# Setup Guide — Native OPC HDA (COM/DCOM)

**No Matrikon Automation required.**  
This client talks directly to the OPC HDA server using custom COM interfaces:

- `IOPCHDA_Server`
- `IOPCHDA_SyncRead`

**Scripts:**
- `opchdareadraw_nativecomdcomwrapper.py` — ReadRaw  
- `opchdareadprocessed_nativecomdcomwrapper.py` — ReadProcessed  

**Config:** `.env`

Use this guide on the **client PC** (the machine that runs Python).

---

## 1. What this approach is

| Item | Detail |
|------|--------|
| Client wrapper | **None** (native COM/DCOM) |
| Not used | `Matrikon.OPCHDA.Automation.1` |
| Server examples | Integration Objects, Matrikon HDA, Kepware HDA, Honeywell PHD HDA, … |
| How connect works | `CoCreateInstance` of server **CLSID** on remote host |
| How read works | `GetItemHandles` → `IOPCHDA_SyncRead.ReadRaw` |

```
Python (32-bit)
    → IPC$ authentication
    → CreateObject(CLSID, machine=host) as IOPCHDA_Server
    → QueryInterface IOPCHDA_SyncRead
    → GetItemHandles(tag)
    → ReadRaw(start, end)
    → print values (UTC)
```

---

## 2. Prerequisites (client PC)

### 2.1 OS

- Windows 10 / 11 (or Windows Server)
- Network reachability to OPC HDA server (ping / firewall allow DCOM/RPC)

### 2.2 Python — **32-bit only**

Most classic OPC servers/proxy stubs are **32-bit**. Use 32-bit Python.

1. Install **Python 3.11 (32-bit)** from python.org (Windows installer **x86**, not x86-64).
2. During install, optionally check **Add python.exe to PATH**.
3. Verify:

```powershell
py -0p
# or
where python
python -c "import struct; print(struct.calcsize('P') * 8)"
```

Expected: **`32`**

### 2.3 Create venv and install packages

From **CMD** (`cmd.exe`), in the project folder:

```bat
cd /d C:\Users\saisu\Downloads\opcresearch38

C:\Users\%USERNAME%\AppData\Local\Programs\Python\Python311-32\python.exe -m venv env32

env32\Scripts\activate.bat

python -m pip install --upgrade pip
python -m pip install pywin32 comtypes
```

Verify:

```bat
python -c "import win32com.client, comtypes; print('pywin32+comtypes OK', comtypes.__version__)"
```

### 2.4 OPC Core Components / HDA proxy-stub (recommended)

Remote `IOPCHDA_*` calls need the OPC HDA **proxy/stub** registered on the client.

1. Install **OPC Core Components Redistributable** (from OPC Foundation / your OPC vendor pack). Prefer the **x86** / 32-bit package when available.
2. After install, reboot if the installer asks.
3. Confirm DCOM is enabled (see §4).

> If `GetHistorianStatus` / `GetItemHandles` work but `ReadRaw` fails oddly, re-check Core Components / proxy-stub install (32-bit).

### 2.5 What you do **NOT** install for this path

- Matrikon OPC HDA Automation  
- `Matrikon.OPCHDA.Automation.1`

---

## 3. Server side (brief — done once on HDA machine)

On the OPC HDA server PC (example: Integration Objects Advanced Simulator):

1. HDA server installed and running (or set to launch on demand).
2. Note **ProgID** and **CLSID**, e.g.:

```powershell
reg query "HKCR\IntegrationObjects.AdvancedSimulator.1\CLSID"
```

Example CLSID used in lab:

```text
{22260278-2009-1112-1982-11121982ABCD}
```

3. DCOM configured (32-bit if server is 32-bit):

```powershell
mmc comexp.msc /32
```

- **My Computer → Properties → Default Properties:** Enable Distributed COM = Yes  
- **COM Security:** allow your domain/local user **Remote Access**, **Remote Launch**, **Remote Activation**  
- App identity: Interactive User or a dedicated service account (as per site policy)

4. Firewall: allow DCOM/RPC between client and server (or temporarily disable for lab test only).

5. Windows account on server that the client will use via IPC$ (same user/password as in `.env`).

---

## 4. Client DCOM / network checklist

1. Client and server on same LAN (or VPN with RPC allowed).
2. Use a real Windows user that exists on the server (or domain account).
3. First connection uses **IPC$**:

```powershell
net use \\<SERVER_IP>\IPC$ /user:<DOMAIN\user> <password>
```

The Python script does this automatically from `.env`.

4. Do **not** set client `OPC_HOST=localhost` unless the HDA server is installed **on this same PC**. Remote CLSID is not registered locally → `Class not registered`.

---

## 5. Project files

Copy at least:

| File | Purpose |
|------|---------|
| `opchdareadraw_nativecomdcomwrapper.py` | Native ReadRaw client |
| `opchdareadprocessed_nativecomdcomwrapper.py` | Native ReadProcessed client |
| `.env` | Host, user, password, CLSID, tag, interval, aggregate |

Optional: keep `env32\` or recreate venv with §2.3.

---

## 6. Configure `.env`

Create / edit `.env` next to the script:

```env
# Native OPC HDA client
OPC_HOST=192.168.1.6
OPC_USER=DOMAIN\User Name
OPC_PASSWORD=your_password_here
OPC_PROGID=IntegrationObjects.AdvancedSimulator.1
OPC_CLSID={22260278-2009-1112-1982-11121982ABCD}
OPC_TAG=Random/Int1
OPC_MINUTES=10
OPC_INTERVAL=10
OPC_AGGREGATE=interpolative
```

| Key | Meaning |
|-----|---------|
| `OPC_HOST` | Server IP or hostname (**not** localhost unless server is local) |
| `OPC_USER` | `DOMAIN\user` or `MACHINE\user` |
| `OPC_PASSWORD` | Password (leave empty to prompt at run) |
| `OPC_PROGID` | Server ProgID (informational / fallback) |
| `OPC_CLSID` | **Required** when ProgID is not registered on the client |
| `OPC_TAG` | HDA item ID |
| `OPC_MINUTES` | Default history window when you press Enter |
| `OPC_INTERVAL` | ReadProcessed resample interval (seconds) |
| `OPC_AGGREGATE` | ReadProcessed aggregate (`interpolative`, `average`, …) |

### How to get CLSID (on the **server**)

```powershell
reg query "HKCR\<YourServerProgID>\CLSID"
```

Paste that `{...}` into `OPC_CLSID`.

---

## 7. How to run (CMD)

Open **Command Prompt** (`cmd.exe`). Replace the folder path with your real project path.

### 7.1 One-time: open folder + activate venv

```bat
cd /d C:\Users\saisu\Downloads\opcresearch38
env32\Scripts\activate.bat
```

You should see `(env32)` at the start of the prompt.

> Without activate, call Python directly:  
> `env32\Scripts\python.exe opchdareadraw_nativecomdcomwrapper.py`

### 7.2 ReadRaw (interactive)

```bat
python opchdareadraw_nativecomdcomwrapper.py
```

Or without activating venv:

```bat
env32\Scripts\python.exe opchdareadraw_nativecomdcomwrapper.py
```

**Prompts:**

```text
Start time [...]:     ← press Enter = last 10 minutes
End time   [...]:     ← press Enter = now
```

Optional time examples: `NOW-10M` | `NOW` | `2026/11/15 21:00:00`

### 7.3 ReadRaw (no prompts)

```bat
python opchdareadraw_nativecomdcomwrapper.py --no-prompt
```

### 7.4 ReadProcessed (interactive)

```bat
python opchdareadprocessed_nativecomdcomwrapper.py
```

**Prompts:**

```text
Start time [...]:                      ← Enter = last 10 minutes
End time   [...]:                      ← Enter = now
Resample interval seconds [10.0]:      ← Enter = 10 seconds
Aggregate [interpolative]:             ← Enter = interpolative
```

### 7.5 ReadProcessed (no prompts)

```bat
python opchdareadprocessed_nativecomdcomwrapper.py --no-prompt --interval 10 --aggregate interpolative
```

Other aggregates (examples): `average` `minimum` `maximum` `total` `count`

### 7.6 Settings come from `.env`

Edit `.env` in the project folder (`OPC_HOST`, `OPC_USER`, `OPC_PASSWORD`, `OPC_CLSID`, `OPC_TAG`, …), then run the commands above. No need to pass host/user on the command line unless you want to override.

### 7.7 PowerShell alternative (optional)

```powershell
cd C:\Users\saisu\Downloads\opcresearch38
.\env32\Scripts\Activate.ps1
python .\opchdareadraw_nativecomdcomwrapper.py
python .\opchdareadprocessed_nativecomdcomwrapper.py
```

---

## 8. Success checklist

You should see roughly:

1. `[INFO] Loaded env from ...\.env`
2. IPC$ OK (or skipped if truly local)
3. `CoCreate via clsid={...} machine=<server-ip>`
4. `IOPCHDA_Server + IOPCHDA_SyncRead acquired`
5. `GetHistorianStatus` vendor string
6. `GetItemHandles server_handle=...`
7. `ReadRaw dwCount=...`
8. Table of values + **TIMESTAMP (UTC)**
9. `[SUCCESS] NATIVE HDA READ COMPLETED`

Compare a few rows with MatrikonOPC HDA Explorer (same tag). Values/ms should match; Explorer often shows **local time**, script shows **UTC** (e.g. IST = UTC+5:30).

---

## 9. Troubleshooting

| Error | Meaning | Fix |
|-------|---------|-----|
| `Invalid class string` (0x800401F3) | ProgID not on client | Set `OPC_CLSID` |
| `Class not registered` (0x80040154) | CoCreate on wrong machine / CLSID missing here | Set `OPC_HOST` to server IP; do not use localhost for remote server |
| `Access is denied` | DCOM / user rights | Fix COM Security + IPC$ user; check server App identity |
| `Server execution failed` (0x80080005) | Server AppID / identity / not running | Fix AppID on server; start simulator; DCOM identity |
| IPC$ failed | Bad password / account | Fix `OPC_USER` / `OPC_PASSWORD` |
| Host shows `localhost` but `.env` has IP | Old shell env var | Script prefers `.env`; or `Remove-Item Env:OPC_HOST` |
| Hard crash on ReadRaw | Bad marshaling / proxy | Reinstall OPC Core Components (x86); use 32-bit Python |

### Quick registry / connectivity checks (client)

```powershell
# Must be 32-bit Python
python -c "import struct; print(struct.calcsize('P')*8)"

# Optional: clear leftover env
Remove-Item Env:OPC_HOST -ErrorAction SilentlyContinue

# Test IPC$
net use \\192.168.1.6\IPC$ /user:"DOMAIN\User" *
```

### Get CLSID again (server)

```powershell
reg query "HKCR\IntegrationObjects.AdvancedSimulator.1\CLSID"
```

---

## 10. Flow summary (for handover)

1. Install **32-bit Python** + venv + `pywin32` + `comtypes`  
2. Install **OPC Core Components** (proxy/stub)  
3. Configure **server DCOM** + account  
4. Fill **`.env`** (`OPC_HOST`, `OPC_CLSID`, user, password, tag)  
5. CMD: `env32\Scripts\activate.bat` then run ReadRaw / ReadProcessed scripts (§7)  
6. Enter/Enter for last 10 minutes → verify vs Explorer  

**No Matrikon HDA Automation on the client for this path.**
