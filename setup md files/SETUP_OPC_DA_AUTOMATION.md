# Setup Guide — OPC DA (`OPC.Automation.1`)

This client uses the standard **OPC Data Access Automation** wrapper on the PC that runs Python:

- ProgID: **`OPC.Automation.1`**

That wrapper connects to a remote (or local) **OPC DA** server by ProgID + host  
(e.g. Integration Objects, Matrikon Simulation, Prosys, Kepware, …).

**Script:** `opc_da_remote_read_test.py`

> This is **OPC DA** (live/current values), not OPC HDA (history).  
> For HDA see `SETUP_NATIVE_OPC_HDA_COM_DCOM.md` or `SETUP_MATRIKON_OPC_HDA_AUTOMATION.md`.

---

## 1. What this approach is

| Item | Detail |
|------|--------|
| Client wrapper | **`OPC.Automation.1`** (must be registered on the client) |
| Typical DLL | Often `OPCDAAuto.dll` (OPC Foundation / vendor redistributable) |
| Server | Any classic OPC DA server |
| How connect works | `Dispatch("OPC.Automation.1")` → `Connect(ServerProgID, host)` |
| What the script does | IPC$ → Connect → browse tags → create group → `AddItem` → read Value / Quality / Timestamp in a loop |

```
Python (32-bit)
    → IPC$ authentication
    → Dispatch("OPC.Automation.1")          ← LOCAL wrapper
    → Connect("IntegrationObjects....", host)
    → Browse items / AddItem(tag)
    → Read current values (polling loop)
```

`AddItem` only registers the tag in the **client group** — it does **not** create tags on the server.

---

## 2. Prerequisites (client PC)

### 2.1 OS

- Windows 10 / 11 (or Windows Server)
- Network reachability to the OPC DA server

### 2.2 Python — **32-bit only**

`OPC.Automation.1` is almost always a **32-bit** COM object. Use 32-bit Python.

1. Install **Python 3.11 (32-bit)** (Windows installer **x86**, not x86-64).
2. Verify in **CMD**:

```bat
python -c "import struct; print(struct.calcsize('P') * 8)"
```

Expected: **`32`**

### 2.3 Create venv and install pywin32 (CMD)

```bat
cd /d C:\Users\saisu\Downloads\opcresearch38

C:\Users\%USERNAME%\AppData\Local\Programs\Python\Python311-32\python.exe -m venv env32

env32\Scripts\activate.bat

python -m pip install --upgrade pip
python -m pip install pywin32
```

Verify:

```bat
python -c "import win32com.client; print('pywin32 OK')"
```

`comtypes` is **not** required for this DA Automation path.

---

## 3. Install / register `OPC.Automation.1` (required on client)

### 3.1 Where it comes from

Install one of these on the **client** (the machine running Python):

- **OPC Core Components Redistributable** (OPC Foundation) — often registers `OPC.Automation.1` / `OPCDAAuto.dll`
- Or your vendor pack that includes **OPC DA Automation**

Prefer **x86 / 32-bit** when the installer offers a choice. Run as Administrator if registration fails.

### 3.2 Verify registration (CMD)

```bat
reg query "HKCR\OPC.Automation.1"
reg query "HKCR\OPC.Automation.1\CLSID"
```

Python check (32-bit venv):

```bat
cd /d C:\Users\saisu\Downloads\opcresearch38
env32\Scripts\python.exe -c "import win32com.client; o=win32com.client.Dispatch('OPC.Automation.1'); print('OK', o)"
```

If you get **`Invalid class string` (0x800401F3)** → `OPC.Automation.1` is **not** installed/registered on this PC.

### 3.3 Optional: confirm DLL (32-bit)

```bat
reg query "HKCR\OPC.Automation.1\CLSID"
rem Then query InprocServer32 for that CLSID, e.g.:
%SystemRoot%\SysWOW64\reg.exe query "HKCR\CLSID\{28E68F9A-8D75-11D1-8DC3-3C302A000000}\InprocServer32"
```

(Your CLSID may differ; use the one from the ProgID query.)

### 3.4 What you do **not** need for this DA path

- Matrikon HDA Automation (`Matrikon.OPCHDA.Automation.1`) — that is for **HDA**, not DA

---

## 4. Server side (brief)

On the OPC DA server PC:

1. OPC DA server installed and running (e.g. Integration Objects Advanced Simulator).
2. Know the **server ProgID**, e.g. `IntegrationObjects.AdvancedSimulator.1`.
3. DCOM configured for remote clients. If the server is 32-bit:

```bat
mmc comexp.msc /32
```

- **My Computer → Properties → Default Properties:** Enable Distributed COM = Yes  
- **COM Security:** allow your user **Remote Access**, **Remote Launch**, **Remote Activation**  
- App identity: Interactive User or a dedicated account (site policy)

4. Firewall allows DCOM/RPC from the client.
5. Windows account exists for IPC$ login from the client.

> With `OPC.Automation.1`, you usually **do not** need the server ProgID registered on the client.  
> The Automation wrapper resolves the remote server during `Connect(progid, host)`.

---

## 5. Client network / IPC$

```bat
net use \\192.168.1.6\IPC$ /user:DOMAIN\User *
```

The script clears and recreates IPC$ automatically before COM (skipped for `localhost`).

---

## 6. Project file

| File | Purpose |
|------|---------|
| `opc_da_remote_read_test.py` | Remote OPC DA browse + continuous read via `OPC.Automation.1` |

---

## 7. How to run (CMD)

Open **Command Prompt** (`cmd.exe`).

### 7.1 Open folder + activate venv

```bat
cd /d C:\Users\saisu\Downloads\opcresearch38
env32\Scripts\activate.bat
```

You should see `(env32)` in the prompt.

### 7.2 Basic run (lab defaults)

Uses defaults in the script (`192.168.1.6`, Integration Objects, etc.):

```bat
python opc_da_remote_read_test.py
```

Or without activate:

```bat
env32\Scripts\python.exe opc_da_remote_read_test.py
```

Enter password when prompted.

### 7.3 Full command (recommended for handover)

```bat
python opc_da_remote_read_test.py --host 192.168.1.6 --progid IntegrationObjects.AdvancedSimulator.1 --domain Swapnildebtasus --user "Swapnil kumar debta" --interval 1
```

With password on the command line (optional; prefer prompt in shared machines):

```bat
python opc_da_remote_read_test.py --host 192.168.1.6 --progid IntegrationObjects.AdvancedSimulator.1 --domain Swapnildebtasus --user "Swapnil kumar debta" --password YOUR_PASSWORD --interval 1
```

### 7.4 Filter tags (optional)

Only tags whose name contains a substring:

```bat
python opc_da_remote_read_test.py --host 192.168.1.6 --progid IntegrationObjects.AdvancedSimulator.1 --domain Swapnildebtasus --user "Swapnil kumar debta" --filter Random
```

Leave `--filter` empty (default) to browse/add **all** discovered items.

### 7.5 Localhost test (server on same PC)

```bat
python opc_da_remote_read_test.py --host localhost --progid IntegrationObjects.AdvancedSimulator.1
```

IPC$ is skipped for localhost.

### 7.6 Useful arguments

| Arg | Meaning |
|-----|---------|
| `--host` | Server IP or hostname |
| `--progid` | Remote OPC DA server ProgID |
| `--domain` | Windows domain / machine name |
| `--user` | Windows user (combined as `DOMAIN\user`) |
| `--password` | Optional; else prompted |
| `--interval` | Seconds between read loops (default `1`) |
| `--filter` | Optional tag name substring |

Stop the continuous loop with **Ctrl+C**.

---

## 8. Success checklist

1. `Dispatch('OPC.Automation.1')` works (§3.2)  
2. IPC$ OK (or skipped on localhost)  
3. `Connect(progid, host)` OK  
4. Browse finds tags (or filter matches)  
5. Group items added; values / quality / timestamps print in the loop  
6. Values match what you see in an OPC DA client (e.g. Matrikon Explorer DA view / Prosys) for the same tags  

---

## 9. Troubleshooting

| Error | Meaning | Fix |
|-------|---------|-----|
| `Invalid class string` on Dispatch | `OPC.Automation.1` missing on **client** | Install OPC Core Components / DA Automation (32-bit); re-run §3.2 |
| Wrong bitness | 64-bit Python vs 32-bit Automation | Use 32-bit Python + `env32` |
| Access denied / Connect fail | DCOM / user rights | Fix COM Security + IPC$ account; check server identity |
| Server execution failed | Server AppID / not running | Start DA server; fix DCOM App identity |
| IPC$ failed | Bad user/password | Fix `--domain` / `--user` / password |
| No tags found | Browse empty / filter too strict | Clear `--filter`; confirm server has items |
| Localhost Class not registered | Server not on this PC | Use real server IP in `--host` |

### Quick verify (CMD)

```bat
reg query "HKCR\OPC.Automation.1\CLSID"

cd /d C:\Users\saisu\Downloads\opcresearch38
env32\Scripts\python.exe -c "import struct,win32com.client; print(struct.calcsize('P')*8); print(win32com.client.Dispatch('OPC.Automation.1'))"
```

---

## 10. Flow summary (for handover)

1. Install **32-bit Python** + `env32` + **pywin32**  
2. Install / register **`OPC.Automation.1`** on the **client**  
3. Confirm `Dispatch('OPC.Automation.1')` prints OK  
4. Configure server DCOM + account  
5. CMD:

```bat
cd /d C:\Users\saisu\Downloads\opcresearch38
env32\Scripts\activate.bat
python opc_da_remote_read_test.py --host 192.168.1.6 --progid IntegrationObjects.AdvancedSimulator.1 --domain Swapnildebtasus --user "Swapnil kumar debta" --interval 1
```

6. Verify live values against an OPC DA explorer  

**Requires `OPC.Automation.1` on the client. Does not use Matrikon HDA Automation.**

---

## 11. Quick compare (DA vs HDA docs)

| | This doc (OPC DA) | HDA Matrikon doc | HDA Native doc |
|--|-------------------|------------------|----------------|
| Wrapper | `OPC.Automation.1` | `Matrikon.OPCHDA.Automation.1` | None (`IOPCHDA_*`) |
| Data | Current / live | History | History |
| Script | `opc_da_remote_read_test.py` | `opc_hda_readraw.py` | `opchdareadraw_nativecomdcomwrapper.py` |
