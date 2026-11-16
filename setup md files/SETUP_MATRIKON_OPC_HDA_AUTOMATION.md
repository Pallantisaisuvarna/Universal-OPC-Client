# Setup Guide — Matrikon OPC HDA Automation

This client uses the **Matrikon HDA Automation wrapper** on the PC that runs Python:

- ProgID: `Matrikon.OPCHDA.Automation.1`

The wrapper then connects to **any** classic OPC HDA server by ProgID + host (Integration Objects, Matrikon, Kepware, Honeywell PHD HDA, …).

**Script:** `opc_hda_readraw.py`  
**Related:** `opc_hda_readprocessed.py` (ReadProcessed)

Use this guide on the **client PC**.

---

## 1. What this approach is

| Item | Detail |
|------|--------|
| Client wrapper | **`Matrikon.OPCHDA.Automation.1`** (must be installed locally) |
| Server | Any OPC HDA–compliant server |
| How connect works | `Dispatch(Automation)` → `Connect(ServerProgID, host)` |
| How read works | `OPCHDAItems.AddItem(tag)` → `item.ReadRaw(...)` |

```
Python (32-bit)
    → IPC$ authentication
    → Dispatch("Matrikon.OPCHDA.Automation.1")   ← LOCAL DLL
    → hda.Connect("IntegrationObjects....", host)
    → OPCHDAItems.AddItem(tag)
    → item.ReadRaw(start, end)
    → print values
```

> This is **not** the native `IOPCHDA_*` path.  
> For native COM/DCOM (no Matrikon), see `SETUP_NATIVE_OPC_HDA_COM_DCOM.md`.

---

## 2. Prerequisites (client PC)

### 2.1 OS

- Windows 10 / 11 (or Windows Server)
- Network access to the OPC HDA server

### 2.2 Python — **32-bit only**

Matrikon HDA Automation is typically a **32-bit** COM in-proc server. Match it with 32-bit Python.

1. Install **Python 3.11 (32-bit)** (Windows **x86** installer).
2. Verify:

```powershell
python -c "import struct; print(struct.calcsize('P') * 8)"
```

Expected: **`32`**

### 2.3 Create venv and install pywin32

From **CMD** (`cmd.exe`):

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

`comtypes` is **not** required for this Matrikon Automation path.

---

## 3. Install Matrikon OPC HDA Automation (required)

### 3.1 Install

1. Download / obtain **Matrikon OPC HDA Automation** (or the Matrikon package that registers HDA Automation).
2. Run the installer **on the client PC** (not only on the server).
3. Prefer **32-bit** / “x86” options if the installer offers architecture choice.
4. Run installer as Administrator if registration fails.

### 3.2 Verify registration

```powershell
reg query "HKCR\Matrikon.OPCHDA.Automation.1"
reg query "HKCR\Matrikon.OPCHDA.Automation.1\CLSID"
```

Python check (32-bit venv):

```powershell
.\env32\Scripts\python.exe -c "import win32com.client; o=win32com.client.Dispatch('Matrikon.OPCHDA.Automation.1'); print('OK', o)"
```

If you get **`Invalid class string` (0x800401F3)** → Automation is **not** installed/registered on this PC.

### 3.3 Optional: 32-bit registry view

If 64-bit tools look wrong, query Wow6432Node / use SysWOW64 reg:

```powershell
& "$env:SystemRoot\SysWOW64\reg.exe" query "HKCR\Matrikon.OPCHDA.Automation.1\CLSID"
```

---

## 4. Server side (brief)

On the HDA server PC:

1. OPC HDA server installed (e.g. Integration Objects Advanced Simulator).
2. Know the **server ProgID**, e.g. `IntegrationObjects.AdvancedSimulator.1`.
3. DCOM configured for remote clients (32-bit `mmc comexp.msc /32` if server is 32-bit):
   - Enable Distributed COM
   - Allow your user Remote Access / Launch / Activation
   - App identity set correctly (Interactive User or service account)
4. Firewall allows DCOM/RPC from the client.
5. Account exists for IPC$ login from the client.

> With Matrikon Automation you usually **do not** need the server ProgID registered on the client.  
> Matrikon resolves the remote server during `Connect(progid, host)`.

---

## 5. Client network / IPC$

```powershell
net use \\<SERVER_IP>\IPC$ /user:<DOMAIN\user> <password>
```

The script `opc_hda_readraw.py` performs IPC$ clear + connect before COM.

---

## 6. Project files

| File | Purpose |
|------|---------|
| `opc_hda_readraw.py` | Remote ReadRaw via Matrikon Automation |
| `opc_hda_readprocessed.py` | ReadProcessed (same Automation stack) |

---

## 7. How to run (CMD)

Open **Command Prompt** (`cmd.exe`). Replace host/user/path with your values.

### 7.1 One-time: open folder + activate venv

```bat
cd /d C:\Users\saisu\Downloads\opcresearch38
env32\Scripts\activate.bat
```

You should see `(env32)` at the start of the prompt.

> Without activate:  
> `env32\Scripts\python.exe opc_hda_readraw.py ...`

### 7.2 ReadRaw (interactive)

```bat
python opc_hda_readraw.py --host 192.168.1.6 --progid IntegrationObjects.AdvancedSimulator.1 --user "DOMAIN\User Name" --tag Random/Int1 --minutes 10
```

Password is prompted if you do not pass `--password`.

**Time prompts:** Enter = last N minutes → now (or type `NOW-10M` / `NOW` / `YYYY/MM/DD HH:MM:SS`).

### 7.3 ReadRaw (no prompts)

```bat
python opc_hda_readraw.py --host 192.168.1.6 --progid IntegrationObjects.AdvancedSimulator.1 --user "DOMAIN\User Name" --tag Random/Int1 --minutes 10 --no-prompt
```

### 7.4 ReadProcessed (interactive)

```bat
python opc_hda_readprocessed.py --host 192.168.1.6 --progid IntegrationObjects.AdvancedSimulator.1 --user "DOMAIN\User Name" --tag Random/Int1 --minutes 10 --interval 10 --aggregate interpolative
```

Prompts for time range, then resample interval / aggregate (Enter keeps defaults).

### 7.5 ReadProcessed (no prompts)

```bat
python opc_hda_readprocessed.py --host 192.168.1.6 --progid IntegrationObjects.AdvancedSimulator.1 --user "DOMAIN\User Name" --tag Random/Int1 --minutes 10 --interval 10 --aggregate interpolative --no-prompt
```

### 7.6 Useful arguments

| Arg | Default / meaning |
|-----|-------------------|
| `--host` | Server IP |
| `--progid` | Remote HDA server ProgID |
| `--automation` | `Matrikon.OPCHDA.Automation.1` |
| `--tag` | Item ID |
| `--user` / `--domain` | Windows account for IPC$ |
| `--password` | Optional; else prompted |
| `--minutes` | Default history window |
| `--interval` | ReadProcessed resample seconds |
| `--aggregate` | e.g. `interpolative`, `average` |
| `--num-values` | Cap (0 = server default / all in range) |
| `--bounds` | Include bounding values (ReadRaw) |
| `--no-prompt` | Skip interactive time / option prompts |

### 7.7 PowerShell alternative (optional)

```powershell
cd C:\Users\saisu\Downloads\opcresearch38
.\env32\Scripts\Activate.ps1
python .\opc_hda_readraw.py --host 192.168.1.6 --progid IntegrationObjects.AdvancedSimulator.1 --user "DOMAIN\User Name" --tag Random/Int1 --minutes 10
python .\opc_hda_readprocessed.py --host 192.168.1.6 --progid IntegrationObjects.AdvancedSimulator.1 --user "DOMAIN\User Name" --tag Random/Int1 --minutes 10 --interval 10 --aggregate interpolative
```

---

## 8. Success checklist

1. IPC$ authentication successful  
2. `EnsureDispatch` / `Dispatch` of `Matrikon.OPCHDA.Automation.1` OK  
3. `Connect(progid, host)` OK  
4. `AddItem` OK  
5. `ReadRaw` returns rows  
6. Values match MatrikonOPC HDA Explorer for the same tag/window  

---

## 9. Troubleshooting

| Error | Meaning | Fix |
|-------|---------|-----|
| `Invalid class string` on Dispatch | Matrikon Automation missing on **client** | Install/register `Matrikon.OPCHDA.Automation.1` (32-bit) |
| `Server execution failed` on Connect | Server DCOM / AppID / identity | Fix server DCOM AppID, identity, start server EXE |
| Access denied | DCOM permissions / wrong user | COM Security + correct IPC$ account |
| Connect hangs | DCOM/firewall/AppID incomplete | Server AppID keys, firewall, pre-start simulator |
| Wrong bitness | 64-bit Python vs 32-bit Automation | Use 32-bit Python + `env32` |
| ProgID wrong | Typo / different server | Confirm ProgID on server with `reg query HKCR\...` |

### Verify Automation only

```powershell
reg query "HKCR\Matrikon.OPCHDA.Automation.1\CLSID"
.\env32\Scripts\python.exe -c "import win32com.client; print(win32com.client.Dispatch('Matrikon.OPCHDA.Automation.1'))"
```

### Verify remote server ProgID (on server)

```powershell
reg query "HKCR\IntegrationObjects.AdvancedSimulator.1\CLSID"
```

---

## 10. Flow summary (for handover)

1. Install **32-bit Python** + venv + **pywin32**  
2. Install **Matrikon OPC HDA Automation** on the **client**  
3. Confirm `Dispatch('Matrikon.OPCHDA.Automation.1')` works  
4. Ensure server DCOM + account ready  
5. CMD: `env32\Scripts\activate.bat` then run `opc_hda_readraw.py` / `opc_hda_readprocessed.py` (§7)  
6. Verify against MatrikonOPC HDA Explorer  

**Requires Matrikon HDA Automation on the client. Does not use native `IOPCHDA_*` bindings.**

---

## 11. Quick compare vs native path

| | Matrikon Automation (this doc) | Native COM/DCOM (other doc) |
|--|-------------------------------|-----------------------------|
| Client install | Matrikon HDA Automation | No Matrikon; `comtypes` + OPC proxy/stub |
| Create local object | `Matrikon.OPCHDA.Automation.1` | Server CLSID via `CreateObject` |
| Connect to server | `Connect(ProgID, host)` | CoCreate remote CLSID |
| Script | `opc_hda_readraw.py` | `opchdareadraw_nativecomdcomwrapper.py` |
| Need server CLSID on client? | Usually no | Yes (if ProgID not registered locally) |
