# Setup Guide — Native OPC DA (COM/DCOM)

**No `OPC.Automation.1` required.**  
This client talks directly to the OPC DA server using custom COM interfaces:

- `IOPCServer`
- `IOPCItemMgt`
- `IOPCSyncIO`

**Script:** `opcda_nativecomanddcom.py`  
**Config:** `.env` (same style as native HDA)

> Automation DA path (wrapper): see `SETUP_OPC_DA_AUTOMATION.md` / `opc_da_remote_read_test.py`.  
> This doc does **not** change that Automation script.

---

## 1. What this approach is

| Item | Detail |
|------|--------|
| Client wrapper | **None** (native COM/DCOM) |
| Not used | `OPC.Automation.1` |
| How connect works | `CoCreateInstance` of server **CLSID** as `IOPCServer` |
| How read works | `AddGroup` → `AddItems` → `IOPCSyncIO.Read` |
| Tags | Pass with `--tag` (or `OPC_TAG` in `.env`). No Automation browser. |

```
Python (32-bit)
    → IPC$ authentication
    → CreateObject(CLSID, machine=host) as IOPCServer
    → AddGroup → IOPCItemMgt + IOPCSyncIO
    → AddItems(tag…)
    → Read (device/cache) in a loop
```

---

## 2. Prerequisites (client PC)

Same base as native HDA:

1. **32-bit Python** + `env32`  
2. `pywin32` + `comtypes`  
3. **OPC Core Components** (DA proxy/stub) recommended  
4. Server DCOM + IPC$ account configured  

```bat
cd /d C:\Users\saisu\Downloads\opcresearch38
env32\Scripts\activate.bat
python -m pip install pywin32 comtypes
python -c "import struct,comtypes; print(struct.calcsize('P')*8, comtypes.__version__)"
```

Expected bitness: **`32`**

### What you do **NOT** need for this path

- `OPC.Automation.1` (optional to have; unused by this script)

---

## 3. `.env` (shared with native HDA)

```env
OPC_HOST=192.168.1.6
OPC_USER=Swapnildebtasus\Swapnil kumar debta
OPC_PASSWORD=
OPC_PROGID=IntegrationObjects.AdvancedSimulator.1
OPC_CLSID={22260278-2009-1112-1982-11121982ABCD}
OPC_TAG=Random/Int1
OPC_DA_INTERVAL=1
```

| Key | Meaning |
|-----|---------|
| `OPC_HOST` | Server IP (**not** localhost unless DA server is local) |
| `OPC_CLSID` | Required when ProgID is not registered on the client |
| `OPC_TAG` | Default ItemID (comma-separated allowed) |
| `OPC_DA_INTERVAL` | Seconds between read cycles |

For Integration Objects, the **same CLSID** often serves both DA and HDA (different interfaces).

---

## 4. How to run (CMD)

Open **Command Prompt** (`cmd.exe`).

### 4.1 Activate venv

```bat
cd /d C:\Users\saisu\Downloads\opcresearch38
env32\Scripts\activate.bat
```

### 4.2 Basic run (uses `.env` tag / host)

```bat
python opcda_nativecomanddcom.py
```

Or:

```bat
env32\Scripts\python.exe opcda_nativecomanddcom.py
```

Press **Ctrl+C** to stop the continuous loop.

### 4.3 Explicit tag + interval

```bat
python opcda_nativecomanddcom.py --tag Random/Int1 --interval 1
```

### 4.4 Multiple tags

```bat
python opcda_nativecomanddcom.py --tag Random/Int1 --tag Random/Int4 --interval 1
```

Or comma-separated:

```bat
python opcda_nativecomanddcom.py --tag "Random/Int1,Random/Int4" --interval 1
```

### 4.5 Single read then exit

```bat
python opcda_nativecomanddcom.py --tag Random/Int1 --once
```

### 4.6 Full override example

```bat
python opcda_nativecomanddcom.py --host 192.168.1.6 --clsid "{22260278-2009-1112-1982-11121982ABCD}" --user "Swapnildebtasus\Swapnil kumar debta" --tag Random/Int1 --interval 1 --source device
```

### 4.7 Useful arguments

| Arg | Meaning |
|-----|---------|
| `--host` | Server IP |
| `--progid` | Server ProgID (fallback) |
| `--clsid` | Server CLSID for native CoCreate |
| `--user` | `DOMAIN\user` |
| `--password` | Optional; else `.env` / prompt |
| `--tag` | ItemID (repeatable) |
| `--interval` | Seconds between cycles |
| `--once` | One read, then exit |
| `--source` | `device` (default) or `cache` |
| `--env` | Path to `.env` |

---

## 5. Success checklist

1. IPC$ OK (or skipped on localhost)  
2. `IOPCServer acquired`  
3. `AddGroup` + `IOPCItemMgt + IOPCSyncIO`  
4. `[ADDED] <tag>`  
5. Live table: TAG / VALUE / QUALITY / TIMESTAMP  
6. Values match Automation DA / Explorer for the same tags  

---

## 6. Troubleshooting

| Error | Fix |
|-------|-----|
| Invalid class string / Class not registered | Set `OPC_CLSID`; use server IP not localhost |
| AddItems skipped / HRESULT | Wrong ItemID; confirm tag path (e.g. `Random.Int1` for Matrikon) |
| **Access is denied** on CoCreate (`0x80070005`) | Server DCOM Launch/Activation (see below) — **not** a Python bug |
| 64-bit Python | Use 32-bit `env32` |

### Access Denied on native CoCreate (Matrikon Simulation)

IPC$ can succeed while CoCreate still fails. Native COM activates the server EXE remotely; Automation `Connect` uses a different path.

**Client-side (already in script):** after IPC$, the script **impersonates** with `LogonUser(NEW_CREDENTIALS)` so DCOM uses `desktop-e9295iq\dell` (not your local interactive token). Re-run after pulling the latest `opcda_nativecomanddcom.py`.

If it **still** fails:

**On the OPC server PC (`192.168.1.141`), as Administrator:**

1. Open 32-bit Component Services:

```bat
mmc comexp.msc /32
```

2. **Console Root → Component Services → Computers → My Computer → Properties**
   - **Default Properties:** Enable Distributed COM = **Yes**
   - **COM Security → Access Permissions → Edit Limits:** add `desktop-e9295iq\dell` → allow **Remote Access**
   - **COM Security → Launch and Activation Permissions → Edit Limits:** same user → allow **Remote Launch** + **Remote Activation**
   - Optionally set **Edit Defaults** the same way

3. **My Computer → DCOM Config** → find **`Matrikon.OPC.Simulation.1`**  
   (or search by CLSID `{F8582CF2-88FB-11D0-B850-00C0F0104305}`)
   - **Properties → Identity:** **The interactive user** (and stay logged on with Simulation running), **or** “This user” = `dell`
   - **Properties → Security:**
     - **Launch and Activation Permissions** → Customize → add your user → **Local + Remote** Launch/Activation
     - **Access Permissions** → Customize → add your user → **Local + Remote** Access

4. Windows Firewall: allow the Matrikon OPC Simulation EXE / DCOM, or disable briefly for a lab test.

5. Confirm ProgID spelling on server (note the **a** in Simulation):

```bat
reg query "HKCR\Matrikon.OPC.Simulation.1\CLSID"
```

6. Retry from client:

```bat
python opcda_nativecomanddcom.py --tag Random.Int1 --interval 1
```

**Why Automation DA may work but native fails:**  
`OPC.Automation.1` → `Connect(progid, host)` often piggybacks a different activation path. Native `CoCreateInstanceEx(CLSID)` always needs **Remote Launch/Activation** on that AppID.

---

## 7. Compare vs Automation DA

| | Native (this doc) | Automation (`SETUP_OPC_DA_AUTOMATION.md`) |
|--|-------------------|-------------------------------------------|
| Wrapper | None | `OPC.Automation.1` |
| Script | `opcda_nativecomanddcom.py` | `opc_da_remote_read_test.py` |
| Browse all tags | No (pass `--tag`) | Yes (`CreateBrowser`) |
| Connect | CoCreate CLSID → `IOPCServer` | `Dispatch` → `Connect(progid, host)` |
