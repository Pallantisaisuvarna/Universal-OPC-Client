"""
Native OPC DA reader (CLI) — COM/DCOM custom interfaces
=======================================================

Uses IOPCServer + IOPCItemMgt + IOPCSyncIO directly
(no OPC.Automation.1 wrapper).

Requires: 32-bit Python + pywin32 + comtypes

Config: .env (OPC_HOST, OPC_CLSID, OPC_USER, OPC_TAG, …) then:
  .\\env32\\Scripts\\python.exe opcda_nativecomanddcom.py
  .\\env32\\Scripts\\python.exe opcda_nativecomanddcom.py --tag Random/Int1 --interval 1
"""

from __future__ import annotations

import argparse
import getpass
import os
import re
import subprocess
import sys
import time
from ctypes import (
    POINTER,
    Structure,
    byref,
    c_byte,
    c_float,
    c_long,
    c_uint,
    c_ushort,
    c_void_p,
    c_wchar_p,
    pointer,
    windll,
)
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from ctypes.wintypes import BOOL, DWORD, FILETIME, WORD

DEFAULT_HOST = "192.168.1.3"
DEFAULT_SERVER_PROGID = "IntegrationObjects.AdvancedSimulator.1"
DEFAULT_SERVER_CLSID = "{22260278-2009-1112-1982-11121982ABCD}"
DEFAULT_TAG = "Random/Int1"
DEFAULT_USER = "Swapnil kumar debta"
DEFAULT_INTERVAL = 1.0
DEFAULT_ENV_FILE = ".env"

# Console mode flags (disable Quick Edit so click-select does not freeze the process)
_ENABLE_QUICK_EDIT = 0x0040
_ENABLE_EXTENDED_FLAGS = 0x0080
_ENABLE_INSERT_MODE = 0x0020


def disable_console_quick_edit() -> None:
    """Prevent Windows console 'Quick Edit' from pausing the OPC loop on click."""
    try:
        kernel32 = windll.kernel32
        hin = kernel32.GetStdHandle(-10)  # STD_INPUT_HANDLE
        if hin in (0, -1, None):
            return
        mode = c_uint()
        if not kernel32.GetConsoleMode(hin, byref(mode)):
            return
        new_mode = mode.value
        new_mode |= _ENABLE_EXTENDED_FLAGS
        new_mode &= ~_ENABLE_QUICK_EDIT
        # Keep insert mode off with quick-edit clear (stable for polling CLIs)
        new_mode &= ~_ENABLE_INSERT_MODE
        if kernel32.SetConsoleMode(hin, new_mode):
            print("[INFO] Console Quick Edit disabled (avoids click-to-pause)")
    except Exception as exc:
        print(f"[WARNING] Could not disable Quick Edit: {exc}")


def _flush_print(*args: Any, **kwargs: Any) -> None:
    kwargs.setdefault("flush", True)
    print(*args, **kwargs)


# OPC DA 2.0 custom interface IIDs
IID_IOPCServer = "{39C13A4D-011E-11D0-9675-0020AFD8ADB3}"
IID_IOPCItemMgt = "{39C13A54-011E-11D0-9675-0020AFD8ADB3}"
IID_IOPCSyncIO = "{39C13A52-011E-11D0-9675-0020AFD8ADB3}"

OPC_DS_CACHE = 1
OPC_DS_DEVICE = 2
VT_EMPTY = 0

_GUID_RE = re.compile(
    r"^\{?[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-"
    r"[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}?$"
)
_FILETIME_EPOCH = datetime(1601, 1, 1, tzinfo=timezone.utc)

_DA_QUALITY_STATUS = {
    0x00: "Bad, Non-specific",
    0x04: "Bad, Configuration Error",
    0x08: "Bad, Not Connected",
    0x0C: "Bad, Device Failure",
    0x10: "Bad, Sensor Failure",
    0x14: "Bad, Last Known Value",
    0x18: "Bad, Comm Failure",
    0x1C: "Bad, Out of Service",
    0x40: "Uncertain, Non-specific",
    0x44: "Uncertain, Last Usable Value",
    0x50: "Uncertain, Sensor Not Accurate",
    0x54: "Uncertain, EU Exceeded",
    0x58: "Uncertain, Sub-Normal",
    0xC0: "Good, Non-specific",
    0xD8: "Good, Local Override",
}


def _normalize_clsid(value: str) -> str:
    s = (value or "").strip()
    if not s:
        return ""
    if not s.startswith("{"):
        s = "{" + s
    if not s.endswith("}"):
        s = s + "}"
    return s.upper()


def _looks_like_clsid(value: str) -> bool:
    return bool(_GUID_RE.match((value or "").strip()))


def _resolve_create_id(progid: str, clsid: Optional[str]) -> Tuple[str, str]:
    if clsid and str(clsid).strip():
        return _normalize_clsid(str(clsid)), "clsid"
    p = (progid or "").strip()
    if _looks_like_clsid(p):
        return _normalize_clsid(p), "clsid"
    if p.lower() == DEFAULT_SERVER_PROGID.lower() and DEFAULT_SERVER_CLSID:
        return _normalize_clsid(DEFAULT_SERVER_CLSID), "clsid"
    return p, "progid"


def _filetime_to_dt(ft: FILETIME) -> datetime:
    t = int(ft.dwLowDateTime) | (int(ft.dwHighDateTime) << 32)
    if t <= 0:
        return _FILETIME_EPOCH
    return _FILETIME_EPOCH + timedelta(microseconds=t // 10)


def load_dotenv(path: str | Path | None = None, *, override: bool = True) -> Dict[str, str]:
    env_path = Path(path or DEFAULT_ENV_FILE)
    if not env_path.is_file():
        alt = Path(__file__).resolve().parent / DEFAULT_ENV_FILE
        if alt.is_file():
            env_path = alt
        else:
            return {}
    loaded: Dict[str, str] = {}
    for raw in env_path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, val = line.partition("=")
        key = key.strip()
        val = val.strip()
        if (val.startswith('"') and val.endswith('"')) or (
            val.startswith("'") and val.endswith("'")
        ):
            val = val[1:-1]
        loaded[key] = val
        if override or key not in os.environ:
            os.environ[key] = val
    print(f"[INFO] Loaded env from {env_path.resolve()}")
    return loaded


def env_or(key: str, default: str = "") -> str:
    return (os.environ.get(key) or default).strip()


def authenticate_ipc(host: str, username: str, password: str) -> bool:
    if host.lower() in ("", "localhost", "127.0.0.1"):
        print("[INFO] Local host — skipping IPC$")
        return True
    ipc_path = rf"\\{host}\IPC$"
    print("\n[1] Clearing existing IPC$ connection...")
    subprocess.run(
        ["net", "use", ipc_path, "/delete", "/y"],
        capture_output=True,
        text=True,
    )
    print("[2] Authenticating IPC$...")
    result = subprocess.run(
        ["net", "use", ipc_path, f"/user:{username}", password],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        print("[FAIL] IPC$ authentication failed")
        print(result.stderr or result.stdout)
        return False
    print("[OK] IPC$ authentication successful")
    return True


def decode_quality(q: int) -> str:
    try:
        qi = int(q)
    except Exception:
        return str(q)
    da = qi & 0xFF
    status_key = da & 0xFC
    return _DA_QUALITY_STATUS.get(
        status_key,
        _DA_QUALITY_STATUS.get(da & 0xC0, f"Quality 0x{da:02X}"),
    )


def format_timestamp_utc(ts: datetime) -> str:
    """OPC FILETIME is UTC — display as UTC (match HDA native scripts)."""
    if ts is None:
        return "N/A"
    if getattr(ts, "year", 2000) < 1970:
        return str(ts)
    if ts.tzinfo is None:
        ts = ts.replace(tzinfo=timezone.utc)
    else:
        ts = ts.astimezone(timezone.utc)
    ms = ts.microsecond // 1000
    return ts.strftime(f"%Y-%m-%d %H:%M:%S.{ms:03d} UTC")


def _build_interfaces():
    from comtypes import COMMETHOD, GUID, HRESULT, IUnknown
    from comtypes.automation import VARIANT

    class OPCITEMDEF(Structure):
        _pack_ = 8
        _fields_ = [
            ("szAccessPath", c_wchar_p),
            ("szItemID", c_wchar_p),
            ("bActive", BOOL),
            ("hClient", DWORD),
            ("dwBlobSize", DWORD),
            ("pBlob", POINTER(c_byte)),
            ("vtRequestedDataType", c_ushort),
            ("wReserved", WORD),
        ]

    class OPCITEMRESULT(Structure):
        _pack_ = 8
        _fields_ = [
            ("hServer", DWORD),
            ("vtCanonicalDataType", c_ushort),
            ("wReserved", WORD),
            ("dwAccessRights", DWORD),
            ("dwBlobSize", DWORD),
            ("pBlob", POINTER(c_byte)),
        ]

    class OPCITEMSTATE(Structure):
        _pack_ = 8
        _fields_ = [
            ("hClient", DWORD),
            ("ftTimeStamp", FILETIME),
            ("wQuality", WORD),
            ("wReserved", WORD),
            ("vDataValue", VARIANT),
        ]

    class IOPCServer(IUnknown):
        _iid_ = GUID(IID_IOPCServer)
        _methods_ = [
            # IDL: [out, iid_is(riid)] LPUNKNOWN *ppUnk  → IUnknown**
            COMMETHOD(
                [],
                HRESULT,
                "AddGroup",
                (["in"], c_wchar_p, "szName"),
                (["in"], BOOL, "bActive"),
                (["in"], DWORD, "dwRequestedUpdateRate"),
                (["in"], DWORD, "hClientGroup"),
                (["in"], POINTER(c_long), "pTimeBias"),
                (["in"], POINTER(c_float), "pPercentDeadband"),
                (["in"], DWORD, "dwLCID"),
                (["out"], POINTER(DWORD), "phServerGroup"),
                (["out"], POINTER(DWORD), "pRevisedUpdateRate"),
                (["in"], POINTER(GUID), "riid"),
                (["out"], POINTER(POINTER(IUnknown)), "ppUnk"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "GetErrorString",
                (["in"], HRESULT, "dwError"),
                (["in"], DWORD, "dwLocale"),
                (["out"], POINTER(c_wchar_p), "ppString"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "GetGroupByName",
                (["in"], c_wchar_p, "szName"),
                (["in"], POINTER(GUID), "riid"),
                (["out"], POINTER(POINTER(IUnknown)), "ppUnk"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "GetStatus",
                (["out"], POINTER(c_void_p), "ppServerStatus"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "RemoveGroup",
                (["in"], DWORD, "hServerGroup"),
                (["in"], BOOL, "bForce"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "CreateGroupEnumerator",
                (["in"], DWORD, "dwScope"),
                (["in"], POINTER(GUID), "riid"),
                (["out"], POINTER(POINTER(IUnknown)), "ppUnk"),
            ),
        ]

    class IOPCItemMgt(IUnknown):
        _iid_ = GUID(IID_IOPCItemMgt)
        _methods_ = [
            COMMETHOD(
                [],
                HRESULT,
                "AddItems",
                (["in"], DWORD, "dwCount"),
                (["in"], POINTER(OPCITEMDEF), "pItemArray"),
                (["out"], POINTER(POINTER(OPCITEMRESULT)), "ppAddResults"),
                (["out"], POINTER(POINTER(HRESULT)), "ppErrors"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "ValidateItems",
                (["in"], DWORD, "dwCount"),
                (["in"], POINTER(OPCITEMDEF), "pItemArray"),
                (["in"], BOOL, "bBlobUpdate"),
                (["out"], POINTER(POINTER(OPCITEMRESULT)), "ppValidationResults"),
                (["out"], POINTER(POINTER(HRESULT)), "ppErrors"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "RemoveItems",
                (["in"], DWORD, "dwCount"),
                (["in"], POINTER(DWORD), "phServer"),
                (["out"], POINTER(POINTER(HRESULT)), "ppErrors"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "SetActiveState",
                (["in"], DWORD, "dwCount"),
                (["in"], POINTER(DWORD), "phServer"),
                (["in"], BOOL, "bActive"),
                (["out"], POINTER(POINTER(HRESULT)), "ppErrors"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "SetClientHandles",
                (["in"], DWORD, "dwCount"),
                (["in"], POINTER(DWORD), "phServer"),
                (["in"], POINTER(DWORD), "phClient"),
                (["out"], POINTER(POINTER(HRESULT)), "ppErrors"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "SetDatatypes",
                (["in"], DWORD, "dwCount"),
                (["in"], POINTER(DWORD), "phServer"),
                (["in"], POINTER(c_ushort), "pRequestedDatatypes"),
                (["out"], POINTER(POINTER(HRESULT)), "ppErrors"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "CreateEnumerator",
                (["in"], POINTER(GUID), "riid"),
                (["out"], POINTER(POINTER(IUnknown)), "ppUnk"),
            ),
        ]

    class IOPCSyncIO(IUnknown):
        _iid_ = GUID(IID_IOPCSyncIO)
        _methods_ = [
            COMMETHOD(
                [],
                HRESULT,
                "Read",
                (["in"], DWORD, "dwSource"),
                (["in"], DWORD, "dwCount"),
                (["in"], POINTER(DWORD), "phServer"),
                (["out"], POINTER(POINTER(OPCITEMSTATE)), "ppItemValues"),
                (["out"], POINTER(POINTER(HRESULT)), "ppErrors"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "Write",
                (["in"], DWORD, "dwCount"),
                (["in"], POINTER(DWORD), "phServer"),
                (["in"], POINTER(VARIANT), "pItemValues"),
                (["out"], POINTER(POINTER(HRESULT)), "ppErrors"),
            ),
        ]

    return IOPCServer, IOPCItemMgt, IOPCSyncIO, OPCITEMDEF, OPCITEMRESULT, OPCITEMSTATE


def _split_domain_user(username: str) -> Tuple[str, str]:
    u = (username or "").strip()
    if "\\" in u:
        domain, user = u.split("\\", 1)
        return domain.strip(), user.strip()
    return ".", u


class NativeOpcDaClient:
    """Minimal native OPC DA client (custom COM interfaces over DCOM)."""

    def __init__(self) -> None:
        self._server = None
        self._item_mgt = None
        self._sync = None
        self._com_ready = False
        self._server_group = 0
        self._item_ids: List[str] = []
        self._server_handles: List[int] = []
        self._OPCITEMDEF = None
        self._OPCITEMSTATE = None
        self._impersonation_token = None
        self.prog_id = ""
        self.clsid = ""
        self.host = ""

    @staticmethod
    def _hresult_at(pp_errors: Any, index: int = 0) -> int:
        if pp_errors is None:
            return 0
        try:
            return int(pp_errors[index])
        except Exception:
            try:
                return int(pp_errors[index].value)
            except Exception:
                return 0

    def _begin_impersonation(self, username: str, password: str) -> None:
        """Use network credentials for DCOM (IPC$ alone is often not enough)."""
        import win32con
        import win32security

        domain, user = _split_domain_user(username)
        print(f"[INFO] Impersonating for DCOM as {domain}\\{user} (NEW_CREDENTIALS)")
        token = win32security.LogonUser(
            user,
            domain,
            password,
            win32con.LOGON32_LOGON_NEW_CREDENTIALS,
            win32con.LOGON32_PROVIDER_WINNT50,
        )
        win32security.ImpersonateLoggedOnUser(token)
        self._impersonation_token = token
        print("[OK] Impersonation active for remote CoCreate")

    def _end_impersonation(self) -> None:
        if self._impersonation_token is None:
            return
        try:
            import win32api
            import win32security

            win32security.RevertToSelf()
            win32api.CloseHandle(self._impersonation_token)
        except Exception:
            pass
        self._impersonation_token = None

    def connect(
        self,
        prog_id: str,
        host: str,
        clsid: Optional[str] = None,
        username: str = "",
        password: str = "",
    ) -> None:
        import pythoncom
        from comtypes import CLSCTX_LOCAL_SERVER, CLSCTX_REMOTE_SERVER
        from comtypes.client import CreateObject

        (
            IOPCServer,
            IOPCItemMgt,
            IOPCSyncIO,
            self._OPCITEMDEF,
            _OPCITEMRESULT,
            self._OPCITEMSTATE,
        ) = _build_interfaces()
        self._IOPCItemMgt = IOPCItemMgt
        self._IOPCSyncIO = IOPCSyncIO

        # Same thread for CoCreate + SyncIO.Read (COM proxies are apartment-bound).
        pythoncom.CoInitialize()
        self._com_ready = True
        try:
            pythoncom.CoInitializeSecurity(
                None,
                None,
                None,
                pythoncom.RPC_C_AUTHN_LEVEL_PKT_INTEGRITY,
                pythoncom.RPC_C_IMP_LEVEL_IMPERSONATE,
                None,
                pythoncom.EOAC_NONE,
                None,
            )
        except pythoncom.com_error as exc:
            if not (exc.args and exc.args[0] == -2147417831):
                raise

        create_id, kind = _resolve_create_id(prog_id, clsid)
        self.prog_id = prog_id
        self.clsid = create_id if kind == "clsid" else ""
        self.host = (host or "").strip()
        machine = None if self.host.lower() in ("", "localhost", "127.0.0.1") else self.host
        clsctx = CLSCTX_LOCAL_SERVER if machine is None else CLSCTX_REMOTE_SERVER

        # Remote DCOM: impersonate after CoInitializeSecurity, before CoCreate.
        if machine and username and password:
            self._begin_impersonation(username, password)

        print(
            f"[INFO] CoCreate via {kind}={create_id} "
            f"machine={machine or 'local'} clsctx={clsctx}"
        )
        try:
            self._server = CreateObject(
                create_id,
                machine=machine,
                interface=IOPCServer,
                clsctx=clsctx,
            )
        except OSError as exc:
            winerr = getattr(exc, "winerror", None)
            if kind == "progid" and winerr == -2147221005:
                raise OSError(
                    f"{exc}\n"
                    f"ProgID {prog_id!r} is not registered on this client. "
                    f"Pass --clsid (e.g. {DEFAULT_SERVER_CLSID})."
                ) from exc
            if winerr == -2147024891 and machine:
                raise PermissionError(
                    f"{exc}\n"
                    f"DCOM Access Denied creating {create_id} on {machine} "
                    f"(even with impersonation). On SERVER check: "
                    f"1) mmc comexp.msc /32 → Matrikon.OPC.Simulation.1 → "
                    f"Launch/Activation Remote for {username or 'your user'} "
                    f"2) My Computer COM Security Limits "
                    f"3) Identity = Interactive User + Simulation running "
                    f"4) Confirm Automation DA works: opc_da_remote_read_test.py"
                ) from exc
            raise
        print("[OK] IOPCServer acquired")

        # AddGroup → request IOPCItemMgt on the group
        print("\n[4] AddGroup...")
        iid_item_mgt = IOPCItemMgt._iid_
        ret = self._server.AddGroup(
            "NativeDaGroup",
            1,  # active
            1000,  # update rate ms
            1,  # client handle
            None,  # time bias
            None,  # deadband
            0,  # LCID
            pointer(iid_item_mgt),
        )
        if not (isinstance(ret, (list, tuple)) and len(ret) >= 3):
            raise OSError(f"Unexpected AddGroup return: {ret!r}")
        self._server_group = int(ret[0])
        revised = int(ret[1])
        unk = ret[2]
        print(f"[OK] AddGroup server_group={self._server_group} revised_rate={revised}")

        # ppUnk may already be IOPCItemMgt or a raw IUnknown
        try:
            self._item_mgt = unk.QueryInterface(IOPCItemMgt)
        except Exception:
            self._item_mgt = unk
        self._sync = self._item_mgt.QueryInterface(IOPCSyncIO)
        print("[OK] IOPCItemMgt + IOPCSyncIO acquired")

    def add_items(self, item_ids: List[str]) -> None:
        assert self._item_mgt is not None and self._OPCITEMDEF is not None
        tags = [t.strip() for t in item_ids if t and t.strip()]
        if not tags:
            raise ValueError("No item IDs to add")

        n = len(tags)
        defs = (self._OPCITEMDEF * n)()
        for i, tag in enumerate(tags):
            defs[i].szAccessPath = ""
            defs[i].szItemID = tag
            defs[i].bActive = 1
            defs[i].hClient = i + 1
            defs[i].dwBlobSize = 0
            defs[i].pBlob = None
            defs[i].vtRequestedDataType = VT_EMPTY
            defs[i].wReserved = 0

        print(f"\n[5] AddItems count={n}...")
        pp_results, pp_errors = self._item_mgt.AddItems(n, defs)

        ok_ids: List[str] = []
        ok_handles: List[int] = []
        for i, tag in enumerate(tags):
            e = self._hresult_at(pp_errors, i)
            if e:
                print(f"[SKIPPED] {tag}: HRESULT=0x{e & 0xFFFFFFFF:08X}")
                continue
            h = int(pp_results[i].hServer)
            ok_ids.append(tag)
            ok_handles.append(h)
            print(f"[ADDED] {tag} server_handle={h}")

        if not ok_handles:
            raise OSError("No OPC items could be added")

        self._item_ids = ok_ids
        self._server_handles = ok_handles
        print(f"[OK] Successfully added {len(ok_handles)} OPC items")

    def read(self, source: int = OPC_DS_DEVICE) -> List[dict]:
        assert self._sync is not None
        n = len(self._server_handles)
        ph = (DWORD * n)(*self._server_handles)
        ret = self._sync.Read(int(source), n, ph)
        if isinstance(ret, (list, tuple)) and len(ret) == 2:
            pp_states, pp_errors = ret
        else:
            raise OSError(f"Unexpected Read return: {ret!r}")

        rows: List[dict] = []
        for i, tag in enumerate(self._item_ids):
            e = self._hresult_at(pp_errors, i)
            if e:
                rows.append(
                    {
                        "tag": tag,
                        "value": None,
                        "quality": 0,
                        "timestamp": None,
                        "error": f"0x{e & 0xFFFFFFFF:08X}",
                    }
                )
                continue
            st = pp_states[i]
            try:
                value: Any = st.vDataValue.value
            except Exception:
                value = repr(st.vDataValue)
            rows.append(
                {
                    "tag": tag,
                    "value": value,
                    "quality": int(st.wQuality),
                    "timestamp": _filetime_to_dt(st.ftTimeStamp),
                    "error": None,
                }
            )
        return rows

    def disconnect(self) -> None:
        try:
            if self._server is not None and self._server_group:
                try:
                    self._server.RemoveGroup(self._server_group, 1)
                except Exception:
                    pass
        finally:
            self._sync = None
            self._item_mgt = None
            self._server = None
            self._item_ids = []
            self._server_handles = []
            self._end_impersonation()
            if self._com_ready:
                try:
                    import pythoncom

                    pythoncom.CoUninitialize()
                except Exception:
                    pass
                self._com_ready = False


def _parse_tags(args_tags: List[str], env_tag: str) -> List[str]:
    tags: List[str] = []
    for t in args_tags or []:
        for part in str(t).split(","):
            part = part.strip()
            if part:
                tags.append(part)
    if not tags and env_tag:
        for part in env_tag.split(","):
            part = part.strip()
            if part:
                tags.append(part)
    # de-dupe preserve order
    seen = set()
    out: List[str] = []
    for t in tags:
        if t not in seen:
            seen.add(t)
            out.append(t)
    return out


def main() -> int:
    if sys.platform != "win32":
        print("[FAIL] Windows only")
        return 1

    import faulthandler

    faulthandler.enable()
    disable_console_quick_edit()
    try:
        sys.stdout.reconfigure(line_buffering=True)  # type: ignore[attr-defined]
    except Exception:
        pass

    pre = argparse.ArgumentParser(add_help=False)
    pre.add_argument("--env", default=DEFAULT_ENV_FILE)
    pre_args, _ = pre.parse_known_args()
    load_dotenv(pre_args.env)

    host_default = env_or("OPC_HOST", DEFAULT_HOST)
    progid_default = env_or("OPC_PROGID", DEFAULT_SERVER_PROGID)
    clsid_default = env_or("OPC_CLSID", DEFAULT_SERVER_CLSID)
    user_default = env_or("OPC_USER", DEFAULT_USER)
    tag_default = env_or("OPC_TAG", DEFAULT_TAG)
    interval_default = float(
        env_or("OPC_DA_INTERVAL", str(DEFAULT_INTERVAL)) or DEFAULT_INTERVAL
    )
    password_env = env_or("OPC_PASSWORD", "")

    parser = argparse.ArgumentParser(
        description="Native OPC DA reader (IOPCServer/IOPCSyncIO, no OPC.Automation.1)"
    )
    parser.add_argument("--env", default=DEFAULT_ENV_FILE)
    parser.add_argument("--host", default=host_default)
    parser.add_argument("--progid", default=progid_default)
    parser.add_argument("--clsid", default=clsid_default)
    parser.add_argument("--user", default=user_default, help="user or DOMAIN\\user")
    parser.add_argument("--domain", default=env_or("OPC_DOMAIN", ""))
    parser.add_argument("--password", default=None)
    parser.add_argument(
        "--tag",
        action="append",
        default=[],
        help="ItemID to read (repeatable). Also accepts comma-separated list.",
    )
    parser.add_argument(
        "--interval",
        type=float,
        default=interval_default,
        help="Seconds between read cycles (default 1)",
    )
    parser.add_argument(
        "--once",
        action="store_true",
        help="Single read then exit (no continuous loop)",
    )
    parser.add_argument(
        "--source",
        choices=("device", "cache"),
        default="cache",
        help="OPC read source (default: cache — closer to Explorer)",
    )
    args = parser.parse_args()

    if args.interval <= 0:
        print("[FAIL] Interval must be greater than zero")
        return 1

    if "\\" in args.user:
        username = args.user
    elif args.domain:
        username = rf"{args.domain}\{args.user}"
    else:
        username = args.user

    password = args.password if args.password is not None else password_env
    if not password and args.host.strip().lower() not in ("", "localhost", "127.0.0.1"):
        password = getpass.getpass(f"Enter password for {username}: ")

    tags = _parse_tags(args.tag, tag_default)
    if not tags:
        print("[FAIL] No tags. Use --tag Random/Int1 or set OPC_TAG in .env")
        return 1

    source = OPC_DS_DEVICE if args.source == "device" else OPC_DS_CACHE
    create_id, create_kind = _resolve_create_id(args.progid, args.clsid)

    _flush_print("=" * 75)
    _flush_print("NATIVE OPC DA READER (IOPC COM/DCOM)")
    _flush_print("=" * 75)
    _flush_print("Host           :", args.host)
    _flush_print("Server ProgID  :", args.progid)
    _flush_print("Server CLSID   :", args.clsid or "(none)")
    _flush_print("CoCreate via   :", f"{create_kind}={create_id}")
    _flush_print("Username       :", username)
    _flush_print("Tags           :", ", ".join(tags))
    _flush_print("Interval       :", args.interval, "s")
    _flush_print("Source         :", args.source)
    _flush_print("Automation     : (none — native IOPCServer)")
    _flush_print("=" * 75)

    if not authenticate_ipc(args.host.strip(), username, password or ""):
        return 2

    client = NativeOpcDaClient()
    try:
        _flush_print("\n[3] Connecting via native COM/DCOM...")
        client.connect(
            args.progid,
            args.host.strip(),
            clsid=args.clsid,
            username=username,
            password=password or "",
        )
        _flush_print("[SUCCESS] Connected")

        client.add_items(tags)

        _flush_print("\n" + "=" * 105)
        _flush_print("LIVE REMOTE OPC DA DATA (NATIVE)")
        if not args.once:
            _flush_print("Press CTRL+C to stop")
        _flush_print("=" * 105)

        cycle = 0
        while True:
            cycle += 1
            _flush_print(
                f"\nREAD CYCLE: {cycle} | "
                f"PC2 TIME: {time.strftime('%Y-%m-%d %H:%M:%S')}"
            )
            _flush_print("-" * 105)
            _flush_print(f"{'TAG':<45}{'VALUE':<18}{'QUALITY':<22}TIMESTAMP (UTC)")
            _flush_print("-" * 105)

            rows = client.read(source=source)
            for row in rows:
                if row.get("error"):
                    _flush_print(f"{row['tag']:<45}READ ERROR: {row['error']}")
                    continue
                _flush_print(
                    f"{row['tag']:<45}"
                    f"{str(row['value']):<18}"
                    f"{decode_quality(int(row['quality'])):<22}"
                    f"{format_timestamp_utc(row['timestamp'])}"
                )
            _flush_print("-" * 105)

            if args.once:
                break
            time.sleep(args.interval)

        _flush_print("\n[SUCCESS] NATIVE OPC DA READ COMPLETED")
        return 0
    except KeyboardInterrupt:
        _flush_print("\n\n[STOPPED] User stopped the OPC reader")
        return 0
    except Exception as exc:
        _flush_print("\n" + "=" * 75)
        _flush_print("[FAIL] NATIVE OPC DA READ FAILED")
        _flush_print("=" * 75)
        _flush_print("ERROR:", exc)
        import traceback

        traceback.print_exc()
        return 1
    finally:
        _flush_print("\n[INFO] Cleaning up OPC connection...")
        client.disconnect()
        _flush_print("[INFO] Client stopped")


if __name__ == "__main__":
    raise SystemExit(main())
