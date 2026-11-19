"""
Native OPC HDA ReadProcessed (CLI) — COM/DCOM custom interfaces
===============================================================

Uses IOPCHDA_Server + IOPCHDA_SyncRead::ReadProcessed directly
(no Matrikon Automation). Same connect path as native ReadRaw.

Requires: 32-bit Python + pywin32 + comtypes

Config: .env (OPC_HOST, OPC_CLSID, OPC_USER, …) then:
  .\\env32\\Scripts\\python.exe opchdareadprocessed_nativecomdcomwrapper.py
  # Enter/Enter for time; Enter/Enter for interval=10s + interpolative
"""

from __future__ import annotations

import argparse
import getpass
import os
import re
import subprocess
import sys
from ctypes import (
    POINTER,
    Structure,
    c_ushort,
    c_void_p,
    c_wchar_p,
    pointer,
)
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from ctypes.wintypes import BOOL, DWORD, FILETIME, WORD

DEFAULT_HOST = "192.168.1.3"
DEFAULT_SERVER_PROGID = "IntegrationObjects.AdvancedSimulator.1"
# Server CLSID (from HKCR\IntegrationObjects.AdvancedSimulator.1\CLSID on PC1).
# Needed when ProgID is not registered on the client (native CoCreate).
DEFAULT_SERVER_CLSID = "{22260278-2009-1112-1982-11121982ABCD}"
DEFAULT_TAG = "Random/Int1"
DEFAULT_USER = "Swapnil kumar debta"
DEFAULT_HISTORY_MINUTES = 10
DEFAULT_INTERVAL_SECONDS = 10.0
DEFAULT_AGGREGATE = "interpolative"
DEFAULT_ENV_FILE = ".env"

# OPC HDA 1.20 aggregate IDs (same as Matrikon Explorer / Automation)
_AGGREGATE_IDS = {
    "interpolative": 1,
    "interpolated": 1,
    "total": 2,
    "average": 3,
    "timeaverage": 4,
    "count": 5,
    "stdev": 6,
    "minimumactualtime": 7,
    "minimum": 8,
    "maximumactualtime": 9,
    "maximum": 10,
    "start": 11,
    "end": 12,
    "delta": 13,
    "regslope": 14,
    "regconst": 15,
    "regdev": 16,
    "variance": 17,
    "range": 18,
    "durationgood": 19,
    "durationbad": 20,
    "percentgood": 21,
    "percentbad": 22,
    "worstquality": 23,
    "annotations": 24,
}
_AGGREGATE_NAMES = {
    v: k.title()
    for k, v in _AGGREGATE_IDS.items()
    if k != "interpolated"
}

IID_IOPCHDA_Server = "{1F1217B0-DEE0-11D2-A5E5-000086339399}"
IID_IOPCHDA_SyncRead = "{1F1217B2-DEE0-11D2-A5E5-000086339399}"

_GUID_RE = re.compile(
    r"^\{?[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-"
    r"[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}?$"
)
_FILETIME_EPOCH = datetime(1601, 1, 1, tzinfo=timezone.utc)


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
    """Return (create_id, kind) where kind is 'clsid' or 'progid'.

    Prefer explicit --clsid. For Integration Objects (ProgID not on client),
    fall back to known CLSID when --clsid is omitted.
    """
    if clsid and str(clsid).strip():
        return _normalize_clsid(str(clsid)), "clsid"
    p = (progid or "").strip()
    if _looks_like_clsid(p):
        return _normalize_clsid(p), "clsid"
    # Remote IO simulator: ProgID lives on server only → use known CLSID.
    if p.lower() == DEFAULT_SERVER_PROGID.lower() and DEFAULT_SERVER_CLSID:
        return _normalize_clsid(DEFAULT_SERVER_CLSID), "clsid"
    return p, "progid"


# ---------------------------------------------------------------------------
# COM structures / interfaces (OPC HDA 1.20)
# ---------------------------------------------------------------------------

class OPCHDA_TIME(Structure):
    # Use c_void_p for szTime — c_wchar_p makes ctypes free/convert the
    # pointer on [in,out] return and hard-crashes the process.
    _pack_ = 8
    _fields_ = [
        ("bString", BOOL),
        ("szTime", c_void_p),
        ("ftTime", FILETIME),
    ]


def _dt_to_filetime(dt: datetime) -> FILETIME:
    if dt.tzinfo is None:
        # Match Automation scripts: treat naive as local wall clock → UTC
        dt = dt.astimezone().astimezone(timezone.utc)
    else:
        dt = dt.astimezone(timezone.utc)
    delta = dt - _FILETIME_EPOCH
    t = int(delta.total_seconds() * 10_000_000)
    return FILETIME(t & 0xFFFFFFFF, (t >> 32) & 0xFFFFFFFF)


def _filetime_to_dt(ft: FILETIME) -> datetime:
    t = ft.dwLowDateTime | (ft.dwHighDateTime << 32)
    return _FILETIME_EPOCH + timedelta(microseconds=t // 10)


def _make_opchda_time(dt: datetime) -> OPCHDA_TIME:
    ht = OPCHDA_TIME()
    ht.bString = 0
    ht.szTime = None
    ht.ftTime = _dt_to_filetime(dt)
    return ht


def _seconds_to_filetime_interval(seconds: float) -> FILETIME:
    """Native ReadProcessed resample interval is FILETIME duration (100ns ticks)."""
    if seconds <= 0:
        raise ValueError("Resample interval must be > 0")
    ticks = int(float(seconds) * 10_000_000)
    return FILETIME(ticks & 0xFFFFFFFF, (ticks >> 32) & 0xFFFFFFFF)


def resolve_aggregate(name_or_id: str) -> Tuple[int, str]:
    raw = str(name_or_id).strip()
    if not raw:
        raise ValueError("Aggregate is empty")
    if raw.isdigit():
        agg_id = int(raw)
        name = _AGGREGATE_NAMES.get(agg_id, f"Aggregate-{agg_id}")
        return agg_id, name
    key = raw.lower().replace(" ", "").replace("_", "").replace("-", "")
    if key not in _AGGREGATE_IDS:
        known = ", ".join(sorted(set(_AGGREGATE_IDS) - {"interpolated"}))
        raise ValueError(f"Unknown aggregate '{name_or_id}'. Known: {known}")
    agg_id = _AGGREGATE_IDS[key]
    return agg_id, key.title()


def _build_interfaces():
    """Lazy-build comtypes interfaces (import-heavy)."""
    from comtypes import COMMETHOD, GUID, HRESULT, IUnknown
    from comtypes.automation import VARIANT

    class OPCHDA_ITEM(Structure):
        _pack_ = 8
        _fields_ = [
            ("hClient", DWORD),
            ("haAggregate", DWORD),
            ("dwCount", DWORD),
            ("pftTimeStamps", POINTER(FILETIME)),
            ("pdwQualities", POINTER(DWORD)),
            ("pvDataValues", POINTER(VARIANT)),
        ]

    class IOPCHDA_Server(IUnknown):
        _iid_ = GUID(IID_IOPCHDA_Server)
        _methods_ = [
            COMMETHOD(
                [],
                HRESULT,
                "GetItemAttributes",
                (["out"], POINTER(DWORD), "pdwCount"),
                (["out"], POINTER(POINTER(DWORD)), "ppdwAttrID"),
                (["out"], POINTER(POINTER(c_wchar_p)), "ppszAttrName"),
                (["out"], POINTER(POINTER(c_wchar_p)), "ppszAttrDesc"),
                (["out"], POINTER(POINTER(c_ushort)), "ppvtAttrDataType"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "GetAggregates",
                (["out"], POINTER(DWORD), "pdwCount"),
                (["out"], POINTER(POINTER(DWORD)), "ppdwAggrID"),
                (["out"], POINTER(POINTER(c_wchar_p)), "ppszAggrName"),
                (["out"], POINTER(POINTER(c_wchar_p)), "ppszAggrDesc"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "GetHistorianStatus",
                (["out"], POINTER(WORD), "pwStatus"),
                (["out"], POINTER(POINTER(FILETIME)), "pftCurrentTime"),
                (["out"], POINTER(POINTER(FILETIME)), "pftStartTime"),
                (["out"], POINTER(WORD), "pwMajorVersion"),
                (["out"], POINTER(WORD), "pwMinorVersion"),
                (["out"], POINTER(WORD), "pwBuildNumber"),
                (["out"], POINTER(DWORD), "pdwMaxReturnValues"),
                (["out"], POINTER(c_wchar_p), "ppszStatusString"),
                (["out"], POINTER(c_wchar_p), "ppszVendorInfo"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "GetItemHandles",
                (["in"], DWORD, "dwCount"),
                (["in"], POINTER(c_wchar_p), "pszItemID"),
                (["in"], POINTER(DWORD), "phClient"),
                (["out"], POINTER(POINTER(DWORD)), "pphServer"),
                (["out"], POINTER(POINTER(HRESULT)), "ppErrors"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "ReleaseItemHandles",
                (["in"], DWORD, "dwCount"),
                (["in"], POINTER(DWORD), "phServer"),
                (["out"], POINTER(POINTER(HRESULT)), "ppErrors"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "ValidateItemIDs",
                (["in"], DWORD, "dwCount"),
                (["in"], POINTER(c_wchar_p), "pszItemID"),
                (["out"], POINTER(POINTER(HRESULT)), "ppErrors"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "CreateBrowse",
                (["in"], DWORD, "dwCount"),
                (["in"], POINTER(DWORD), "pdwAttrID"),
                (["in"], c_void_p, "pOperator"),
                (["in"], POINTER(VARIANT), "vFilter"),
                (["out"], POINTER(c_void_p), "pphBrowser"),
                (["out"], POINTER(POINTER(HRESULT)), "ppErrors"),
            ),
        ]

    class IOPCHDA_SyncRead(IUnknown):
        _iid_ = GUID(IID_IOPCHDA_SyncRead)
        _methods_ = [
            # [in,out] must be POINTER(T) — comtypes reads argtype._type_.
            # Pass pointer(OPCHDA_TIME()) at the call site (not bare struct).
            COMMETHOD(
                [],
                HRESULT,
                "ReadRaw",
                (["in", "out"], POINTER(OPCHDA_TIME), "htStartTime"),
                (["in", "out"], POINTER(OPCHDA_TIME), "htEndTime"),
                (["in"], DWORD, "dwNumValues"),
                (["in"], BOOL, "bBounds"),
                (["in"], DWORD, "dwNumItems"),
                (["in"], POINTER(DWORD), "phServer"),
                (["out"], POINTER(POINTER(OPCHDA_ITEM)), "ppItemValues"),
                (["out"], POINTER(POINTER(HRESULT)), "ppErrors"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "ReadProcessed",
                (["in", "out"], POINTER(OPCHDA_TIME), "htStartTime"),
                (["in", "out"], POINTER(OPCHDA_TIME), "htEndTime"),
                (["in"], FILETIME, "ftResampleInterval"),
                (["in"], DWORD, "dwNumItems"),
                (["in"], POINTER(DWORD), "phServer"),
                (["in"], POINTER(DWORD), "haAggregate"),
                (["out"], POINTER(POINTER(OPCHDA_ITEM)), "ppItemValues"),
                (["out"], POINTER(POINTER(HRESULT)), "ppErrors"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "ReadAtTime",
                (["in"], DWORD, "dwNumTimeStamps"),
                (["in"], POINTER(FILETIME), "ftTimeStamps"),
                (["in"], DWORD, "dwNumItems"),
                (["in"], POINTER(DWORD), "phServer"),
                (["out"], POINTER(POINTER(OPCHDA_ITEM)), "ppItemValues"),
                (["out"], POINTER(POINTER(HRESULT)), "ppErrors"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "ReadModified",
                (["in", "out"], POINTER(OPCHDA_TIME), "htStartTime"),
                (["in", "out"], POINTER(OPCHDA_TIME), "htEndTime"),
                (["in"], DWORD, "dwNumValues"),
                (["in"], DWORD, "dwNumItems"),
                (["in"], POINTER(DWORD), "phServer"),
                (["out"], POINTER(c_void_p), "ppItemValues"),
                (["out"], POINTER(POINTER(HRESULT)), "ppErrors"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "ReadAttribute",
                (["in", "out"], POINTER(OPCHDA_TIME), "htStartTime"),
                (["in", "out"], POINTER(OPCHDA_TIME), "htEndTime"),
                (["in"], DWORD, "hServer"),
                (["in"], DWORD, "dwNumAttributes"),
                (["in"], POINTER(DWORD), "pdwAttributeIDs"),
                (["out"], POINTER(c_void_p), "ppAttributeValues"),
                (["out"], POINTER(POINTER(HRESULT)), "ppErrors"),
            ),
        ]

    return IOPCHDA_Server, IOPCHDA_SyncRead, OPCHDA_ITEM


class NativeOpcHdaClient:
    """Minimal native OPC HDA client (custom COM interfaces over DCOM)."""

    def __init__(self) -> None:
        self._server = None
        self._sync = None
        self._com_ready = False
        self._OPCHDA_ITEM = None
        self.prog_id = ""
        self.clsid = ""
        self.host = ""

    def connect(
        self,
        prog_id: str,
        host: str,
        clsid: Optional[str] = None,
    ) -> None:
        import pythoncom
        from comtypes.client import CreateObject
        from comtypes import CLSCTX_REMOTE_SERVER, CLSCTX_LOCAL_SERVER

        IOPCHDA_Server, IOPCHDA_SyncRead, self._OPCHDA_ITEM = _build_interfaces()

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

        print(
            f"[INFO] CoCreate via {kind}={create_id} "
            f"machine={machine or 'local'} clsctx={clsctx}"
        )
        try:
            self._server = CreateObject(
                create_id,
                machine=machine,
                interface=IOPCHDA_Server,
                clsctx=clsctx,
            )
        except OSError as exc:
            # WinError -2147221005 = CO_E_CLASSSTRING (Invalid class string)
            if kind == "progid" and getattr(exc, "winerror", None) == -2147221005:
                raise OSError(
                    f"{exc}\n"
                    f"ProgID {prog_id!r} is not registered on this client. "
                    f"Pass --clsid (e.g. {DEFAULT_SERVER_CLSID}) instead — "
                    f"native CoCreate needs CLSID when ProgID is only on the server."
                ) from exc
            raise
        self._sync = self._server.QueryInterface(IOPCHDA_SyncRead)
        print("[OK] IOPCHDA_Server + IOPCHDA_SyncRead acquired")

    def probe(self) -> str:
        """comtypes: all GetHistorianStatus args are ['out'] → call with no args."""
        assert self._server is not None
        (
            status,
            _pft_cur,
            _pft_start,
            maj,
            minor,
            build,
            max_ret,
            status_s,
            vendor,
        ) = self._server.GetHistorianStatus()
        v = vendor if isinstance(vendor, str) else (vendor or "")
        s = status_s if isinstance(status_s, str) else (status_s or "")
        return (
            f"status={int(status)} ver={int(maj)}.{int(minor)}.{int(build)} "
            f"maxReturn={int(max_ret)} vendor={v!r} statusStr={s!r}"
        )

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

    def read_processed(
        self,
        item_id: str,
        start_dt: datetime,
        end_dt: datetime,
        interval_seconds: float,
        aggregate_id: int,
    ) -> Tuple[List[dict], OPCHDA_TIME, OPCHDA_TIME]:
        # comtypes: do NOT pass ['out'] args — they are returned.
        # ['in','out'] args are passed in and also returned.
        from ctypes import windll

        assert self._server is not None and self._sync is not None

        count = 1
        ids = (c_wchar_p * count)(item_id)
        client_hs = (DWORD * count)(1)

        pph_server, pp_errors = self._server.GetItemHandles(count, ids, client_hs)
        e0 = self._hresult_at(pp_errors, 0)
        if e0:
            raise OSError(f"GetItemHandles item error HRESULT=0x{e0 & 0xFFFFFFFF:08X}")

        server_handle = int(pph_server[0])
        print(f"[OK] GetItemHandles server_handle={server_handle}")

        samples: List[dict] = []
        ht_start = _make_opchda_time(start_dt)
        ht_end = _make_opchda_time(end_dt)
        ft_interval = _seconds_to_filetime_interval(interval_seconds)
        ha_agg = (DWORD * 1)(int(aggregate_id))

        try:
            ph = (DWORD * 1)(server_handle)
            print(
                f"[INFO] ReadProcessed interval_s={interval_seconds} "
                f"aggregate_id={aggregate_id} handle={server_handle}",
                flush=True,
            )
            # ReadProcessed(htStart, htEnd, ftInterval, dwNumItems, phServer, haAggregate)
            # → (htStart, htEnd, ppItemValues, ppErrors)
            ret = self._sync.ReadProcessed(
                pointer(ht_start),
                pointer(ht_end),
                ft_interval,
                1,
                ph,
                ha_agg,
            )
            print(f"[INFO] ReadProcessed returned type={type(ret).__name__}", flush=True)
            if not (isinstance(ret, (list, tuple)) and len(ret) in (2, 4)):
                raise OSError(f"Unexpected ReadProcessed return: {ret!r}")
            if len(ret) == 4:
                _ht_s, _ht_e, pp_items, pp_read_errs = ret
            else:
                pp_items, pp_read_errs = ret

            e1 = self._hresult_at(pp_read_errs, 0)
            if e1:
                raise OSError(
                    f"ReadProcessed item error HRESULT=0x{e1 & 0xFFFFFFFF:08X}"
                )

            if pp_items is None:
                raise OSError("ReadProcessed returned null ppItemValues")
            item = pp_items[0]
            n = int(item.dwCount)
            print(f"[OK] ReadProcessed dwCount={n}", flush=True)
            for i in range(n):
                ts = _filetime_to_dt(item.pftTimeStamps[i])
                q = int(item.pdwQualities[i])
                var = item.pvDataValues[i]
                try:
                    value: Any = var.value
                except Exception:
                    value = repr(var)
                samples.append(
                    {
                        "timestamp": ts,
                        "value": value,
                        "quality": q,
                    }
                )
        finally:
            ph_rel = (DWORD * 1)(server_handle)
            try:
                self._server.ReleaseItemHandles(1, ph_rel)
            except Exception:
                pass
            try:
                windll.ole32.CoTaskMemFree(pph_server)
            except Exception:
                pass

        return samples, ht_start, ht_end

    def disconnect(self) -> None:
        self._sync = None
        self._server = None
        if self._com_ready:
            try:
                import pythoncom

                pythoncom.CoUninitialize()
            except Exception:
                pass
            self._com_ready = False


def load_dotenv(path: str | Path | None = None, *, override: bool = True) -> Dict[str, str]:
    """Load KEY=VALUE from .env into os.environ.

    override=True (default): .env wins over ambient shell vars so OPC_HOST
    in .env is not silently replaced by a leftover localhost in the session.
    """
    env_path = Path(path or DEFAULT_ENV_FILE)
    if not env_path.is_file():
        # Also try next to this script
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


def parse_time_input(text: str, default_dt: datetime) -> datetime:
    raw = (text or "").strip()
    if not raw:
        return default_dt
    upper = raw.upper().replace(" ", "")
    if upper == "NOW":
        return datetime.now()
    if upper.startswith("NOW-"):
        body = upper[4:]
        for suffix in ("MINUTES", "MIN", "M"):
            if body.endswith(suffix):
                body = body[: -len(suffix)]
                break
        return datetime.now() - timedelta(minutes=int(body))
    formats = (
        "%Y/%m/%d %H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%d-%m-%Y %H:%M:%S",
        "%d/%m/%Y %H:%M:%S",
        "%Y/%m/%d %H:%M",
        "%Y-%m-%d %H:%M",
        "%d-%m-%Y %H:%M",
        "%d/%m/%Y %H:%M",
    )
    for fmt in formats:
        try:
            return datetime.strptime(raw, fmt)
        except ValueError:
            continue
    raise ValueError(f"Unrecognized time: {text}")


def _floor_to_second(dt: datetime) -> datetime:
    """Drop sub-seconds so resample grid matches Explorer (.000)."""
    return dt.replace(microsecond=0)


def prompt_time_range(history_minutes: int) -> Tuple[datetime, datetime]:
    # Floor to whole seconds — Explorer uses .000; fractional start shifts
    # interpolative values even when the clock minute looks the same.
    end_default = _floor_to_second(datetime.now())
    start_default = end_default - timedelta(minutes=history_minutes)
    start_default_s = start_default.strftime("%Y/%m/%d %H:%M:%S")
    end_default_s = end_default.strftime("%Y/%m/%d %H:%M:%S")

    print("\n" + "=" * 70)
    print("READ PROCESSED - TIME RANGE")
    print("=" * 70)
    print(f"Default window: past {history_minutes} minutes -> now")
    print("Press Enter to keep defaults.")
    print("Examples: NOW-10M | NOW | 2026/11/15 21:00:00")
    print("-" * 70)

    while True:
        try:
            start_in = input(f"Start time [{start_default_s}]: ").strip()
            end_in = input(f"End time [{end_default_s}]: ").strip()
        except EOFError:
            start_in = ""
            end_in = ""
        try:
            start_dt = parse_time_input(start_in, start_default)
            end_dt = parse_time_input(end_in, end_default)
        except ValueError as exc:
            print("[ERROR]", exc)
            continue
        if start_dt >= end_dt:
            print("[ERROR] Start time must be before end time.")
            continue
        break

    print("-" * 70)
    print("Resolved start:", start_dt)
    print("Resolved end  :", end_dt)
    print("=" * 70)
    return start_dt, end_dt


def prompt_processed_options(
    default_interval: float, default_aggregate: str
) -> Tuple[float, int, str]:
    print("\n" + "=" * 70)
    print("READ PROCESSED — RESAMPLE / AGGREGATE")
    print("=" * 70)
    print(f"Default interval  : {default_interval} s")
    print(f"Default aggregate : {default_aggregate} (ID from OPC HDA 1.20)")
    print("Press Enter to keep defaults.")
    print("Aggregates: interpolative, average, minimum, maximum, …")
    print("-" * 70)

    while True:
        try:
            interval_in = input(
                f"Resample interval seconds [{default_interval}]: "
            ).strip()
            aggregate_in = input(f"Aggregate [{default_aggregate}]: ").strip()
        except EOFError:
            interval_in = ""
            aggregate_in = ""
        try:
            interval = float(interval_in) if interval_in else float(default_interval)
            if interval <= 0:
                raise ValueError("Resample interval must be > 0")
            aggregate_id, aggregate_name = resolve_aggregate(
                aggregate_in or default_aggregate
            )
        except ValueError as exc:
            print("[ERROR]", exc)
            continue
        break

    print("-" * 70)
    print(f"Using interval  : {interval} s")
    print(f"Using aggregate : {aggregate_name} (ID={aggregate_id})")
    print("=" * 70)
    return interval, aggregate_id, aggregate_name


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


def decode_quality(q: int) -> str:
    """Match Matrikon Explorer-style DA status + HDA flags."""
    try:
        qi = int(q)
    except Exception:
        return str(q)
    da = qi & 0xFF
    status_key = da & 0xFC
    base = _DA_QUALITY_STATUS.get(
        status_key,
        _DA_QUALITY_STATUS.get(da & 0xC0, f"Quality 0x{da:02X}"),
    )
    flags = []
    if qi & 0x00040000:
        flags.append("Raw")
    if qi & 0x00020000:
        flags.append("Interpolated")
    if qi & 0x00080000:
        flags.append("Calculated")
    if qi & 0x00010000:
        flags.append("NoBound")
    if qi & 0x00200000:
        flags.append("NoData")
    if flags:
        return f"{base} [{', '.join(flags)}]"
    return base


def format_timestamp_utc(ts: datetime) -> str:
    if ts.tzinfo is None:
        ts = ts.replace(tzinfo=timezone.utc)
    else:
        ts = ts.astimezone(timezone.utc)
    ms = ts.microsecond // 1000
    return ts.strftime(f"%Y-%m-%d %H:%M:%S.{ms:03d} UTC")


def main() -> int:
    if sys.platform != "win32":
        print("[FAIL] Windows only")
        return 1

    import faulthandler

    faulthandler.enable()

    pre = argparse.ArgumentParser(add_help=False)
    pre.add_argument("--env", default=DEFAULT_ENV_FILE)
    pre_args, _ = pre.parse_known_args()
    load_dotenv(pre_args.env)

    host_default = env_or("OPC_HOST", DEFAULT_HOST)
    progid_default = env_or("OPC_PROGID", DEFAULT_SERVER_PROGID)
    clsid_default = env_or("OPC_CLSID", DEFAULT_SERVER_CLSID)
    user_default = env_or("OPC_USER", DEFAULT_USER)
    tag_default = env_or("OPC_TAG", DEFAULT_TAG)
    minutes_default = int(
        env_or("OPC_MINUTES", str(DEFAULT_HISTORY_MINUTES)) or DEFAULT_HISTORY_MINUTES
    )
    interval_default = float(
        env_or("OPC_INTERVAL", str(DEFAULT_INTERVAL_SECONDS))
        or DEFAULT_INTERVAL_SECONDS
    )
    aggregate_default = env_or("OPC_AGGREGATE", DEFAULT_AGGREGATE) or DEFAULT_AGGREGATE
    password_env = env_or("OPC_PASSWORD", "")

    parser = argparse.ArgumentParser(
        description=(
            "Native OPC HDA ReadProcessed (IOPCHDA COM/DCOM, no Matrikon Automation)"
        )
    )
    parser.add_argument("--env", default=DEFAULT_ENV_FILE, help="Path to .env file")
    parser.add_argument("--host", default=host_default)
    parser.add_argument(
        "--progid",
        default=progid_default,
        help="Server ProgID (only works if registered locally)",
    )
    parser.add_argument(
        "--clsid",
        default=clsid_default,
        help=(
            "Server CLSID for native CoCreate when ProgID is not registered "
            f"locally (IO default auto: {DEFAULT_SERVER_CLSID})"
        ),
    )
    parser.add_argument("--user", default=user_default, help="user or DOMAIN\\user")
    parser.add_argument("--domain", default=env_or("OPC_DOMAIN", ""))
    parser.add_argument(
        "--password",
        default=None,
        help="Override .env OPC_PASSWORD (omit to use .env or prompt)",
    )
    parser.add_argument("--tag", default=tag_default)
    parser.add_argument("--minutes", type=int, default=minutes_default)
    parser.add_argument(
        "--interval",
        type=float,
        default=interval_default,
        help="Resample interval in seconds (native FILETIME duration)",
    )
    parser.add_argument(
        "--aggregate",
        default=aggregate_default,
        help="Aggregate name or numeric ID (default: interpolative)",
    )
    parser.add_argument(
        "--no-prompt",
        action="store_true",
        help="Skip time/interval prompts; use defaults",
    )
    args = parser.parse_args()

    if "\\" in args.user:
        username = args.user
    elif args.domain:
        username = rf"{args.domain}\{args.user}"
    else:
        username = args.user

    password = args.password if args.password is not None else password_env
    if not password:
        password = getpass.getpass(f"Enter password for {username}: ")

    try:
        default_agg_id, default_agg_name = resolve_aggregate(args.aggregate)
    except ValueError as exc:
        print("[FAIL]", exc)
        return 1

    if args.no_prompt:
        end_dt = _floor_to_second(datetime.now())
        start_dt = end_dt - timedelta(minutes=args.minutes)
        interval_s = float(args.interval)
        aggregate_id = default_agg_id
        aggregate_name = default_agg_name
    else:
        start_dt, end_dt = prompt_time_range(args.minutes)
        interval_s, aggregate_id, aggregate_name = prompt_processed_options(
            args.interval, args.aggregate
        )

    create_id, create_kind = _resolve_create_id(args.progid, args.clsid)

    print("=" * 75)
    print("NATIVE OPC HDA READPROCESSED (IOPCHDA COM/DCOM)")
    print("=" * 75)
    print("Host           :", args.host)
    print("Server ProgID  :", args.progid)
    print("Server CLSID   :", args.clsid or "(none)")
    print("CoCreate via   :", f"{create_kind}={create_id}")
    print("Username       :", username)
    print("Tag / ItemID   :", args.tag)
    print("Start          :", start_dt)
    print("End            :", end_dt)
    print("Interval       :", f"{interval_s} s")
    print("Aggregate      :", f"{aggregate_name} (ID={aggregate_id})")
    print("Automation     : (none — native IOPCHDA)")
    print("=" * 75)

    if not authenticate_ipc(args.host.strip(), username, password):
        return 2

    client = NativeOpcHdaClient()
    try:
        print("\n[3] Connecting via native COM/DCOM...")
        client.connect(args.progid, args.host.strip(), clsid=args.clsid)
        print("[SUCCESS] Connected")

        print("\n[4] GetHistorianStatus probe...")
        try:
            print("[OK]", client.probe())
        except Exception as exc:
            print("[WARNING] probe:", exc)

        print("\n[5] ReadProcessed...")
        samples, ht_s, ht_e = client.read_processed(
            args.tag, start_dt, end_dt, interval_s, aggregate_id
        )

        print("\n" + "=" * 100)
        print(
            f"SERVER VALUES — READPROCESSED ({aggregate_name}, interval={interval_s}s) — UTC"
        )
        print("=" * 100)
        print(f"{'NO':<6}{'DATAVALUE':<18}{'TIMESTAMP (UTC)':<32}{'QUALITY'}")
        print("-" * 100)
        for i, s in enumerate(samples, 1):
            print(
                f"{i:<6}"
                f"{repr(s['value']):<18}"
                f"{format_timestamp_utc(s['timestamp']):<32}"
                f"{decode_quality(int(s['quality']))}"
            )
        print("-" * 100)
        print(f"Total records: {len(samples)}")
        print("=" * 100)
        print("\n[SUCCESS] NATIVE HDA READPROCESSED COMPLETED")
        return 0
    except Exception as exc:
        print("\n" + "=" * 75)
        print("[FAIL] NATIVE OPC HDA READPROCESSED FAILED")
        print("=" * 75)
        print("ERROR:", exc)
        import traceback

        traceback.print_exc()
        return 1
    finally:
        client.disconnect()
        print("[INFO] Disconnected")


if __name__ == "__main__":
    raise SystemExit(main())
