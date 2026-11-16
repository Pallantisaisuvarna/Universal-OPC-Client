"""
Remote Matrikon OPC HDA Reader - ReadProcessed (Interpolative)
==============================================================

PC1 (server): Matrikon.OPC.Simulation.1
PC2 (client): Python 3.11 32-bit + Matrikon.OPCHDA.Automation.1

Same remote pattern as opc_hda_remote1.py (ReadRaw), but calls ReadProcessed
with INTERPOLATIVE aggregate — like Matrikon HDA Explorer "Read Processed".

Features:
1. IPC$ authentication
2. COM security
3. Remote HDA connection
4. AddItem
5. ReadProcessed (Interpolative) with resample interval
6. UTC timestamp display + Explorer-style quality decode

IMPORTANT:
- ResampleInterval is a VT_DATE *duration* (seconds/86400 days), not a 2020 calendar date.
- UseUTC=True — same as working opc_hda_remote1.py ReadRaw.
- TimeStamp: print whatever Automation returns, displayed in UTC only.
- No timestamp reconstruction / invented history times.
- Returned DataValue / TimeStamp / Quality come from the server as-is.
"""

from __future__ import annotations

import argparse
import gc
import getpass
import subprocess
import sys

from datetime import datetime, timedelta, timezone


DEFAULT_HOST = "192.168.1.141"
DEFAULT_HDA_AUTOMATION = "Matrikon.OPCHDA.Automation.1"
DEFAULT_SERVER_PROGID = "Matrikon.OPC.Simulation.1"
DEFAULT_TAG = "Random.Int1"
DEFAULT_USER = "DELL"
DEFAULT_DOMAIN = "DESKTOP-E9295IQ"
DEFAULT_HISTORY_MINUTES = 10
DEFAULT_INTERVAL_SECONDS = 10.0
DEFAULT_AGGREGATE = "interpolative"


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

_HDA_RAW = 0x00040000
_HDA_INTERPOLATED = 0x00020000
_HDA_CALCULATED = 0x00080000
_HDA_NOBOUND = 0x00010000
_HDA_NODATA = 0x00200000
_HDA_DATALOST = 0x00100000
_HDA_CONVERSION = 0x00400000
_HDA_PARTIAL = 0x00800000


def print_error(exc):
    print(f"ERROR: {exc}")
    if getattr(exc, "args", None):
        print("Details:", exc.args)
        hr = exc.args[0]
        if isinstance(hr, int):
            print(f"HRESULT: 0x{hr & 0xFFFFFFFF:08X}")


def authenticate_ipc(host, username, password):
    ipc_path = rf"\\{host}\IPC$"

    print("\n[1] Clearing existing IPC$ connection...")
    subprocess.run(
        ["net", "use", ipc_path, "/delete", "/y"],
        capture_output=True,
        text=True,
    )

    print("[2] Authenticating IPC$...")
    result = subprocess.run(
        [
            "net",
            "use",
            ipc_path,
            f"/user:{username}",
            password,
        ],
        capture_output=True,
        text=True,
    )

    if result.returncode != 0:
        print("[FAIL] IPC$ authentication failed")
        print(result.stderr or result.stdout)
        return False

    print("[OK] IPC$ authentication successful")
    return True


def resolve_aggregate(name_or_id):
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
        raise ValueError(
            f"Unknown aggregate '{name_or_id}'. Known: {known}"
        )
    agg_id = _AGGREGATE_IDS[key]
    return agg_id, key.title()


def ole_duration_days(seconds: float) -> float:
    """True OPC/OLE resample duration: seconds as fraction of a day."""
    if seconds <= 0:
        raise ValueError("Resample interval must be > 0")
    return float(seconds) / 86400.0


def seconds_to_hda_interval_candidates(seconds: float):
    """
    ResampleInterval must be a VT_DATE *duration* (= seconds/86400 days).

    IMPORTANT:
      datetime(2020,1,1,0,0,10) is NOT a 10-second duration.
      Its OLE DATE is ~43831 days → server returns almost no useful points
      (often a single NoData record).

    Prefer a real duration value. UTC-aware OLE epoch avoids IST Errno 22.
    """
    import pywintypes

    seconds = float(seconds)
    ole_days = ole_duration_days(seconds)
    candidates = []

    # 1) True duration via UTC-aware OLE epoch + seconds (VT_DATE serial ≈ ole_days)
    try:
        ole_epoch_utc = datetime(1899, 12, 30, tzinfo=timezone.utc)
        duration_time = pywintypes.Time(
            ole_epoch_utc + timedelta(seconds=seconds)
        )
        candidates.append(
            ("pywintypes.Time OLE-epoch UTC duration", duration_time, "wrapper")
        )
    except Exception as exc:
        print("[WARNING] Could not build UTC OLE-epoch duration Time:", exc)

    # 2) Direct InvokeTypes with OLE float days (bypass wrapper Time conversion)
    candidates.append(
        ("InvokeTypes VT_DATE ole_days float", ole_days, "invoke")
    )

    if not candidates:
        raise RuntimeError("Could not build any interval encoding")

    return candidates


def call_read_processed(hda_item, start_dt, end_dt, interval_value, aggregate_id, mode: str):
    """
    mode='wrapper' → normal ReadProcessed(...)
    mode='invoke'  → low-level InvokeTypes with VT_DATE as float days
    """
    if mode == "wrapper":
        return hda_item.ReadProcessed(
            start_dt,
            end_dt,
            interval_value,
            aggregate_id,
        )

    # DISPID ReadProcessed = 1610743813 from Matrikon typelib
    # Arg types: VARIANT, VARIANT, DATE, I4
    return hda_item._oleobj_.InvokeTypes(
        1610743813,
        0,
        1,
        (12, 0),
        ((16396, 3), (16396, 3), (7, 1), (3, 1)),
        start_dt,
        end_dt,
        float(interval_value),
        int(aggregate_id),
    )


def parse_time_input(text, default_dt):
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


def prompt_time_range(history_minutes):
    end_default = datetime.now()
    start_default = end_default - timedelta(minutes=history_minutes)
    start_default_s = start_default.strftime("%Y/%m/%d %H:%M:%S")
    end_default_s = end_default.strftime("%Y/%m/%d %H:%M:%S")

    print("\n" + "=" * 70)
    print("READ PROCESSED — TIME RANGE")
    print("=" * 70)
    print(f"Default window: past {history_minutes} minutes -> now")
    print("Press Enter to keep defaults.")
    print("Examples: NOW-10M | NOW | 2026/11/09 10:30:00")
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
    print("Resolved start:", repr(start_dt))
    print("Resolved end  :", repr(end_dt))
    print("=" * 70)
    return start_dt, end_dt


def prompt_processed_options(default_interval: float, default_aggregate: str):
    print("\n" + "=" * 70)
    print("READ PROCESSED — RESAMPLE / AGGREGATE")
    print("=" * 70)
    print(f"Default interval  : {default_interval} s  (Explorer: 10 Seconds)")
    print(f"Default aggregate : {default_aggregate} (Explorer: INTERPOLATIVE)")
    print("Press Enter to keep defaults.")
    print("-" * 70)

    while True:
        try:
            interval_in = input(
                f"Resample interval seconds [{default_interval}]: "
            ).strip()
            aggregate_in = input(
                f"Aggregate [{default_aggregate}]: "
            ).strip()
        except EOFError:
            interval_in = ""
            aggregate_in = ""

        try:
            interval = (
                float(interval_in) if interval_in else float(default_interval)
            )
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


def decode_quality(quality) -> str:
    try:
        q = int(quality)
    except Exception:
        return str(quality)

    da_status = q & 0xFF
    status_key = da_status & 0xFC
    da_text = _DA_QUALITY_STATUS.get(
        status_key,
        _DA_QUALITY_STATUS.get(da_status & 0xC0, f"Quality 0x{da_status:02X}"),
    )

    hda_flags = []
    if q & _HDA_RAW:
        hda_flags.append("Raw")
    if q & _HDA_INTERPOLATED:
        hda_flags.append("Interpolated")
    if q & _HDA_CALCULATED:
        hda_flags.append("Calculated")
    if q & _HDA_NOBOUND:
        hda_flags.append("NoBound")
    if q & _HDA_NODATA:
        hda_flags.append("NoData")
    if q & _HDA_DATALOST:
        hda_flags.append("DataLost")
    if q & _HDA_CONVERSION:
        hda_flags.append("Conversion")
    if q & _HDA_PARTIAL:
        hda_flags.append("Partial")

    if hda_flags:
        return f"{da_text} [{', '.join(hda_flags)}]"
    return da_text


def format_timestamp_utc(timestamp) -> str:
    if timestamp is None:
        return "None"
    try:
        if hasattr(timestamp, "astimezone"):
            if getattr(timestamp, "tzinfo", None) is not None:
                utc_ts = timestamp.astimezone(timezone.utc)
            else:
                utc_ts = timestamp.replace(tzinfo=timezone.utc)
            ms = int(getattr(utc_ts, "microsecond", 0) // 1000)
            return utc_ts.strftime(f"%Y-%m-%d %H:%M:%S.{ms:03d} UTC")
        return f"{timestamp} UTC"
    except Exception:
        return f"{timestamp!r}"


def display_history(result, interval_s: float, aggregate_name: str):
    print("\n" + "=" * 115)
    print(
        f"MATRIKON OPC HDA READPROCESSED RESPONSE "
        f"({aggregate_name}, interval={interval_s}s)"
    )
    print("=" * 115)

    if result is None:
        print("[WARNING] ReadProcessed returned None")
        return False

    if isinstance(result, tuple):
        if not result:
            print("[WARNING] Empty result tuple")
            return False
        history = result[0]
        if len(result) > 1:
            print(
                "\nRETURNED START (UTC display):",
                format_timestamp_utc(result[1]),
            )
            print("RETURNED START (raw repr)   :", repr(result[1]))
        if len(result) > 2:
            print(
                "RETURNED END   (UTC display):",
                format_timestamp_utc(result[2]),
            )
            print("RETURNED END   (raw repr)   :", repr(result[2]))
    else:
        history = result

    if history is None:
        print("[WARNING] Historical collection is None")
        return False

    try:
        count = history.Count
    except Exception as exc:
        print("[FAIL] Could not read history.Count")
        print_error(exc)
        return False

    print("\nTOTAL HISTORICAL RECORDS:", repr(count))
    print("Aggregate              :", aggregate_name)
    print("Resample interval      :", interval_s, "s")

    if count == 0:
        print("[WARNING] No historical records found")
        return True

    print("\n" + "=" * 145)
    print("SERVER VALUES — TIMESTAMP SHOWN IN UTC (NO RECONSTRUCTION)")
    print("=" * 145)
    print(
        f"{'NO':<8}"
        f"{'DATAVALUE':<18}"
        f"{'TIMESTAMP (UTC)':<36}"
        f"{'QUALITY'}"
    )
    print("-" * 145)

    successful = 0
    for index in range(1, count + 1):
        try:
            record = history.Item(index)
            raw_value = record.DataValue
            raw_timestamp = record.TimeStamp
            raw_quality = record.Quality
            print(
                f"{index:<8}"
                f"{repr(raw_value):<18}"
                f"{format_timestamp_utc(raw_timestamp):<36}"
                f"{decode_quality(raw_quality)}"
            )
            successful += 1
        except Exception as exc:
            print(f"\n[ERROR] Could not read record {index}: {exc}")

    print("-" * 145)
    print("Total records       :", repr(count))
    print("Successfully printed:", successful)
    print("Timestamp display   : UTC format of COM TimeStamp as returned")
    print("=" * 145)
    return True


def main():
    if sys.platform != "win32":
        print("[FAIL] Windows only")
        return 1

    parser = argparse.ArgumentParser(
        description=(
            "Remote Matrikon OPC HDA ReadProcessed client "
            "(Interpolative / Explorer-style)"
        )
    )
    parser.add_argument("--host", default=DEFAULT_HOST)
    parser.add_argument("--automation", default=DEFAULT_HDA_AUTOMATION)
    parser.add_argument("--progid", default=DEFAULT_SERVER_PROGID)
    parser.add_argument("--tag", default=DEFAULT_TAG)
    parser.add_argument("--user", default=DEFAULT_USER)
    parser.add_argument("--domain", default=DEFAULT_DOMAIN)
    parser.add_argument("--password", default=None)
    parser.add_argument(
        "--minutes",
        type=int,
        default=DEFAULT_HISTORY_MINUTES,
    )
    parser.add_argument(
        "--interval",
        type=float,
        default=DEFAULT_INTERVAL_SECONDS,
        help="Resample interval in seconds (Explorer default often 10)",
    )
    parser.add_argument(
        "--aggregate",
        default=DEFAULT_AGGREGATE,
        help="Aggregate name or ID (default: interpolative)",
    )
    parser.add_argument("--no-prompt", action="store_true")

    args = parser.parse_args()

    try:
        default_agg_id, default_agg_name = resolve_aggregate(args.aggregate)
    except ValueError as exc:
        print(f"[FAIL] {exc}")
        return 1

    if args.interval <= 0:
        print("[FAIL] --interval must be > 0")
        return 1

    host = args.host.strip()
    username = rf"{args.domain}\{args.user}"
    password = args.password
    if password is None:
        password = getpass.getpass(f"Enter password for {username}: ")

    print("=" * 75)
    print("REMOTE MATRIKON OPC HDA READPROCESSED READER")
    print("=" * 75)
    print("PC1 Host      :", host)
    print("Username      :", username)
    print("HDA Automation:", args.automation)
    print("Server ProgID :", args.progid)
    print("Tag           :", args.tag)
    print("Interval      :", args.interval, "s")
    print(
        "Aggregate     :",
        f"{default_agg_name} (ID={default_agg_id})",
    )
    print("=" * 75)

    if not authenticate_ipc(host, username, password):
        return 2

    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        print("[FAIL] pywin32 is not installed")
        print_error(exc)
        return 3

    com_initialized = False
    hda = None
    hda_item = None
    result = None

    try:
        print("\n[3] Initializing COM security...")
        pythoncom.CoInitialize()
        com_initialized = True

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
            print("[OK] COM security initialized")
        except pythoncom.com_error as exc:
            hr = exc.args[0] if exc.args else None
            if hr == -2147417831:
                print("[WARNING] COM security already initialized")
            else:
                raise

        print("\n[4] Creating HDA Automation client...")
        try:
            hda = win32com.client.gencache.EnsureDispatch(args.automation)
            print("[OK] EnsureDispatch successful")
        except Exception as exc:
            print("[WARNING] EnsureDispatch failed:", exc)
            hda = win32com.client.Dispatch(args.automation)
            print("[OK] Dispatch successful")

        print("\n[5] Connecting to remote HDA server...")
        print("ProgID:", args.progid)
        print("Host  :", host)
        hda.Connect(args.progid, host)
        print("[SUCCESS] Remote HDA connection established")

        try:
            # Same as working ReadRaw (opc_hda_remote1.py).
            hda.UseUTC = True
            print("[OK] UseUTC = True (same as ReadRaw)")
        except Exception as exc:
            print("[WARNING] Could not set UseUTC:", exc)

        print("\n[6] Adding historical item...")
        print("Tag:", args.tag)
        hda_item = hda.OPCHDAItems.AddItem(args.tag, 1)
        print("[OK] Historical item added")

        print("\n[7] Selecting ReadProcessed time range...")
        if args.no_prompt:
            end_dt = datetime.now()
            start_dt = end_dt - timedelta(minutes=args.minutes)
            interval_s = float(args.interval)
            aggregate_id = default_agg_id
            aggregate_name = default_agg_name
        else:
            start_dt, end_dt = prompt_time_range(args.minutes)
            interval_s, aggregate_id, aggregate_name = prompt_processed_options(
                args.interval, args.aggregate
            )

        # Query bounds only (do not alter returned history values).
        start_dt = start_dt.replace(microsecond=0)
        end_dt = end_dt.replace(microsecond=0)

        print("\n[OK] Time range / processed options prepared")
        print("Start         :", repr(start_dt))
        print("End           :", repr(end_dt))
        print("Interval (s)  :", interval_s)
        print("Aggregate     :", f"{aggregate_name} (ID={aggregate_id})")

        print("\n[8] Calling ReadProcessed...")
        print("Tag       :", args.tag)
        print("Start     :", repr(start_dt))
        print("End       :", repr(end_dt))
        print("Interval  :", interval_s, "s")
        print("Aggregate :", f"{aggregate_name} (ID={aggregate_id})")

        result = None
        last_exc = None
        ole_days = ole_duration_days(interval_s)
        print(f"[INFO] True OLE duration days (10s => ~{10/86400:.12f}): {ole_days:.12f}")

        for label, interval_value, mode in seconds_to_hda_interval_candidates(interval_s):
            print(f"\n[INFO] Trying interval encoding: {label}")
            print("Interval value:", repr(interval_value), type(interval_value))
            print("Call mode     :", mode)
            try:
                result = call_read_processed(
                    hda_item,
                    start_dt,
                    end_dt,
                    interval_value,
                    aggregate_id,
                    mode,
                )
                print(f"[OK] ReadProcessed succeeded with: {label}")
                break
            except Exception as exc:
                last_exc = exc
                print(f"[WARNING] ReadProcessed failed with {label}: {exc}")

        if result is None:
            raise last_exc if last_exc is not None else RuntimeError(
                "ReadProcessed failed for all interval encodings"
            )

        print("\n[OK] ReadProcessed executed successfully")
        print("ReadProcessed return type:", type(result))

        print("\n[9] Displaying COM response...")
        ok = display_history(result, interval_s, aggregate_name)
        if not ok:
            return 4

        print("\n[SUCCESS] HDA READPROCESSED COMPLETED")
        print("[INFO] Same pattern as ReadRaw: UseUTC=True, server TimeStamp shown in UTC only.")
        print("[INFO] No timestamp reconstruction. Quality decoded to Explorer text.")

    except Exception as exc:
        print("\n" + "=" * 75)
        print("[FAIL] REMOTE OPC HDA READPROCESSED FAILED")
        print("=" * 75)
        print_error(exc)
        return 1

    finally:
        print("\n[INFO] Cleaning up...")
        result = None
        hda_item = None
        if hda is not None:
            try:
                hda.Disconnect()
                print("[OK] HDA server disconnected")
            except Exception as exc:
                print("[WARNING] Disconnect failed:", exc)
        hda = None
        if com_initialized:
            gc.collect()
            pythoncom.CoUninitialize()
        print("[INFO] COM cleanup completed")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
