
"""
Remote Matrikon OPC HDA Reader - UTC DISPLAY
============================================

PC1 (server):
    Matrikon.OPC.Simulation.1

PC2 (client):
    Python 3.11 32-bit
    Matrikon.OPCHDA.Automation.1
    pywin32

Features:
    1. IPC$ authentication
    2. COM security initialization
    3. Remote HDA connection
    4. Add historical item
    5. ReadRaw historical data
    6. Print DataValue / Quality as returned
    7. Print TimeStamp in UTC display format only

IMPORTANT:
    - ReadRaw only.
    - No timestamp reconstruction / invented intervals.
    - No year fixing / offset hacks.
    - Timestamp display only: show the COM value in UTC form.
"""

from __future__ import annotations

import argparse
import gc
import getpass
import subprocess
import sys

from datetime import datetime, timedelta, timezone


# ============================================================
# CONFIGURATION
# ============================================================

DEFAULT_HOST = "192.168.1.6"

DEFAULT_HDA_AUTOMATION = "Matrikon.OPCHDA.Automation.1"

DEFAULT_SERVER_PROGID = "IntegrationObjects.AdvancedSimulator.1"

DEFAULT_TAG = "Random/Int1"

DEFAULT_USER = "Swapnil kumar debta"

DEFAULT_DOMAIN = "Swapnildebtasus"

DEFAULT_HISTORY_MINUTES = 10

DEFAULT_NUM_VALUES = 0

DEFAULT_INCLUDE_BOUNDS = False

# OPC DA quality status → HDA Explorer wording
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


# ============================================================
# ERROR HANDLING
# ============================================================

def print_error(exc):

    print(f"ERROR: {exc}")

    if getattr(exc, "args", None):

        print("Details:", exc.args)

        hr = exc.args[0]

        if isinstance(hr, int):

            print(f"HRESULT: 0x{hr & 0xFFFFFFFF:08X}")


# ============================================================
# IPC$ AUTHENTICATION
# ============================================================

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


# ============================================================
# TIME INPUT PARSER
# ============================================================

def parse_time_input(text, default_dt):

    raw = (text or "").strip()

    if not raw:

        return default_dt

    upper = raw.upper().replace(" ", "")

    # NOW
    if upper == "NOW":

        return datetime.now()

    # NOW-10M, NOW-10MIN, NOW-10MINUTES
    if upper.startswith("NOW-"):

        body = upper[4:]

        for suffix in ("MINUTES", "MIN", "M"):

            if body.endswith(suffix):

                body = body[:-len(suffix)]

                break

        return (
            datetime.now()
            - timedelta(minutes=int(body))
        )

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

    raise ValueError(
        f"Unrecognized time: {text}"
    )


# ============================================================
# TIME RANGE INPUT
# ============================================================

def prompt_time_range(history_minutes):

    end_default = datetime.now()

    start_default = (
        end_default
        - timedelta(minutes=history_minutes)
    )

    start_default_s = start_default.strftime(
        "%Y/%m/%d %H:%M:%S"
    )

    end_default_s = end_default.strftime(
        "%Y/%m/%d %H:%M:%S"
    )

    print("\n" + "=" * 70)

    print("READ RAW - TIME RANGE")

    print("=" * 70)

    print(
        f"Default window: past {history_minutes} minutes -> now"
    )

    print("Press Enter to keep defaults.")

    print(
        "Examples: NOW-10M | NOW | 2026/11/09 10:30:00"
    )

    print("-" * 70)

    while True:

        try:

            start_in = input(
                f"Start time [{start_default_s}]: "
            ).strip()

            end_in = input(
                f"End time [{end_default_s}]: "
            ).strip()

        except EOFError:

            start_in = ""

            end_in = ""

        try:

            # IMPORTANT:
            # Relative expressions are resolved here.
            # ReadRaw receives datetime objects, not strings.

            start_dt = parse_time_input(
                start_in,
                start_default,
            )

            end_dt = parse_time_input(
                end_in,
                end_default,
            )

        except ValueError as exc:

            print("[ERROR]", exc)

            continue

        if start_dt >= end_dt:

            print(
                "[ERROR] Start time must be before end time."
            )

            continue

        break

    print("-" * 70)

    print("Resolved start:", repr(start_dt))

    print("Resolved end  :", repr(end_dt))

    print("Input type    :", type(start_dt))

    print("Input type    :", type(end_dt))

    print("=" * 70)

    return start_dt, end_dt


# ============================================================
# QUALITY DECODE (display only — raw COM int unchanged)
# ============================================================

def decode_quality(quality) -> str:
    """Decode OPC HDA Quality integer to Explorer-style text."""
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


# ============================================================
# UTC DISPLAY HELPER (display only — no reconstruction)
# ============================================================

def format_timestamp_utc(timestamp) -> str:
    """
    Display the COM TimeStamp in UTC form.
    Does not invent, reconstruct, or offset the value.
    """
    if timestamp is None:
        return "None"

    try:
        if hasattr(timestamp, "astimezone"):
            if getattr(timestamp, "tzinfo", None) is not None:
                utc_ts = timestamp.astimezone(timezone.utc)
            else:
                # Naive COM DATE: show as UTC clock fields as returned.
                utc_ts = timestamp.replace(tzinfo=timezone.utc)

            ms = int(getattr(utc_ts, "microsecond", 0) // 1000)
            return utc_ts.strftime(f"%Y-%m-%d %H:%M:%S.{ms:03d} UTC")

        return f"{timestamp} UTC"
    except Exception:
        return f"{timestamp!r}"


# ============================================================
# RAW HDA RESULT DISPLAY
# ============================================================

def display_history(result):

    print("\n" + "=" * 115)

    print("MATRIKON OPC HDA READRAW RESPONSE (UTC TIMESTAMP DISPLAY)")

    print("=" * 115)

    if result is None:

        print("[WARNING] ReadRaw returned None")

        return False

    # Extract the returned historical collection.
    # Do not alter its records.

    if isinstance(result, tuple):

        if not result:

            print("[WARNING] Empty result tuple")

            return False

        history = result[0]

        print(
            "\nRETURNED START (UTC display):",
            format_timestamp_utc(result[1]) if len(result) > 1 else None,
        )
        print(
            "RETURNED START (raw repr)   :",
            repr(result[1]) if len(result) > 1 else None,
        )

        print(
            "RETURNED END   (UTC display):",
            format_timestamp_utc(result[2]) if len(result) > 2 else None,
        )
        print(
            "RETURNED END   (raw repr)   :",
            repr(result[2]) if len(result) > 2 else None,
        )

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

            # Read original COM properties directly.

            raw_value = record.DataValue

            raw_timestamp = record.TimeStamp

            raw_quality = record.Quality

            # Value unchanged. Timestamp → UTC display. Quality → Explorer text.

            print(
                f"{index:<8}"
                f"{repr(raw_value):<18}"
                f"{format_timestamp_utc(raw_timestamp):<36}"
                f"{decode_quality(raw_quality)}"
            )

            successful += 1

        except Exception as exc:

            print(
                f"\n[ERROR] Could not read record {index}: {exc}"
            )

    print("-" * 145)

    print("Total records       :", repr(count))

    print("Successfully printed:", successful)

    print("Timestamp display   : UTC format of COM TimeStamp as returned")

    print("=" * 145)

    return True


# ============================================================
# MAIN
# ============================================================

def main():

    if sys.platform != "win32":

        print("[FAIL] Windows only")

        return 1

    parser = argparse.ArgumentParser(
        description="Remote Matrikon OPC HDA RAW ReadRaw client"
    )

    parser.add_argument(
        "--host",
        default=DEFAULT_HOST,
    )

    parser.add_argument(
        "--automation",
        default=DEFAULT_HDA_AUTOMATION,
    )

    parser.add_argument(
        "--progid",
        default=DEFAULT_SERVER_PROGID,
    )

    parser.add_argument(
        "--tag",
        default=DEFAULT_TAG,
    )

    parser.add_argument(
        "--user",
        default=DEFAULT_USER,
    )

    parser.add_argument(
        "--domain",
        default=DEFAULT_DOMAIN,
    )

    parser.add_argument(
        "--password",
        default=None,
    )

    parser.add_argument(
        "--minutes",
        type=int,
        default=DEFAULT_HISTORY_MINUTES,
    )

    parser.add_argument(
        "--num-values",
        type=int,
        default=DEFAULT_NUM_VALUES,
    )

    parser.add_argument(
        "--bounds",
        action="store_true",
        default=DEFAULT_INCLUDE_BOUNDS,
    )

    parser.add_argument(
        "--no-prompt",
        action="store_true",
    )

    args = parser.parse_args()

    host = args.host.strip()

    username = rf"{args.domain}\{args.user}"

    password = args.password

    if password is None:

        password = getpass.getpass(
            f"Enter password for {username}: "
        )

    print("=" * 75)

    print("REMOTE MATRIKON OPC HDA RAW READER")

    print("=" * 75)

    print("PC1 Host      :", host)

    print("Username      :", username)

    print("HDA Automation:", args.automation)

    print("Server ProgID :", args.progid)

    print("Tag           :", args.tag)

    print("=" * 75)

    # STEP 1: IPC authentication

    if not authenticate_ipc(
        host,
        username,
        password,
    ):

        return 2

    # STEP 2: Import COM libraries

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

        # STEP 3: COM initialization

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

                print(
                    "[WARNING] COM security already initialized"
                )

            else:

                raise

        # STEP 4: Create HDA Automation client

        print("\n[4] Creating HDA Automation client...")

        try:

            hda = win32com.client.gencache.EnsureDispatch(
                args.automation
            )

            print("[OK] EnsureDispatch successful")

        except Exception as exc:

            print("[WARNING] EnsureDispatch failed:", exc)

            hda = win32com.client.Dispatch(
                args.automation
            )

            print("[OK] Dispatch successful")

        # STEP 5: Connect to remote HDA server

        print("\n[5] Connecting to remote HDA server...")

        print("ProgID:", args.progid)

        print("Host  :", host)

        hda.Connect(
            args.progid,
            host,
        )

        print("[SUCCESS] Remote HDA connection established")

        try:

            # Request UTC-oriented timestamps from Automation.
            hda.UseUTC = True

            print("[OK] UseUTC = True")

        except Exception as exc:

            print("[WARNING] Could not set UseUTC:", exc)

        # STEP 6: Add historical item

        print("\n[6] Adding historical item...")

        print("Tag:", args.tag)

        hda_item = hda.OPCHDAItems.AddItem(
            args.tag,
            1,
        )

        print("[OK] Historical item added")

        # STEP 7: Select time range

        print("\n[7] Selecting ReadRaw time range...")

        if args.no_prompt:

            end_dt = datetime.now()

            start_dt = (
                end_dt
                - timedelta(minutes=args.minutes)
            )

        else:

            start_dt, end_dt = prompt_time_range(
                args.minutes
            )

        print("\n[OK] Time range prepared")

        print("Start:", repr(start_dt))

        print("End  :", repr(end_dt))

        print("Start type:", type(start_dt))

        print("End type  :", type(end_dt))

        # STEP 8: Execute ReadRaw

        print("\n[8] Calling ReadRaw...")

        print("Tag          :", args.tag)

        print("Start        :", repr(start_dt))

        print("End          :", repr(end_dt))

        print("NumValues    :", args.num_values)

        print("IncludeBounds:", args.bounds)

        print("\n[INFO] Passing datetime objects directly to ReadRaw.")

        # IMPORTANT CHANGE:
        # Do not send "NOW-10M" or "NOW" strings.
        # Pass actual datetime objects to the COM method.

        result = hda_item.ReadRaw(
            start_dt,
            end_dt,
            args.num_values,
            args.bounds,
        )

        print("\n[OK] ReadRaw executed successfully")

        print("ReadRaw return type:", type(result))

        # STEP 9: Display original COM response

        print("\n[9] Displaying original COM response...")

        ok = display_history(result)

        if not ok:

            return 4

        print("\n[SUCCESS] HDA READ COMPLETED")

        print(
            "[INFO] DataValue printed as returned from COM."
        )

        print(
            "[INFO] TimeStamp displayed in UTC form only "
            "(no reconstruction / invented intervals)."
        )

        print(
            "[INFO] Quality decoded to Explorer-style text "
            "(e.g. Good, Non-specific [Raw])."
        )

    except Exception as exc:

        print("\n" + "=" * 75)

        print("[FAIL] REMOTE OPC HDA READ FAILED")

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
