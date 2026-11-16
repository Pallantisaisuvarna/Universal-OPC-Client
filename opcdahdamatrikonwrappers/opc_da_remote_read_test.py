
"""
Dynamic Remote Matrikon OPC DA Reader
======================================

PC1: Matrikon OPC Simulation Server
     IP: 192.168.1.47

PC2: Python 32-bit OPC DA Client

Features:
1. IPC$ authentication
2. COM security initialization
3. Remote OPC Automation connection
4. Automatic browsing of existing OPC server items
5. Creation of client-side OPC group
6. Addition of all discovered server items
7. Reading all item values, quality and timestamps
8. UTC to PC2 local time conversion
9. Continuous terminal output

IMPORTANT:
- Does not create new tags on PC1.
- Reads existing OPC items exposed by the server.
- Runs the Python client only on PC2.
- Requires OPC.Automation.1 registered on PC2.
"""

from __future__ import annotations

import argparse
import getpass
import subprocess
import sys
import time

from datetime import timezone


DEFAULT_HOST = "192.168.1.6"
DEFAULT_PROGID = "IntegrationObjects.AdvancedSimulator.1"
DEFAULT_USER = "Swapnil kumar debta"
DEFAULT_DOMAIN = "Swapnildebtasus"
DEFAULT_INTERVAL = 1.0
DEFAULT_FILTER = ""


def print_error(exc):
    print(f"Error: {exc}")

    if getattr(exc, "args", None):
        print(f"Details: {exc.args}")

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


def convert_to_local_time(timestamp):

    """
    Convert OPC UTC timestamp to PC2 local time.
    Safely handles uninitialized / 1601 epoch timestamps.
    """

    if timestamp is None:
        return "N/A"

    try:
        if getattr(timestamp, "year", 2000) < 1970:
            return str(timestamp)

        if timestamp.tzinfo is None:
            timestamp = timestamp.replace(
                tzinfo=timezone.utc
            )

        return timestamp.astimezone().strftime("%Y-%m-%d %H:%M:%S.%f %Z")
    except Exception:
        return str(timestamp)


def browse_opc_tags(opc_server):

    """
    Discover existing OPC item IDs recursively.

    Does not create tags on the OPC server.
    """

    print("\n[6] Browsing existing OPC server items...")

    browser = opc_server.CreateBrowser()

    discovered_tags = set()

    def browse_level():

        # Get branches at the current location.
        browser.ShowBranches()

        branch_names = [
            str(browser.Item(i))
            for i in range(1, browser.Count + 1)
        ]

        # Get leaf items at the current location.
        browser.ShowLeafs()

        leaf_names = [
            str(browser.Item(i))
            for i in range(1, browser.Count + 1)
        ]

        # Resolve leaf names to fully qualified OPC item IDs.
        for leaf in leaf_names:

            try:

                item_id = str(
                    browser.GetItemID(leaf)
                )

                if item_id:
                    discovered_tags.add(item_id)

            except Exception as exc:

                print(
                    f"[WARNING] Could not resolve item "
                    f"{leaf}: {exc}"
                )

        # Visit child branches recursively.
        for branch in branch_names:

            try:

                browser.MoveDown(branch)

                browse_level()

                browser.MoveUp()

            except Exception as exc:

                print(
                    f"[WARNING] Could not browse branch "
                    f"{branch}: {exc}"
                )

    browse_level()

    return sorted(discovered_tags), browser


def main():

    if sys.platform != "win32":

        print("[FAIL] Windows only")

        return 1

    parser = argparse.ArgumentParser(
        description="Dynamic remote Matrikon OPC DA reader"
    )

    parser.add_argument(
        "--host",
        default=DEFAULT_HOST
    )

    parser.add_argument(
        "--progid",
        default=DEFAULT_PROGID
    )

    parser.add_argument(
        "--user",
        default=DEFAULT_USER
    )

    parser.add_argument(
        "--domain",
        default=DEFAULT_DOMAIN
    )

    parser.add_argument(
        "--interval",
        type=float,
        default=DEFAULT_INTERVAL
    )

    parser.add_argument(
        "--password",
        default=None
    )

    parser.add_argument(
        "--filter",
        default=DEFAULT_FILTER,
        help="Optional tag substring filter (e.g. 'Random' or leave empty for all tags)"
    )

    args = parser.parse_args()

    if args.interval <= 0:
        print("[FAIL] Interval must be greater than zero")
        return 1

    host = args.host.strip()

    username = rf"{args.domain}\{args.user}"

    password = args.password

    if password is None:

        password = getpass.getpass(
            f"Enter password for {username}: "
        )

    print("=" * 75)
    print("DYNAMIC REMOTE MATRIKON OPC DA READER")
    print("=" * 75)

    print(f"PC1 IP        : {host}")
    print(f"Windows user  : {username}")
    print(f"OPC server    : {args.progid}")
    print(f"Read interval : {args.interval} seconds")

    is_localhost = host in ("127.0.0.1", "localhost", "::1")

    # STEP 1: IPC authentication (Skip if testing on localhost)

    if not is_localhost:
        if not authenticate_ipc(
            host,
            username,
            password
        ):
            return 2
    else:
        print("\n[1] Running on localhost - skipping IPC$ network authentication")

    # STEP 2: Import COM libraries

    try:

        import pythoncom
        import win32com.client

    except ImportError as exc:

        print("[FAIL] pywin32 is not installed")

        print_error(exc)

        return 3

    com_initialized = False

    opc_server = None
    opc_group = None
    browser = None

    discovered_tags = []
    opc_items = []

    try:

        # STEP 3: Initialize COM

        pythoncom.CoInitialize()

        com_initialized = True

        print("\n[3] Initializing COM security...")

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

            print("[OK] Authentication: PKT_INTEGRITY")
            print("[OK] Impersonation: IMPERSONATE")

        except pythoncom.com_error as exc:

            hr = exc.args[0] if exc.args else None

            if hr == -2147417831:

                print(
                    "[WARNING] COM security already initialized"
                )

            else:

                raise

       

        print("\n[4] Creating OPC Automation client...")

        opc_server = win32com.client.Dispatch(
            "OPC.Automation.1"
        )  

        print("[OK] OPC Automation object created")

        # STEP 5: Connect to remote OPC server

        print("\n[5] Connecting to remote Matrikon server...")

        opc_server.Connect(
            args.progid,
            host
        )

        print(
            "[SUCCESS] Remote OPC DA connection established"
        )

        # STEP 6: Browse all existing server items

        discovered_tags, browser = browse_opc_tags(
            opc_server
        )
        if args.filter:
            discovered_tags = [
                tag for tag in discovered_tags
                if args.filter.lower() in tag.lower()
            ]

        if not discovered_tags:

            print("\n[FAIL] No OPC items discovered")

            print(
                "Check whether the server exposes "
                "its items through OPC browsing."
            )

            return 4

        print(
            f"\n[OK] Discovered {len(discovered_tags)} OPC items"
        )

        print("\nDISCOVERED OPC ITEM IDs")
        print("-" * 65)

        for index, tag in enumerate(
            discovered_tags,
            start=1
        ):

            print(f"{index:03d}. {tag}")

        print("-" * 65)

        # STEP 7: Create client-side OPC group

        print("\n[7] Creating OPC group...")

        opc_groups = opc_server.OPCGroups

        opc_groups.DefaultGroupIsActive = True

        opc_groups.DefaultGroupUpdateRate = 1000

        opc_group = opc_groups.Add(
            "PythonDynamicRemoteReadGroup"
        )

        print("[OK] OPC group created")

        # STEP 8: Add all discovered items to client group

        print("\n[8] Adding discovered OPC items...")

        opc_items_collection = opc_group.OPCItems

        for index, tag in enumerate(
            discovered_tags,
            start=1
        ):

            try:

                item = opc_items_collection.AddItem(
                    tag,
                    index
                )

                opc_items.append(
                    (tag, item)
                )

                print(f"[ADDED] {tag}")

            except Exception as exc:

                print(f"[SKIPPED] {tag}: {exc}")

        if not opc_items:

            print("[FAIL] No OPC items could be added")

            return 5

        print(
            f"\n[OK] Successfully added "
            f"{len(opc_items)} OPC items"
        )

        # STEP 9: Continuously read all items

        print("\n" + "=" * 105)

        print("LIVE REMOTE OPC DA DATA")

        print("Press CTRL+C to stop")

        print("=" * 105)

        cycle = 0

        while True:

            cycle += 1

            print(
                f"\nREAD CYCLE: {cycle} | "
                f"PC2 TIME: {time.strftime('%Y-%m-%d %H:%M:%S')}"
            )

            print("-" * 105)

            print(
                f"{'TAG':<45}"
                f"{'VALUE':<18}"
                f"{'QUALITY':<12}"
                f"TIMESTAMP"
            )

            print("-" * 105)

            for tag, item in opc_items:

                try:

                    value, quality, timestamp = item.Read(2)

                    local_timestamp_str = convert_to_local_time(
                        timestamp
                    )

                    quality_number = int(quality)

                    quality_text = (
                        "GOOD"
                        if (quality_number & 0xC0) == 0xC0
                        else str(quality)
                    )

                    print(
                        f"{tag:<45}"
                        f"{str(value):<18}"
                        f"{quality_text:<12}"
                        f"{local_timestamp_str}"
                    )

                except Exception as exc:

                    print(
                        f"{tag:<45}"
                        f"READ ERROR: {exc}"
                    )

            print("-" * 105)

            time.sleep(args.interval)

    except KeyboardInterrupt:

        print("\n\n[STOPPED] User stopped the OPC reader")

    except Exception as exc:

        print("\n" + "=" * 75)

        print("[FAIL] REMOTE OPC DA READER FAILED")

        print("=" * 75)

        print_error(exc)

        return 1

    finally:

        print("\n[INFO] Cleaning up OPC connection...")

        if opc_server is not None:

            try:

                opc_server.Disconnect()

                print("[OK] OPC server disconnected")

            except Exception:
                pass

        opc_items.clear()

        opc_group = None
        browser = None
        opc_server = None

        if com_initialized:

            import gc

            gc.collect()

            pythoncom.CoUninitialize()

        print("[INFO] Client stopped")

    return 0


if __name__ == "__main__":

    raise SystemExit(main())
