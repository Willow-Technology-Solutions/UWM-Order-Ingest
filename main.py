"""
Dwelling Blocks Order Ingest Script

Pull vendor assignments from the Dwelling Blocks Vendor API and append new rows
to the ERP PDF Automation workbook, matching existing Excel table formatting.
"""

from __future__ import annotations

import argparse
import logging
import os
import sys
import threading
import time

from dotenv import load_dotenv

from helpers import (
    appendErpRows,
    backupAutomationWorkbook,
    cleanupOldFiles,
    getBaseDirectory,
    resolveAutomationWorkbookPath,
    sendSuccessHealthcheck,
)
from logger import setupLogger
from dwelling_blocks import (
    VIEW_FILTERS,
    DwellingBlocksClient,
    assignment_to_erp_row,
    enrich_assignment_with_detail,
    load_credentials,
    load_erp_columns,
)

load_dotenv(getBaseDirectory() / ".env")

logger = setupLogger()

try:
    import sentry_sdk
    from sentry_sdk.integrations.logging import LoggingIntegration
except ImportError:
    sentry_sdk = None
    LoggingIntegration = None

sentry_dsn = os.getenv("SENTRY_DSN")
if sentry_dsn and sentry_sdk is not None:
    sentry_sdk.init(
        dsn=sentry_dsn,
        integrations=[
            LoggingIntegration(
                level=logging.INFO,
                event_level=logging.ERROR,
            ),
        ],
        traces_sample_rate=1.0,
        profiles_sample_rate=1.0,
        send_default_pii=True,
        environment=os.getenv("SCRIPT_ENVIRONMENT", "production"),
    )
    logger.info("Sentry initialized successfully")
elif sentry_dsn:
    logger.warning("SENTRY_DSN is set but sentry_sdk is not installed")
else:
    logger.warning("SENTRY_DSN not set, Sentry error tracking is disabled")

SCRIPT_EXEC_INTERVAL_SECONDS = int(os.getenv("SCRIPT_EXEC_INTERVAL_SECONDS", "300"))
DEFAULT_PAGE_SIZE = int(os.getenv("DEFAULT_PAGE_SIZE", "100"))


def import_pywin32_modules():
    """Import and return the pywin32 service modules."""
    try:  # noqa: PLC0415
        import servicemanager
        import win32service
        import win32serviceutil
    except ImportError as error:
        raise RuntimeError(
            "pywin32 is required to use service commands. "
            "Install it with 'pip install pywin32' on Windows."
        ) from error

    return servicemanager, win32service, win32serviceutil


def log_service_info(message: str) -> None:
    """Write an informational message to the Windows service log when available."""
    try:
        servicemanager, _, _ = import_pywin32_modules()
        servicemanager.LogInfoMsg(message)
    except Exception:
        logger.info(message)


def log_service_error(message: str) -> None:
    """Write an error message to the Windows service log when available."""
    try:
        servicemanager, _, _ = import_pywin32_modules()
        servicemanager.LogErrorMsg(message)
    except Exception:
        logger.error(message)


def pull_erp_rows(
    *,
    view: str = "in-progress",
    page_size: int = DEFAULT_PAGE_SIZE,
    max_results: int | None = None,
) -> list[dict[str, str]]:
    """Pull Dwelling Blocks assignments and map them to ERP column rows."""
    columns = load_erp_columns()
    client_id, client_secret = load_credentials()
    client = DwellingBlocksClient(client_id, client_secret)

    logger.info("Authenticating to Dwelling Blocks...")
    client.get_access_token()
    logger.info("Access token acquired.")

    filters = VIEW_FILTERS[view]
    logger.info(
        "Pulling assignments via POST /vendors/orders/search "
        "(view=%s, page_size=%s)...",
        view,
        page_size,
    )
    orders = list(
        client.iter_vendor_orders(
            page_size=page_size,
            max_results=max_results,
            use_search=True,
            filters=filters,
        )
    )
    logger.info("Enriching %s assignment(s) with contacts/properties/transfers...", len(orders))
    enriched = [enrich_assignment_with_detail(client, item) for item in orders]
    return [assignment_to_erp_row(order, columns) for order in enriched]


def run_once(
    *,
    view: str = "in-progress",
    page_size: int = DEFAULT_PAGE_SIZE,
    max_results: int | None = None,
) -> None:
    """Run a single Dwelling Blocks ingest cycle."""
    logger.info("Starting Dwelling Blocks order ingest script...")

    for directory_name in ("logs", "archive", "exports"):
        directory_path = os.path.join(getBaseDirectory(), directory_name)
        os.makedirs(directory_path, exist_ok=True)
        logger.info(f"Ensured directory exists: {directory_path}")

    destination_file_path = resolveAutomationWorkbookPath()
    logger.info(f"Resolved automation workbook path: {destination_file_path}")
    backup_destination_file = backupAutomationWorkbook(destination_file_path)

    erp_rows = pull_erp_rows(
        view=view,
        page_size=page_size,
        max_results=max_results,
    )
    logger.info("Mapped %s ERP row(s) from Dwelling Blocks", len(erp_rows))

    new_orders_count, added_loan_numbers = appendErpRows(
        erp_rows,
        destination_file_path,
    )
    if added_loan_numbers:
        logger.info(
            "Added %s new orders to destination workbook for loan numbers: %s",
            new_orders_count,
            ", ".join(str(loan_number) for loan_number in added_loan_numbers),
        )
    else:
        logger.info("Added 0 new orders to destination workbook")

    if (
        new_orders_count == 0
        and backup_destination_file
        and os.path.exists(backup_destination_file)
    ):
        os.remove(backup_destination_file)
        logger.info("Removed backup file: %s", backup_destination_file)

    cleanupOldFiles()
    sendSuccessHealthcheck()


def run_main_loop(
    stop_event: threading.Event | None = None,
    *,
    view: str = "in-progress",
    page_size: int = DEFAULT_PAGE_SIZE,
) -> None:
    """Run the ingest cycle repeatedly using the configured interval."""
    interval_seconds = max(1, SCRIPT_EXEC_INTERVAL_SECONDS)

    while not (stop_event and stop_event.is_set()):
        cycle_start = time.monotonic()

        try:
            run_once(view=view, page_size=page_size)
        except Exception:
            logger.exception("Dwelling Blocks ingest cycle failed")

        elapsed_seconds = time.monotonic() - cycle_start
        remaining_seconds = max(0, interval_seconds - elapsed_seconds)

        while remaining_seconds > 0 and not (stop_event and stop_event.is_set()):
            wait_seconds = min(5, remaining_seconds)
            if stop_event and stop_event.wait(wait_seconds):
                return
            if not stop_event:
                time.sleep(wait_seconds)
            remaining_seconds -= wait_seconds


def build_python_service_class():
    """Create and return the Windows service wrapper class."""
    _, win32service, win32serviceutil = import_pywin32_modules()

    class PythonService(win32serviceutil.ServiceFramework):
        _svc_name_ = "DEADwellingBlocksOrderIngest"
        _svc_display_name_ = "DEA Dwelling Blocks Order Ingest"
        _svc_description_ = (
            "Pulls Dwelling Blocks assignments into ERP PDF Automation.xlsx "
            "on a recurring interval."
        )

        def __init__(self, args):
            super().__init__(args)
            self.stop_event = threading.Event()

        def SvcStop(self):
            self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
            self.stop_event.set()

        def SvcDoRun(self):
            try:
                self.ReportServiceStatus(win32service.SERVICE_RUNNING)
                log_service_info("DEA Dwelling Blocks Order Ingest service started.")
                run_main_loop(stop_event=self.stop_event)
            except Exception:
                log_service_error("DEA Dwelling Blocks Order Ingest service runtime failed.")
                logger.exception("Service runtime failed")
                raise
            finally:
                log_service_info("DEA Dwelling Blocks Order Ingest service stopped.")

    return PythonService


def run_service_command(service_command: str) -> int:
    """Execute one Windows service management command."""
    _, _, win32serviceutil = import_pywin32_modules()
    python_service = build_python_service_class()
    win32serviceutil.HandleCommandLine(python_service, argv=[sys.argv[0], service_command])
    return 0


def parse_args(argv: list[str]):
    """Parse CLI options and optional Windows service management commands."""
    parser = argparse.ArgumentParser(
        description="DEA Dwelling Blocks Order Ingest host."
    )
    parser.add_argument(
        "--once",
        action="store_true",
        help="Run a single ingest cycle and exit.",
    )
    parser.add_argument(
        "--view",
        choices=sorted(VIEW_FILTERS),
        default="in-progress",
        help="Assignments view to pull (default: in-progress)",
    )
    parser.add_argument(
        "--page-size",
        type=int,
        default=DEFAULT_PAGE_SIZE,
        help=f"numResults per page (default {DEFAULT_PAGE_SIZE})",
    )
    parser.add_argument(
        "--max",
        type=int,
        default=None,
        dest="max_results",
        help="Stop after this many assignments (default: all matching)",
    )
    parser.add_argument(
        "--print-automation-workbook-path",
        action="store_true",
        help="Resolve the automation workbook path, print it, and exit.",
    )
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument(
        "--service",
        dest="service_command",
        metavar="COMMAND",
        choices=["install", "start", "stop", "remove", "restart"],
        help=(
            "Manage the Windows service instead of running the app normally. "
            "Command choices are install, start, stop, remove, and restart."
        ),
    )
    return parser.parse_args(argv)


def main() -> None:
    """Run normally or dispatch a Windows service command."""
    args = parse_args(sys.argv[1:])

    if args.print_automation_workbook_path:
        print(resolveAutomationWorkbookPath())
        return

    if args.service_command is not None:
        run_service_command(args.service_command)
        return

    if args.once:
        run_once(
            view=args.view,
            page_size=args.page_size,
            max_results=args.max_results,
        )
        return

    run_main_loop(view=args.view, page_size=args.page_size)


if __name__ == "__main__":
    if os.name == "nt" and len(sys.argv) == 1:
        try:
            servicemanager, _, _ = import_pywin32_modules()
            servicemanager.Initialize()
            servicemanager.PrepareToHostSingle(build_python_service_class())
            servicemanager.StartServiceCtrlDispatcher()
        except Exception:
            main()
    else:
        main()
