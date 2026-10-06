# UWM Order Ingest

Automated tool to ingest vendor assignments from the Dwelling Blocks Vendor API into the ERP PDF Automation workbook.

## What it does

- Authenticates to the Dwelling Blocks Vendor API using credentials from the `.env` file.
- Pulls in-progress assignments and skips order dates older than the configured lookback.
- Creates a timestamped backup of `ERP PDF Automation.xlsx`.
- Appends new rows to `ERP PDF Automation.xlsx`, skipping loan numbers already stored for the company in GSG Connect.

## Requirements

- Python 3.10 or newer
- UV package manager

## Setup

1. Create a `.env` file in the project root:

   ```
   DWELLING_BLOCKS_CLIENT_ID=your-client-id
   DWELLING_BLOCKS_CLIENT_SECRET=your-client-secret

   ERP_PDF_AUTOMATION_XLSX_PATH=

   GSG_CONNECT_BASE_URL=...
   GSG_CONNECT_USERNAME=...
   GSG_CONNECT_PASSWORD=...
   COMPANY_ID=...

   LOGGING_LOCATION=
   SENTRY_DSN=optional-dsn
   SCRIPT_ENVIRONMENT=production
   SCRIPT_EXEC_INTERVAL_SECONDS=300

   HEALTHCHECK_IO_URL=
   DEFAULT_PAGE_SIZE=100
   FILE_CLEANUP_DAYS_THRESHOLD=7
   HTTP_TIMEOUT=60
   DNS_SERVER=
   CURRENT_ORDERS_ORDER_DATE_MONTH_LOOKBACK=6
   ```

2. Install dependencies:

   ```bash
   uv sync
   ```

3. Place `ERP PDF Automation.xlsx` in this directory, or set `ERP_PDF_AUTOMATION_XLSX_PATH` to its full path.

Leave `LOGGING_LOCATION` blank to write `log.txt` under `logs/` in this directory. `CURRENT_ORDERS_ORDER_DATE_MONTH_LOOKBACK` is how many months of order dates to keep. The default is 6.

If `GSG_CONNECT_BASE_URL`, `GSG_CONNECT_USERNAME`, `GSG_CONNECT_PASSWORD`, or `COMPANY_ID` is empty, the script still pulls assignments and appends no rows.

## How to use

Run the script on a loop:

```bash
uv run python main.py
```

The loop waits `SCRIPT_EXEC_INTERVAL_SECONDS` between cycle starts. The default is 300 seconds. Time spent in a cycle counts toward that wait.

Run one cycle and exit:

```bash
uv run python main.py --once
```

`--max 10` stops after that many assignments. `--view all` pulls every assignment the search returns. `--view in-progress` is the default. `--print-automation-workbook-path` prints the workbook path and exits.

## Windows service

On Windows, install pywin32, then:

```bash
uv run python main.py --service install
uv run python main.py --service start
uv run python main.py --service stop
uv run python main.py --service restart
uv run python main.py --service remove
```

The service name is `UWMOrderIngest`. Starting `main.py` with no arguments on Windows tries to run as the installed service. Use `--once` or `--service` there.

## Building the executable

```bash
uv run --group build pyinstaller --onedir --name "UWMOrderIngest" main.py
```
