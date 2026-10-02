# Dwelling Blocks order ingest

This script pulls vendor assignments from the Dwelling Blocks Vendor API and appends new rows to the ERP PDF Automation workbook. It skips loan numbers that already exist for the configured company in GSG Connect.

## Setup

Python 3.10 or newer is required. From this directory:

```bash
uv sync
cp .env.example .env
```

Fill in `.env` using the variables in [Configuration](#configuration). Place `ERP PDF Automation.xlsx` in this directory, or set `ERP_PDF_AUTOMATION_XLSX_PATH` to its full path.

`main.py` loads `.env` on import and exits if the file is missing. That includes `--help`.

## Windows build and release

GitHub Actions builds a Windows `--onedir` executable on pushes to `main`, pull requests, and manual runs. Download the `UWMOrderIngest-Windows` artifact and extract `UWMOrderIngest-Windows.zip`. Keep the entire `UWMOrderIngest` folder together. A `v*` tag also publishes that zip as a GitHub release.

Before running the executable, copy `.env.example` to `UWMOrderIngest/.env` and fill in the runtime settings. Place `ERP PDF Automation.xlsx` beside `UWMOrderIngest.exe` or set `ERP_PDF_AUTOMATION_XLSX_PATH` to its path. Neither the workbook nor `.env` is included in GitHub builds or releases. GitHub Actions does not need the Dwelling Blocks or GSG Connect credentials to package the app; keep them on the machine that runs it.

Run `UWMOrderIngest.exe --once` for one cycle, or use the service commands below with `UWMOrderIngest.exe` in place of `python main.py`.

## Run one cycle

```bash
uv run python main.py --once
```

The command authenticates, pulls the in-progress view, writes new rows into the workbook, and exits. It saves a timestamped copy of the workbook under `archive/` before the write. If no new rows are added, it deletes that backup.

Stdout and `logs/log.txt` both receive the log.

Limit the pull while testing:

```bash
uv run python main.py --once --max 10
```

Flags:

- `--view in-progress` pulls accepted assignments with no report uploaded, or with an open revision due. This is the default.
- `--view all` pulls every assignment the search endpoint returns.
- `--page-size 100` sets how many assignments are requested per page. The default comes from `DEFAULT_PAGE_SIZE`.
- `--max 10` stops after that many assignments.
- `--print-automation-workbook-path` prints the workbook path and exits.

## Run on a loop

```bash
uv run python main.py
```

The loop repeats until you stop the process. It waits `SCRIPT_EXEC_INTERVAL_SECONDS` between cycle starts. The default wait is 300 seconds, and time already spent in a cycle counts toward that wait.

On Windows, starting `main.py` with no arguments tries to run as the installed service. Use `--once` or `--service` there.

## Windows service

Install `pywin32`, then:

```bash
uv run python main.py --service install
uv run python main.py --service start
uv run python main.py --service stop
uv run python main.py --service restart
uv run python main.py --service remove
```

The service name is `DEADwellingBlocksOrderIngest`.

## Configuration

`DWELLING_BLOCKS_CLIENT_ID` and `DWELLING_BLOCKS_CLIENT_SECRET` authenticate to `https://api.dwellingblocks.com`.

`ERP_PDF_AUTOMATION_XLSX_PATH` is the workbook that receives new rows. If that path is unset or the file is missing, the script looks for `ERP PDF Automation.xlsx` in this directory.

`GSG_CONNECT_BASE_URL`, `GSG_CONNECT_USERNAME`, `GSG_CONNECT_PASSWORD`, and `COMPANY_ID` identify past orders. The script uses them to skip loan numbers already stored for that company. If any of them is missing, it pulls assignments and appends nothing.

`SCRIPT_EXEC_INTERVAL_SECONDS` is the wait between loop cycles. The default is 300.

`DEFAULT_PAGE_SIZE` is the number of assignments requested per API page. The default is 100.

`FILE_CLEANUP_DAYS_THRESHOLD` is how many days files in `archive/` and `exports/` are kept. The default is 7.

`HTTP_TIMEOUT` is the timeout in seconds for GSG Connect requests. The default is 60.

`HEALTHCHECK_IO_URL` receives a GET after a successful cycle. Leave it blank to skip the ping.

`SENTRY_DSN` turns on Sentry when `sentry-sdk` is installed. Leave it blank to skip error reporting. `SCRIPT_ENVIRONMENT` is the Sentry environment name. The default is production.
