# XLMonitor

XLMonitor watches a folder for Excel workbooks and converts each `.xlsx` file
to CSV. It writes the CSV to a configured output folder, copies the original
workbook to a second output folder, then moves the input workbook to an
archive. The second column of each CSV row has commas inside parentheses
replaced with spaces; other values are left unchanged.

## Requirements

- Python
- `openpyxl`
- `python-dotenv`

## Setup

Install the runtime dependencies:

```powershell
python -m pip install openpyxl python-dotenv
```

Create a `.env` file next to `XLMonitor.py` (or next to `XLMonitor.exe` when
running a packaged build):

```dotenv
WATCH_DIR=C:\XLMonitor\input
OUTPUT_1F_DIR=C:\XLMonitor\output\1f
OUTPUT_CSV_DIR=C:\XLMonitor\output\csv
ARCHIVE_DIR=C:\XLMonitor\archive
```

`WATCH_DIR` and `ARCHIVE_DIR` are created automatically if they do not exist.
`OUTPUT_1F_DIR` and `OUTPUT_CSV_DIR` must already exist and be accessible to
the process (for example, a mounted network share). Configure these paths for
your environment before starting the monitor.

## Run

Start the monitor from the project directory:

```powershell
python XLMonitor.py
```

The process checks the watch folder every five seconds by default. Files are
processed in place, so make sure input workbooks are no longer being edited
when they are picked up. Temporary Excel lock files and non-Excel files are
ignored. The application uses the workbook's active sheet and cached cell
values when creating the CSV.

To stop cleanly, set `KILL_ME_NOW=true` in the `.env` file. The monitor checks
this setting during its polling loop and exits after the next check.

## Configuration

All settings are optional unless noted. Defaults are used when values are
omitted or invalid.

| Variable | Default | Description |
| --- | --- | --- |
| `WATCH_DIR` | *(empty)* | Folder to scan for `.xlsx` workbooks. |
| `OUTPUT_1F_DIR` | *(empty)* | Existing folder where a copy of each source workbook is published. |
| `OUTPUT_CSV_DIR` | *(empty)* | Existing folder where each converted CSV is published. |
| `ARCHIVE_DIR` | *(empty)* | Folder where processed source workbooks are moved. |
| `POLL_INTERVAL` | `5` | Seconds between scans. |
| `MAX_ARCHIVE_FILE_AGE` | `30` | Days to retain files in the archive. |
| `TRIM_INTERVAL` | `86400` | Seconds between archive cleanup passes. |
| `TEST_MODE` | `false` | Write CSVs to the local `working` folder instead of publishing outputs. |
| `KILL_ME_NOW` | `false` | Set to `true` to request a graceful shutdown. |
| `LOGGER_LEVEL` | `INFO` | Python logging level. |
| `LOGGER_FILE` | `app.log` | Log filename, created next to the script or executable. |

In `TEST_MODE`, processed workbooks are still moved to the archive, but output
copies are not published. The CSV remains in the local `working` folder until
the next workbook is processed.

## Tests

Install the test dependency and run the suite:

```powershell
python -m pip install pytest
python -m pytest
```
