import os
import re
import time
import shutil
import csv
from openpyxl import load_workbook

from logger import logger

WATCH_DIR = os.getenv('WATCH_DIR', '')
OUTPUT_1F_DIR = os.getenv('OUTPUT_1F_DIR', '')
OUTPUT_CSV_DIR = os.getenv('OUTPUT_CSV_DIR', '')
ARCHIVE_DIR = os.getenv('ARCHIVE_DIR', '')
POLL_INTERVAL = int(os.getenv('POLL_INTERVAL', 5))
MAX_ARCHIVE_FILE_AGE = int(os.getenv('MAX_ARCHIVE_FILE_AGE', 30))
TRIM_INTERVAL = int(os.getenv('TRIM_INTERVAL', 86400))


def replace_commas_in_parentheses(text):
    """Replace commas found inside parentheses with spaces.

    e.g. "Dimension 2d Distance (PNT19,PNT20)" ->
         "Dimension 2d Distance (PNT19 PNT20)"
    """
    if not isinstance(text, str):
        return text

    return re.sub(
        r"\(([^)]*)\)",
        lambda m: f"({m.group(1).replace(',', ' ')})",
        text,
    )


def clean_row(row):
    """Return a copy of row with commas inside parentheses removed from the
    second column."""
    row = list(row)
    if len(row) > 1:
        row[1] = replace_commas_in_parentheses(row[1])
    return row


def format_csv_filename(base_name):
    """Bracket the dash-delimited DOT machine name in a CSV base filename."""
    return re.sub(
        r"-(DOT\b[^-]*?\d)\s*-",
        r"-[\1] -",
        base_name,
    )


def trim_archive():
    if not ARCHIVE_DIR:
        return

    current_time = time.time()

    for filename in os.listdir(ARCHIVE_DIR):
        file_path = os.path.join(ARCHIVE_DIR, filename)
        if os.path.isfile(file_path):
            file_age = current_time - os.path.getmtime(file_path)
            if file_age > MAX_ARCHIVE_FILE_AGE * 86400:  # Convert days to seconds
                try:
                    os.remove(file_path)
                    logger.info(f"Deleted old archive file: {filename}")
                except Exception as e:
                    logger.error(f"Error deleting {filename}: {e}")


def convert_excel_to_csv():
    for folder in [WATCH_DIR, OUTPUT_1F_DIR, OUTPUT_CSV_DIR, ARCHIVE_DIR]:
        if folder and not os.path.exists(folder):
            os.makedirs(folder)

    for filename in os.listdir(WATCH_DIR):
        if filename.lower().endswith(('.xlsx', '.xls')) and not filename.startswith('~$'):
            excel_path = os.path.join(WATCH_DIR, filename)
            time.sleep(1)
            try:
                logger.info(f"Found file: {filename}. Processing...")
                base_name = os.path.splitext(filename)[0]
                csv_base_name = format_csv_filename(base_name)
                csv_path = os.path.join(OUTPUT_CSV_DIR, f"{csv_base_name}.csv")
                wb = load_workbook(excel_path, data_only=True)
                ws = wb.active

                with open(csv_path, "w", newline="", encoding="utf-8") as f:
                    writer = csv.writer(f)
                    for row in ws.iter_rows(values_only=True):
                        writer.writerow(clean_row(row))
                shutil.copy(excel_path, os.path.join(OUTPUT_1F_DIR, filename))
                shutil.move(excel_path, os.path.join(ARCHIVE_DIR, filename))

            except Exception as e:
                logger.error(f"Error processing {filename}: {e}")


if __name__ == "__main__":
    last_trim_time = 0
    while True:
        convert_excel_to_csv()

        if time.time() - last_trim_time >= TRIM_INTERVAL:
            trim_archive()
            last_trim_time = time.time()

        time.sleep(POLL_INTERVAL)
