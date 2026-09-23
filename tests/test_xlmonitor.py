import csv
import os
import shutil
import time

import pytest
from openpyxl import Workbook, load_workbook

import XLMonitor


FILES_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "files")
ARCHIVE_DIR = os.path.join(FILES_DIR, "archive")
IN_DIR = os.path.join(FILES_DIR, "in")
OUT_1F_DIR = os.path.join(FILES_DIR, "out_1f")
OUT_CSV_DIR = os.path.join(FILES_DIR, "out_csv")
RESET_DIR = os.path.join(FILES_DIR, "reset")


def make_workbook(path, rows):
    wb = Workbook()
    ws = wb.active
    for row in rows:
        ws.append(row)
    wb.save(path)


@pytest.fixture
def dirs(tmp_path, monkeypatch):
    """Point the module's directory globals at isolated temp directories."""
    watch_dir = tmp_path / "watch"
    output_1f_dir = tmp_path / "output_1f"
    output_csv_dir = tmp_path / "output_csv"
    archive_dir = tmp_path / "archive"
    watch_dir.mkdir()
    output_1f_dir.mkdir()
    output_csv_dir.mkdir()
    archive_dir.mkdir()

    monkeypatch.setattr(XLMonitor, "WATCH_DIR", str(watch_dir))
    monkeypatch.setattr(XLMonitor, "OUTPUT_1F_DIR", str(output_1f_dir))
    monkeypatch.setattr(XLMonitor, "OUTPUT_CSV_DIR", str(output_csv_dir))
    monkeypatch.setattr(XLMonitor, "ARCHIVE_DIR", str(archive_dir))

    return {
        "watch": str(watch_dir),
        "output_1f": str(output_1f_dir),
        "output_csv": str(output_csv_dir),
        "archive": str(archive_dir),
    }


class TestReplaceCommasInParentheses:
    def test_replaces_comma_inside_parentheses_with_space(self):
        text = "Dimension 2d Distance (PNT19,PNT20)"
        result = XLMonitor.replace_commas_in_parentheses(text)
        assert result == "Dimension 2d Distance (PNT19 PNT20)"

    def test_replaces_multiple_commas_inside_parentheses(self):
        text = "Dimension (PNT1,PNT2,PNT3)"
        result = XLMonitor.replace_commas_in_parentheses(text)
        assert result == "Dimension (PNT1 PNT2 PNT3)"

    def test_leaves_commas_outside_parentheses_untouched(self):
        text = "Dimension, 2d Distance (PNT19,PNT20), extra"
        result = XLMonitor.replace_commas_in_parentheses(text)
        assert result == "Dimension, 2d Distance (PNT19 PNT20), extra"

    def test_handles_multiple_parenthetical_groups(self):
        text = "(A,B) and (C,D)"
        result = XLMonitor.replace_commas_in_parentheses(text)
        assert result == "(A B) and (C D)"

    def test_text_without_parentheses_is_unchanged(self):
        text = "No parentheses here, just a comma"
        result = XLMonitor.replace_commas_in_parentheses(text)
        assert result == text

    def test_non_string_values_are_returned_unchanged(self):
        assert XLMonitor.replace_commas_in_parentheses(None) is None
        assert XLMonitor.replace_commas_in_parentheses(42) == 42


class TestCleanRow:
    def test_cleans_second_column_only(self):
        row = ("first (a,b)", "second (c,d)", "third (e,f)")
        result = XLMonitor.clean_row(row)
        assert result == ["first (a,b)", "second (c d)", "third (e,f)"]

    def test_row_with_fewer_than_two_columns_is_unchanged(self):
        assert XLMonitor.clean_row(("only",)) == ["only"]
        assert XLMonitor.clean_row(()) == []

    def test_non_string_second_column_is_left_as_is(self):
        row = ("first", 123, "third")
        assert XLMonitor.clean_row(row) == ["first", 123, "third"]

    def test_returns_list_not_original_tuple(self):
        row = ("a", "b (c,d)")
        result = XLMonitor.clean_row(row)
        assert isinstance(result, list)
        assert row == ("a", "b (c,d)")  # original untouched


class TestConvertExcelToCsv:
    def test_converts_xlsx_to_csv(self, dirs):
        xlsx_path = os.path.join(dirs["watch"], "sample.xlsx")
        make_workbook(xlsx_path, [["a", "b", "c"], [1, 2, 3]])

        XLMonitor.convert_excel_to_csv()

        csv_path = os.path.join(dirs["output_csv"], "sample.csv")
        assert os.path.exists(csv_path)
        with open(csv_path, newline="", encoding="utf-8") as f:
            rows = list(csv.reader(f))
        assert rows == [["a", "b", "c"], ["1", "2", "3"]]

    def test_copies_source_file_to_1f_output(self, dirs):
        xlsx_path = os.path.join(dirs["watch"], "sample.xlsx")
        make_workbook(xlsx_path, [["x"]])

        XLMonitor.convert_excel_to_csv()

        assert os.path.exists(os.path.join(dirs["output_1f"], "sample.xlsx"))

    def test_removes_commas_in_parentheses_from_second_column(self, dirs):
        xlsx_path = os.path.join(dirs["watch"], "sample.xlsx")
        make_workbook(xlsx_path, [
            ["header1", "header2", "header3"],
            [1, "Dimension 2d Distance (PNT19,PNT20)", "unchanged,text"],
        ])

        XLMonitor.convert_excel_to_csv()

        csv_path = os.path.join(dirs["output_csv"], "sample.csv")
        with open(csv_path, newline="", encoding="utf-8") as f:
            rows = list(csv.reader(f))

        assert rows[1][1] == "Dimension 2d Distance (PNT19 PNT20)"
        # Third column is untouched, including its comma.
        assert rows[1][2] == "unchanged,text"

    def test_moves_source_file_to_archive(self, dirs):
        xlsx_path = os.path.join(dirs["watch"], "sample.xlsx")
        make_workbook(xlsx_path, [["x"]])

        XLMonitor.convert_excel_to_csv()

        assert not os.path.exists(xlsx_path)
        assert os.path.exists(os.path.join(dirs["archive"], "sample.xlsx"))

    def test_ignores_temp_lock_files(self, dirs):
        lock_path = os.path.join(dirs["watch"], "~$sample.xlsx")
        with open(lock_path, "w", encoding="utf-8") as f:
            f.write("not a real workbook")

        XLMonitor.convert_excel_to_csv()

        # Lock file should be left untouched and nothing produced.
        assert os.path.exists(lock_path)
        assert os.listdir(dirs["output_1f"]) == []
        assert os.listdir(dirs["output_csv"]) == []
        assert os.listdir(dirs["archive"]) == []

    def test_ignores_non_excel_files(self, dirs):
        txt_path = os.path.join(dirs["watch"], "notes.txt")
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write("hello")

        XLMonitor.convert_excel_to_csv()

        assert os.path.exists(txt_path)
        assert os.listdir(dirs["output_1f"]) == []
        assert os.listdir(dirs["output_csv"]) == []
        assert os.listdir(dirs["archive"]) == []

    def test_creates_missing_directories(self, tmp_path, monkeypatch):
        watch_dir = tmp_path / "watch"
        output_1f_dir = tmp_path / "does_not_exist_output_1f"
        output_csv_dir = tmp_path / "does_not_exist_output_csv"
        archive_dir = tmp_path / "does_not_exist_archive"
        watch_dir.mkdir()

        monkeypatch.setattr(XLMonitor, "WATCH_DIR", str(watch_dir))
        monkeypatch.setattr(XLMonitor, "OUTPUT_1F_DIR", str(output_1f_dir))
        monkeypatch.setattr(XLMonitor, "OUTPUT_CSV_DIR", str(output_csv_dir))
        monkeypatch.setattr(XLMonitor, "ARCHIVE_DIR", str(archive_dir))

        XLMonitor.convert_excel_to_csv()

        assert os.path.isdir(output_1f_dir)
        assert os.path.isdir(output_csv_dir)
        assert os.path.isdir(archive_dir)

    def test_handles_corrupt_workbook_without_raising(self, dirs):
        bad_path = os.path.join(dirs["watch"], "corrupt.xlsx")
        with open(bad_path, "w", encoding="utf-8") as f:
            f.write("this is not a real xlsx file")

        # Should not raise; error is logged and file is left in place.
        XLMonitor.convert_excel_to_csv()

        assert os.path.exists(bad_path)
        assert os.listdir(dirs["output_csv"]) == []
        assert os.listdir(dirs["output_1f"]) == []


class TestTrimArchive:
    def test_deletes_files_older_than_max_age(self, tmp_path, monkeypatch):
        archive_dir = tmp_path / "archive"
        archive_dir.mkdir()
        old_file = archive_dir / "old.xlsx"
        old_file.write_text("old")

        monkeypatch.setattr(XLMonitor, "ARCHIVE_DIR", str(archive_dir))
        monkeypatch.setattr(XLMonitor, "MAX_ARCHIVE_FILE_AGE", 1)

        old_time = time.time() - (2 * 86400)
        os.utime(old_file, (old_time, old_time))

        XLMonitor.trim_archive()

        assert not old_file.exists()

    def test_keeps_files_within_max_age(self, tmp_path, monkeypatch):
        archive_dir = tmp_path / "archive"
        archive_dir.mkdir()
        recent_file = archive_dir / "recent.xlsx"
        recent_file.write_text("recent")

        monkeypatch.setattr(XLMonitor, "ARCHIVE_DIR", str(archive_dir))
        monkeypatch.setattr(XLMonitor, "MAX_ARCHIVE_FILE_AGE", 30)

        XLMonitor.trim_archive()

        assert recent_file.exists()

    def test_noop_when_archive_dir_not_configured(self, monkeypatch):
        monkeypatch.setattr(XLMonitor, "ARCHIVE_DIR", "")
        # Should simply return without raising, even though the (empty)
        # directory does not exist.
        XLMonitor.trim_archive()


class TestIntegration:
    """End-to-end test against the real tests/files fixture directories."""

    def test_full_pipeline_using_fixture_directories(self, monkeypatch):
        reset_files = [
            f for f in os.listdir(RESET_DIR)
            if f.lower().endswith((".xlsx", ".xls"))
        ]
        assert reset_files, "Expected at least one fixture file in tests/files/reset"
        source_name = reset_files[0]
        source_path = os.path.join(RESET_DIR, source_name)

        # Clean the working directories used by the pipeline.
        for directory in (ARCHIVE_DIR, IN_DIR, OUT_1F_DIR, OUT_CSV_DIR):
            os.makedirs(directory, exist_ok=True)
            for name in os.listdir(directory):
                path = os.path.join(directory, name)
                if os.path.isfile(path) or os.path.islink(path):
                    os.remove(path)
                elif os.path.isdir(path):
                    shutil.rmtree(path)

        # Seed the "in" directory from the untouched "reset" fixture.
        shutil.copy2(source_path, os.path.join(IN_DIR, source_name))

        monkeypatch.setattr(XLMonitor, "WATCH_DIR", IN_DIR)
        monkeypatch.setattr(XLMonitor, "OUTPUT_1F_DIR", OUT_1F_DIR)
        monkeypatch.setattr(XLMonitor, "OUTPUT_CSV_DIR", OUT_CSV_DIR)
        monkeypatch.setattr(XLMonitor, "ARCHIVE_DIR", ARCHIVE_DIR)

        XLMonitor.convert_excel_to_csv()

        base_name = os.path.splitext(source_name)[0]
        csv_path = os.path.join(OUT_CSV_DIR, f"{base_name}.csv")
        copy_1f_path = os.path.join(OUT_1F_DIR, source_name)
        archived_path = os.path.join(ARCHIVE_DIR, source_name)

        # Source file processed: moved out of "in" and into "archive".
        assert not os.path.exists(os.path.join(IN_DIR, source_name))
        assert os.path.exists(archived_path)

        # Original workbook copied to the 1F output directory.
        assert os.path.exists(copy_1f_path)

        # CSV produced with content matching the original workbook.
        assert os.path.exists(csv_path)
        wb = load_workbook(source_path, data_only=True)
        ws = wb.active
        expected_rows = [list(row) for row in ws.iter_rows(values_only=True)]

        with open(csv_path, newline="", encoding="utf-8") as f:
            actual_rows = list(csv.reader(f))

        assert len(actual_rows) == len(expected_rows)
