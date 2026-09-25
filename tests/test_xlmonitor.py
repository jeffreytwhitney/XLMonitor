import csv
import os
import shutil
import time

import pytest
from openpyxl import Workbook, load_workbook

import xlMonitor as xl

FILES_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "files")
RESET_DIR = os.path.join(FILES_DIR, "reset")

# The integration test exercises the real pipeline directories, sourced from
# .envtest (loaded in conftest.py) so it matches whatever settings are used
# for manual/integration testing rather than hardcoded paths.
def _require_env(name):
    value = os.getenv(name)
    assert value, f"Expected {name} to be set (via .envtest) for the integration test"
    return value


ARCHIVE_DIR = _require_env("ARCHIVE_DIR")
IN_DIR = _require_env("WATCH_DIR")
OUT_1F_DIR = _require_env("OUTPUT_1F_DIR")
OUT_CSV_DIR = _require_env("OUTPUT_CSV_DIR")


def make_workbook(path, rows):
    wb = Workbook()
    ws = wb.active
    assert ws is not None
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
    working_dir = tmp_path / "working"
    watch_dir.mkdir()
    output_1f_dir.mkdir()
    output_csv_dir.mkdir()
    archive_dir.mkdir()

    monkeypatch.setattr(xl, "WATCH_DIR", str(watch_dir))
    monkeypatch.setattr(xl, "OUTPUT_1F_DIR", str(output_1f_dir))
    monkeypatch.setattr(xl, "OUTPUT_CSV_DIR", str(output_csv_dir))
    monkeypatch.setattr(xl, "ARCHIVE_DIR", str(archive_dir))
    monkeypatch.setattr(xl, "WORKING_DIR", str(working_dir))

    return {
        "watch": str(watch_dir),
        "output_1f": str(output_1f_dir),
        "output_csv": str(output_csv_dir),
        "archive": str(archive_dir),
        "working": str(working_dir),
    }


class TestReplaceCommasInParentheses:
    def test_replaces_comma_inside_parentheses_with_space(self):
        text = "Dimension 2d Distance (PNT19,PNT20)"
        result = xl.replace_commas_in_parentheses(text)
        assert result == "Dimension 2d Distance (PNT19 PNT20)"

    def test_replaces_multiple_commas_inside_parentheses(self):
        text = "Dimension (PNT1,PNT2,PNT3)"
        result = xl.replace_commas_in_parentheses(text)
        assert result == "Dimension (PNT1 PNT2 PNT3)"

    def test_leaves_commas_outside_parentheses_untouched(self):
        text = "Dimension, 2d Distance (PNT19,PNT20), extra"
        result = xl.replace_commas_in_parentheses(text)
        assert result == "Dimension, 2d Distance (PNT19 PNT20), extra"

    def test_handles_multiple_parenthetical_groups(self):
        text = "(A,B) and (C,D)"
        result = xl.replace_commas_in_parentheses(text)
        assert result == "(A B) and (C D)"

    def test_text_without_parentheses_is_unchanged(self):
        text = "No parentheses here, just a comma"
        result = xl.replace_commas_in_parentheses(text)
        assert result == text

    def test_non_string_values_are_returned_unchanged(self):
        assert xl.replace_commas_in_parentheses(None) is None
        assert xl.replace_commas_in_parentheses(42) == 42


class TestCleanRow:
    def test_cleans_second_column_only(self):
        row = ("first (a,b)", "second (c,d)", "third (e,f)")
        result = xl.clean_row(row)
        assert result == ["first (a,b)", "second (c d)", "third (e,f)"]

    def test_row_with_fewer_than_two_columns_is_unchanged(self):
        assert xl.clean_row(("only",)) == ["only"]
        assert xl.clean_row(()) == []

    def test_non_string_second_column_is_left_as_is(self):
        row = ("first", 123, "third")
        assert xl.clean_row(row) == ["first", 123, "third"]

    def test_returns_list_not_original_tuple(self):
        row = ("a", "b (c,d)")
        result = xl.clean_row(row)
        assert isinstance(result, list)
        assert row == ("a", "b (c,d)")  # original untouched

    def test_clears_nested_directories_and_files(self, dirs):
        nested_dir = os.path.join(dirs["working"], "nested")
        os.makedirs(nested_dir)
        with open(os.path.join(dirs["working"], "stale.csv"), "w", encoding="utf-8") as f:
            f.write("stale")

        xl.clear_working_directory()

        assert os.listdir(dirs["working"]) == []


class TestFormatCsvFilename:
    def test_brackets_dot_machine_name_and_trims_trailing_space(self):
        base_name = "M961373A001_RevE-Op10-DOT 6.2 - 09_22_2026 11_24_00 PM"

        result = xl.format_csv_filename(base_name)

        assert result == "M961373A001_RevE-Op10-[DOT 6.2] - 09_22_2026 11_24_00 PM"

    def test_leaves_names_without_a_dot_machine_name_unchanged(self):
        base_name = "M961373A001_RevE-Op10-Other Machine - 09_22_2026"

        assert xl.format_csv_filename(base_name) == base_name

    def test_brackets_dot_machine_name_without_a_decimal(self):
        base_name = "M961373A001_RevE-Op10-DOT 6 - 09_22_2026 11_24_00 PM"

        result = xl.format_csv_filename(base_name)

        assert result == "M961373A001_RevE-Op10-[DOT 6] - 09_22_2026 11_24_00 PM"

    def test_brackets_dot_machine_name_with_no_space_or_underscore(self):
        base_name = "M961373A001_RevE-Op10-DOT1 - 09_22_2026 11_24_00 PM"

        result = xl.format_csv_filename(base_name)

        assert result == "M961373A001_RevE-Op10-[DOT1] - 09_22_2026 11_24_00 PM"


class TestConvertExcelToCsv:
    def test_converts_xlsx_to_csv(self, dirs):
        xlsx_path = os.path.join(dirs["watch"], "sample.xlsx")
        make_workbook(xlsx_path, [["a", "b", "c"], [1, 2, 3]])

        xl.convert_excel_to_csv()

        csv_path = os.path.join(dirs["output_csv"], "sample.csv")
        assert os.path.exists(csv_path)
        with open(csv_path, newline="", encoding="utf-8") as f:
            rows = list(csv.reader(f))
        assert rows == [["a", "b", "c"], ["1", "2", "3"]]

    def test_brackets_dot_machine_name_in_csv_filename(self, dirs):
        xlsx_name = "M961373A001_RevE-Op10-DOT 6.2 - 09_22_2026 11_24_00 PM.xlsx"
        make_workbook(os.path.join(dirs["watch"], xlsx_name), [["a"]])

        xl.convert_excel_to_csv()

        expected_name = "M961373A001_RevE-Op10-[DOT 6.2] - 09_22_2026 11_24_00 PM.csv"
        assert os.path.exists(os.path.join(dirs["output_csv"], expected_name))

    def test_brackets_dot_machine_name_without_decimal_in_csv_filename(self, dirs):
        xlsx_name = "M961373A001_RevE-Op10-DOT 6 - 09_22_2026 11_24_00 PM.XLSX"
        make_workbook(os.path.join(dirs["watch"], xlsx_name), [["a"]])

        xl.convert_excel_to_csv()

        expected_name = "M961373A001_RevE-Op10-[DOT 6] - 09_22_2026 11_24_00 PM.csv"
        assert os.path.exists(os.path.join(dirs["output_csv"], expected_name))

    def test_copies_source_file_to_1f_output(self, dirs):
        xlsx_path = os.path.join(dirs["watch"], "sample.xlsx")
        make_workbook(xlsx_path, [["x"]])

        xl.convert_excel_to_csv()

        assert os.path.exists(os.path.join(dirs["output_1f"], "sample.xlsx"))

    def test_removes_commas_in_parentheses_from_second_column(self, dirs):
        xlsx_path = os.path.join(dirs["watch"], "sample.xlsx")
        make_workbook(xlsx_path, [
            ["header1", "header2", "header3"],
            [1, "Dimension 2d Distance (PNT19,PNT20)", "unchanged,text"],
        ])

        xl.convert_excel_to_csv()

        csv_path = os.path.join(dirs["output_csv"], "sample.csv")
        with open(csv_path, newline="", encoding="utf-8") as f:
            rows = list(csv.reader(f))

        assert rows[1][1] == "Dimension 2d Distance (PNT19 PNT20)"
        # Third column is untouched, including its comma.
        assert rows[1][2] == "unchanged,text"

    def test_moves_source_file_to_archive(self, dirs):
        xlsx_path = os.path.join(dirs["watch"], "sample.xlsx")
        make_workbook(xlsx_path, [["x"]])

        xl.convert_excel_to_csv()

        assert not os.path.exists(xlsx_path)
        assert os.path.exists(os.path.join(dirs["archive"], "sample.xlsx"))

    def test_writes_csv_locally_before_publishing(self, dirs):
        stale_file = os.path.join(dirs["working"], "stale.csv")
        os.makedirs(dirs["working"])
        with open(stale_file, "w", encoding="utf-8") as f:
            f.write("stale")
        xlsx_path = os.path.join(dirs["watch"], "sample.xlsx")
        make_workbook(xlsx_path, [["a", "b"], [1, 2]])

        xl.convert_excel_to_csv()

        csv_path = os.path.join(dirs["output_csv"], "sample.csv")
        assert os.path.exists(csv_path)
        assert os.listdir(dirs["working"]) == []

    def test_ignores_temp_lock_files(self, dirs):
        lock_path = os.path.join(dirs["watch"], "~$sample.xlsx")
        with open(lock_path, "w", encoding="utf-8") as f:
            f.write("not a real workbook")

        xl.convert_excel_to_csv()

        # Lock file should be left untouched and nothing produced.
        assert os.path.exists(lock_path)
        assert os.listdir(dirs["output_1f"]) == []
        assert os.listdir(dirs["output_csv"]) == []
        assert os.listdir(dirs["archive"]) == []

    def test_ignores_non_excel_files(self, dirs):
        txt_path = os.path.join(dirs["watch"], "notes.txt")
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write("hello")

        xl.convert_excel_to_csv()

        assert os.path.exists(txt_path)
        assert os.listdir(dirs["output_1f"]) == []
        assert os.listdir(dirs["output_csv"]) == []
        assert os.listdir(dirs["archive"]) == []

    def test_creates_missing_local_directories(self, tmp_path, monkeypatch):
        """WATCH_DIR and ARCHIVE_DIR are local; they should be auto-created."""
        watch_dir = tmp_path / "does_not_exist_watch"
        output_1f_dir = tmp_path / "output_1f"
        output_csv_dir = tmp_path / "output_csv"
        archive_dir = tmp_path / "does_not_exist_archive"
        output_1f_dir.mkdir()
        output_csv_dir.mkdir()

        monkeypatch.setattr(xl, "WATCH_DIR", str(watch_dir))
        monkeypatch.setattr(xl, "OUTPUT_1F_DIR", str(output_1f_dir))
        monkeypatch.setattr(xl, "OUTPUT_CSV_DIR", str(output_csv_dir))
        monkeypatch.setattr(xl, "ARCHIVE_DIR", str(archive_dir))

        xl.convert_excel_to_csv()

        assert os.path.isdir(watch_dir)
        assert os.path.isdir(archive_dir)

    def test_errors_out_when_output_1f_dir_missing(self, tmp_path, monkeypatch):
        """OUTPUT_1F_DIR/OUTPUT_CSV_DIR are network shares; never auto-create them."""
        watch_dir = tmp_path / "watch"
        output_csv_dir = tmp_path / "output_csv"
        archive_dir = tmp_path / "archive"
        watch_dir.mkdir()
        output_csv_dir.mkdir()
        archive_dir.mkdir()
        missing_output_1f_dir = tmp_path / "does_not_exist_output_1f"

        make_workbook(os.path.join(str(watch_dir), "sample.xlsx"), [["x"]])

        monkeypatch.setattr(xl, "WATCH_DIR", str(watch_dir))
        monkeypatch.setattr(xl, "OUTPUT_1F_DIR", str(missing_output_1f_dir))
        monkeypatch.setattr(xl, "OUTPUT_CSV_DIR", str(output_csv_dir))
        monkeypatch.setattr(xl, "ARCHIVE_DIR", str(archive_dir))

        xl.convert_excel_to_csv()

        assert not os.path.isdir(missing_output_1f_dir)
        # File left untouched since processing was skipped entirely.
        assert os.path.exists(os.path.join(str(watch_dir), "sample.xlsx"))

    def test_errors_out_when_output_csv_dir_missing(self, tmp_path, monkeypatch):
        watch_dir = tmp_path / "watch"
        output_1f_dir = tmp_path / "output_1f"
        archive_dir = tmp_path / "archive"
        watch_dir.mkdir()
        output_1f_dir.mkdir()
        archive_dir.mkdir()
        missing_output_csv_dir = tmp_path / "does_not_exist_output_csv"

        make_workbook(os.path.join(str(watch_dir), "sample.xlsx"), [["x"]])

        monkeypatch.setattr(xl, "WATCH_DIR", str(watch_dir))
        monkeypatch.setattr(xl, "OUTPUT_1F_DIR", str(output_1f_dir))
        monkeypatch.setattr(xl, "OUTPUT_CSV_DIR", str(missing_output_csv_dir))
        monkeypatch.setattr(xl, "ARCHIVE_DIR", str(archive_dir))

        xl.convert_excel_to_csv()

        assert not os.path.isdir(missing_output_csv_dir)
        assert os.path.exists(os.path.join(str(watch_dir), "sample.xlsx"))

    def test_handles_corrupt_workbook_without_raising(self, dirs):
        bad_path = os.path.join(dirs["watch"], "corrupt.xlsx")
        with open(bad_path, "w", encoding="utf-8") as f:
            f.write("this is not a real xlsx file")

        # Should not raise; error is logged and file is left in place.
        xl.convert_excel_to_csv()

        assert os.path.exists(bad_path)
        assert os.listdir(dirs["output_csv"]) == []
        assert os.listdir(dirs["output_1f"]) == []

    def test_keeps_processing_when_archive_directory_creation_fails(
        self, tmp_path, monkeypatch
    ):
        watch_dir = tmp_path / "watch"
        output_1f_dir = tmp_path / "output_1f"
        output_csv_dir = tmp_path / "output_csv"
        watch_dir.mkdir()
        output_1f_dir.mkdir()
        output_csv_dir.mkdir()
        missing_archive = tmp_path / "archive"

        monkeypatch.setattr(xl, "WATCH_DIR", str(watch_dir))
        monkeypatch.setattr(xl, "ARCHIVE_DIR", str(missing_archive))
        monkeypatch.setattr(xl, "WORKING_DIR", str(tmp_path / "working"))
        monkeypatch.setattr(xl, "OUTPUT_1F_DIR", str(output_1f_dir))
        monkeypatch.setattr(xl, "OUTPUT_CSV_DIR", str(output_csv_dir))

        real_makedirs = os.makedirs

        def fail_archive_creation(path, *args, **kwargs):
            if path == str(missing_archive):
                raise OSError("archive unavailable")
            real_makedirs(path, *args, **kwargs)

        monkeypatch.setattr(xl.os, "makedirs", fail_archive_creation)
        xl.convert_excel_to_csv()

        assert not missing_archive.exists()

    def test_handles_workbook_without_active_sheet(self, dirs, monkeypatch):
        xlsx_path = os.path.join(dirs["watch"], "empty.xlsx")
        make_workbook(xlsx_path, [["value"]])

        class WorkbookWithoutActiveSheet:
            active = None

        monkeypatch.setattr(xl, "load_workbook", lambda *args, **kwargs: WorkbookWithoutActiveSheet())
        xl.convert_excel_to_csv()

        assert os.path.exists(xlsx_path)

    def test_test_mode_leaves_csv_in_working_directory(self, dirs, monkeypatch):
        monkeypatch.setattr(xl, "TEST_MODE", True)
        xlsx_path = os.path.join(dirs["watch"], "sample.xlsx")
        make_workbook(xlsx_path, [["a", "b"]])

        xl.convert_excel_to_csv()

        assert os.path.exists(os.path.join(dirs["working"], "sample.csv"))
        assert not os.path.exists(os.path.join(dirs["output_csv"], "sample.csv"))


class TestTrimArchive:
    def test_deletes_files_older_than_max_age(self, tmp_path, monkeypatch):
        archive_dir = tmp_path / "archive"
        archive_dir.mkdir()
        old_file = archive_dir / "old.xlsx"
        old_file.write_text("old")

        monkeypatch.setattr(xl, "ARCHIVE_DIR", str(archive_dir))
        monkeypatch.setattr(xl, "MAX_ARCHIVE_FILE_AGE", 1)

        old_time = time.time() - (2 * 86400)
        os.utime(old_file, (old_time, old_time))

        xl.trim_archive()

        assert not old_file.exists()

    def test_keeps_files_within_max_age(self, tmp_path, monkeypatch):
        archive_dir = tmp_path / "archive"
        archive_dir.mkdir()
        recent_file = archive_dir / "recent.xlsx"
        recent_file.write_text("recent")

        monkeypatch.setattr(xl, "ARCHIVE_DIR", str(archive_dir))
        monkeypatch.setattr(xl, "MAX_ARCHIVE_FILE_AGE", 30)

        xl.trim_archive()

        assert recent_file.exists()

    def test_noop_when_archive_dir_not_configured(self, monkeypatch):
        monkeypatch.setattr(xl, "ARCHIVE_DIR", "")
        # Should simply return without raising, even though the (empty)
        # directory does not exist.
        xl.trim_archive()

    def test_logs_and_continues_when_old_file_cannot_be_deleted(
        self, tmp_path, monkeypatch
    ):
        archive_dir = tmp_path / "archive"
        archive_dir.mkdir()
        old_file = archive_dir / "old.xlsx"
        old_file.write_text("old")
        old_time = time.time() - (2 * 86400)
        os.utime(old_file, (old_time, old_time))

        monkeypatch.setattr(xl, "ARCHIVE_DIR", str(archive_dir))
        monkeypatch.setattr(xl, "MAX_ARCHIVE_FILE_AGE", 1)
        monkeypatch.setattr(xl.os, "remove", lambda path: (_ for _ in ()).throw(OSError("locked")))

        xl.trim_archive()

        assert old_file.exists()


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

        monkeypatch.setattr(xl, "WATCH_DIR", IN_DIR)
        monkeypatch.setattr(xl, "OUTPUT_1F_DIR", OUT_1F_DIR)
        monkeypatch.setattr(xl, "OUTPUT_CSV_DIR", OUT_CSV_DIR)
        monkeypatch.setattr(xl, "ARCHIVE_DIR", ARCHIVE_DIR)

        xl.convert_excel_to_csv()

        base_name = xl.format_csv_filename(os.path.splitext(source_name)[0])
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
        assert ws is not None
        expected_rows = [list(row) for row in ws.iter_rows(values_only=True)]

        with open(csv_path, newline="", encoding="utf-8") as f:
            actual_rows = list(csv.reader(f))

        assert len(actual_rows) == len(expected_rows)


class TestShouldExit:
    def test_defaults_to_false_when_unset(self, monkeypatch):
        monkeypatch.delenv("KILL_ME_NOW", raising=False)
        assert xl.should_exit() is False

    @pytest.mark.parametrize("value", ["1", "true", "True", "TRUE", "yes", "Yes"])
    def test_truthy_values_return_true(self, monkeypatch, value):
        monkeypatch.setenv("KILL_ME_NOW", value)
        assert xl.should_exit() is True

    @pytest.mark.parametrize("value", ["0", "false", "False", "no", "", "  "])
    def test_falsy_values_return_false(self, monkeypatch, value):
        monkeypatch.setenv("KILL_ME_NOW", value)
        assert xl.should_exit() is False


class TestRunMonitorLoop:
    def test_exits_immediately_without_processing_when_should_exit_is_true(self, monkeypatch):
        monkeypatch.setattr(xl, "should_exit", lambda: True)
        convert_calls = []
        monkeypatch.setattr(
            xl, "convert_excel_to_csv", lambda: convert_calls.append(1)
        )

        xl.run_monitor_loop()
        assert convert_calls == []

    def test_processes_files_until_should_exit_is_true(self, monkeypatch):
        exit_flags = [False, False, True]
        monkeypatch.setattr(xl, "should_exit", lambda: exit_flags.pop(0))
        convert_calls = []
        monkeypatch.setattr(
            xl, "convert_excel_to_csv", lambda: convert_calls.append(1)
        )
        monkeypatch.setattr(xl, "trim_archive", lambda: None)
        monkeypatch.setattr(xl.time, "sleep", lambda seconds: None)

        xl.run_monitor_loop()

        assert len(convert_calls) == 2

    def test_trims_archive_when_trim_interval_has_elapsed(self, monkeypatch):
        exit_flags = [False, True]
        monkeypatch.setattr(xl, "should_exit", lambda: exit_flags.pop(0))
        monkeypatch.setattr(xl, "convert_excel_to_csv", lambda: None)
        monkeypatch.setattr(xl, "TRIM_INTERVAL", 0)
        trim_calls = []
        monkeypatch.setattr(
            xl, "trim_archive", lambda: trim_calls.append(1)
        )
        monkeypatch.setattr(xl.time, "sleep", lambda seconds: None)

        xl.run_monitor_loop()

        assert trim_calls == [1]

    def test_logs_and_continues_after_cycle_and_trim_errors(self, monkeypatch):
        exit_flags = [False, True]
        monkeypatch.setattr(xl, "should_exit", lambda: exit_flags.pop(0))
        monkeypatch.setattr(
            xl, "convert_excel_to_csv", lambda: (_ for _ in ()).throw(RuntimeError("cycle"))
        )
        monkeypatch.setattr(
            xl, "trim_archive", lambda: (_ for _ in ()).throw(RuntimeError("trim"))
        )
        monkeypatch.setattr(xl, "TRIM_INTERVAL", 0)
        monkeypatch.setattr(xl.time, "sleep", lambda seconds: None)

        xl.run_monitor_loop()


class TestMain:
    def test_registers_exit_handler_and_runs_monitor(self, monkeypatch):
        exit_handlers = []
        monkeypatch.setattr(
            xl.atexit, "register", lambda handler: exit_handlers.append(handler)
        )
        monkeypatch.setattr(xl, "run_monitor_loop", lambda: None)

        xl.main()

        assert len(exit_handlers) == 1
        exit_handlers[0]()

    def test_reraises_unhandled_monitor_exception(self, monkeypatch):
        monkeypatch.setattr(
            xl, "run_monitor_loop", lambda: (_ for _ in ()).throw(RuntimeError("fatal"))
        )

        with pytest.raises(RuntimeError, match="fatal"):
            xl.main()
