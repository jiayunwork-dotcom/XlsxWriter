###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

import os
import tempfile
import unittest
import warnings
from xml.etree import ElementTree

from xlsxwriter.exceptions import OverlappingRange
from xlsxwriter.workbook import Workbook

MAIN_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


###############################################################################
#
# Independent XML parsing helpers. These deliberately don't use the XlsxWriter
# XML helpers so that the comparison is against a separate parser.
#
###############################################################################
def _read_sheet(path):
    import zipfile

    with zipfile.ZipFile(path) as zip_file:
        return zip_file.read("xl/worksheets/sheet1.xml").decode("utf-8")


def _read_shared_strings(path):
    import zipfile

    with zipfile.ZipFile(path) as zip_file:
        try:
            xml = zip_file.read("xl/sharedStrings.xml").decode("utf-8")
        except KeyError:
            return []

    root = ElementTree.fromstring(xml)
    strings = []
    for si in root.iter(MAIN_NS + "si"):
        strings.append("".join(node.text or "" for node in si.iter(MAIN_NS + "t")))

    return strings


def _cell_value(cell, shared_strings):
    # Return a storage independent representation of a cell value.
    cell_type = cell.get("t")
    formula = cell.find(MAIN_NS + "f")
    value = cell.find(MAIN_NS + "v")
    inline = cell.find(MAIN_NS + "is")

    if inline is not None:
        return ("str", "".join(node.text or "" for node in inline.iter(MAIN_NS + "t")))

    if formula is not None:
        return ("f", formula.text, value.text if value is not None else None)

    if cell_type == "s" and value is not None:
        return ("str", shared_strings[int(value.text)])

    if value is not None:
        return ("num", value.text)

    if cell_type == "inlineStr":
        return ("str", "")

    return ("blank",)


def _cell_map(path):
    # Return {(row, col_ref): (style, value)} for the whole worksheet.
    shared_strings = _read_shared_strings(path)
    root = ElementTree.fromstring(_read_sheet(path))

    cells = {}
    for row in root.iter(MAIN_NS + "row"):
        for cell in row:
            if not cell.tag.endswith("c"):
                continue
            ref = cell.get("r")
            cells[(int(row.get("r")), ref)] = (
                cell.get("s"),
                _cell_value(cell, shared_strings),
            )

    return cells


def _dimension(path):
    root = ElementTree.fromstring(_read_sheet(path))
    dimension = root.find(MAIN_NS + "dimension")
    return dimension.get("ref")


###############################################################################
#
# Test scenarios.
#
###############################################################################
def _write_scenario(path, constant_memory, scenario):
    options = {"constant_memory": True, "in_memory": False} if constant_memory else {}
    workbook = Workbook(path, options)
    worksheet = workbook.add_worksheet()
    cell_format = workbook.add_format({"num_format": "0.00"})

    scenario(worksheet, workbook, cell_format)

    workbook.close()


def _scenario_basic(worksheet, workbook, cell_format):
    worksheet.add_table(
        "A1:C5",
        {
            "total_row": True,
            "columns": [
                {"header": "A"},
                {"header": "B", "total_function": "sum"},
                {
                    "header": "C",
                    "formula": "=A2+B2",
                    "total_function": "sum",
                    "format": cell_format,
                },
            ],
        },
    )

    for row in range(1, 5):
        worksheet.write(row, 0, row)
        worksheet.write(row, 1, row * 10)


def _scenario_no_user_rows(worksheet, workbook, cell_format):
    # The caller never writes any cell: headers, calculated formulas and the
    # totals row all have to be generated.
    worksheet.add_table(
        "A1:C5",
        {
            "total_row": True,
            "columns": [
                {"header": "A"},
                {"header": "B", "total_function": "sum"},
                {"header": "C", "formula": "=A2+B2", "total_function": "sum"},
            ],
        },
    )


def _scenario_skipped_rows(worksheet, workbook, cell_format):
    # Write only some of the table rows.
    worksheet.add_table(
        "A1:C6",
        {
            "total_row": True,
            "columns": [
                {"header": "A"},
                {"header": "B", "total_function": "sum"},
                {"header": "C", "formula": "=A2+B2"},
            ],
        },
    )

    for row in (1, 4):
        worksheet.write(row, 0, row)
        worksheet.write(row, 1, row * 3)


def _scenario_user_overwrite(worksheet, workbook, cell_format):
    # Write over calculated/header cells, and set row properties.
    worksheet.add_table(
        "A1:C4",
        {
            "total_row": True,
            "columns": [
                {"header": "A"},
                {"header": "B", "total_function": "sum"},
                {"header": "C", "formula": "=A2+B2"},
            ],
        },
    )

    worksheet.write(2, 2, "USER")
    worksheet.write(2, 0, 99)
    worksheet.set_row(2, 30)
    worksheet.set_row(1, 18, None, {"hidden": True})


def _scenario_data_option(worksheet, workbook, cell_format):
    data = [
        ["Apples", 10000, 5000],
        ["Pears", 2000, 3000],
    ]

    worksheet.add_table(
        "B3:E8",
        {
            "data": data,
            "total_row": True,
            "columns": [
                {"header": "Product"},
                {"header": "Q1"},
                {"header": "Q2", "total_function": "sum"},
                {
                    "header": "Total",
                    "formula": "=SUM(B4:D4)",
                    "total_function": "sum",
                },
            ],
        },
    )

    for row in (3, 4, 5, 6):
        worksheet.set_row(row, 18)

    worksheet.write("B6", "Manual")
    worksheet.write("C7", 33)


def _scenario_two_tables(worksheet, workbook, cell_format):
    # Two tables interleaved with each other, in different columns.
    worksheet.add_table(
        "A1:C6",
        {
            "total_row": True,
            "columns": [
                {"header": "A"},
                {"header": "B", "total_function": "sum"},
                {"header": "C", "formula": "=A2+B2"},
            ],
        },
    )
    worksheet.add_table(
        "E1:G6",
        {
            "total_row": True,
            "columns": [
                {"header": "X"},
                {"header": "Y", "total_function": "sum"},
                {"header": "Z", "formula": "=E2+F2"},
            ],
        },
    )

    for row in range(1, 5):
        worksheet.write(row, 0, row)
        worksheet.write(row, 1, row * 2)
        worksheet.write(row, 4, row * 10)
        worksheet.write(row, 5, row * 20)


def _scenario_second_table_later(worksheet, workbook, cell_format):
    # A second table that starts after rows of the first have been flushed.
    worksheet.add_table(
        "A1:C4",
        {"columns": [{"header": "A"}, {"header": "B"}, {"header": "C"}]},
    )

    for row in range(1, 4):
        worksheet.write(row, 0, row)

    worksheet.add_table(
        "E4:G9",
        {
            "total_row": True,
            "columns": [
                {"header": "X"},
                {"header": "Y", "total_function": "sum"},
                {"header": "Z", "formula": "=E5+F5"},
            ],
        },
    )

    for row in range(4, 9):
        worksheet.write(row, 4, row * 3)


def _scenario_merge(worksheet, workbook, cell_format):
    worksheet.add_table(
        "A1:C3",
        {"columns": [{"header": "A"}, {"header": "B"}, {"header": "C"}]},
    )
    worksheet.merge_range("E1:G2", "merged")
    worksheet.write(1, 0, 1)


SCENARIOS = {
    "basic": _scenario_basic,
    "no_user_rows": _scenario_no_user_rows,
    "skipped_rows": _scenario_skipped_rows,
    "user_overwrite": _scenario_user_overwrite,
    "data_option": _scenario_data_option,
    "two_tables": _scenario_two_tables,
    "second_table_later": _scenario_second_table_later,
    "merge": _scenario_merge,
}


class TestConstantMemoryTable(unittest.TestCase):
    """
    Test add_table() in constant_memory mode against normal mode output.

    """

    def setUp(self):
        self.tmpdir = tempfile.mkdtemp()

    def tearDown(self):
        for filename in os.listdir(self.tmpdir):
            os.unlink(os.path.join(self.tmpdir, filename))
        os.rmdir(self.tmpdir)

    def _build(self, name, constant_memory, scenario):
        path = os.path.join(
            self.tmpdir, f"{name}_{'cm' if constant_memory else 'normal'}.xlsx"
        )
        _write_scenario(path, constant_memory, scenario)
        return path

    def test_table_output_matches_normal_mode(self):
        """Test that constant_memory tables match normal mode cell by cell."""

        for name, scenario in SCENARIOS.items():
            with self.subTest(scenario=name):
                normal = self._build(name, False, scenario)
                cm = self._build(name, True, scenario)

                self.assertEqual(_dimension(normal), _dimension(cm))
                self.assertEqual(_cell_map(normal), _cell_map(cm))

    def test_table_returns_success(self):
        """Test that add_table() returns 0 in constant_memory mode."""

        path = os.path.join(self.tmpdir, "return.xlsx")
        workbook = Workbook(
            path, {"constant_memory": True, "in_memory": False}
        )
        worksheet = workbook.add_worksheet()

        self.assertEqual(worksheet.add_table("A1:C5"), 0)

        workbook.close()

    def test_table_starting_before_flushed_row_returns_minus_3(self):
        """Test that a table that starts in a flushed row returns -3."""

        path = os.path.join(self.tmpdir, "flushed.xlsx")
        workbook = Workbook(
            path, {"constant_memory": True, "in_memory": False}
        )
        worksheet = workbook.add_worksheet()

        worksheet.write(5, 0, 1)

        with warnings.catch_warnings(record=True):
            warnings.simplefilter("always")
            result = worksheet.add_table("A1:C3")

        self.assertEqual(result, -3)

        workbook.close()

    def test_table_and_merge_overlap_in_either_order(self):
        """Test table/merge overlap detection in either call order."""

        path = os.path.join(self.tmpdir, "overlap.xlsx")

        workbook = Workbook(
            path, {"constant_memory": True, "in_memory": False}
        )
        worksheet = workbook.add_worksheet()
        worksheet.add_table("A1:C5")
        with self.assertRaises(OverlappingRange):
            worksheet.merge_range("B2:D6", "")
        workbook.close()

        os.unlink(path)

        workbook = Workbook(
            path, {"constant_memory": True, "in_memory": False}
        )
        worksheet = workbook.add_worksheet()
        worksheet.merge_range("B2:D6", "")
        with self.assertRaises(OverlappingRange):
            worksheet.add_table("A1:C5")
        workbook.close()

    def test_two_tables_overlap(self):
        """Test that two overlapping tables raise in constant_memory mode."""

        path = os.path.join(self.tmpdir, "table_overlap.xlsx")
        workbook = Workbook(
            path, {"constant_memory": True, "in_memory": False}
        )
        worksheet = workbook.add_worksheet()
        worksheet.add_table("A1:C5")

        with self.assertRaises(OverlappingRange):
            worksheet.add_table("B2:D6")

        workbook.close()

    def test_overlap_reported_even_over_flushed_rows(self):
        """Test that an overlap is reported before the flushed row check."""

        path = os.path.join(self.tmpdir, "flushed_overlap.xlsx")
        workbook = Workbook(
            path, {"constant_memory": True, "in_memory": False}
        )
        worksheet = workbook.add_worksheet()
        worksheet.add_table("A1:C3")
        worksheet.write(4, 0, 1)

        with self.assertRaises(OverlappingRange):
            worksheet.merge_range("A2:B2", "")

        workbook.close()

    def test_no_table_constant_memory_unchanged(self):
        """Test that constant_memory without tables is unchanged."""

        def scenario(worksheet, workbook, cell_format):
            for row in range(10):
                worksheet.write(row, 0, row)
                worksheet.write(row, 1, f"row {row}")

        normal = self._build("plain", False, scenario)
        cm = self._build("plain", True, scenario)

        self.assertEqual(_dimension(normal), _dimension(cm))
        self.assertEqual(_cell_map(normal), _cell_map(cm))


###############################################################################
#
# Peak memory check. This is run in a separate process for each row count and
# is skipped unless XLSXWRITER_RUN_MEMORY_TEST is set, since it writes 1M rows.
#
###############################################################################
def _memory_worker(row_count, path):
    # Write a 3 column table with a calculated column and a totals row.
    import resource

    workbook = Workbook(path, {"constant_memory": True, "in_memory": False})
    worksheet = workbook.add_worksheet()
    worksheet.add_table(
        0,
        0,
        row_count,
        2,
        {
            "total_row": True,
            "columns": [
                {"header": "A"},
                {"header": "B", "total_function": "sum"},
                {"header": "C", "formula": "=A2+B2", "total_function": "sum"},
            ],
        },
    )

    for row in range(1, row_count):
        worksheet.write(row, 0, row)
        worksheet.write(row, 1, row * 2)

    workbook.close()

    return resource.getrusage(resource.RUSAGE_SELF).ru_maxrss


class TestConstantMemoryTablePeakMemory(unittest.TestCase):
    """
    Test that a table in constant_memory mode has flat peak memory.

    Set XLSXWRITER_RUN_MEMORY_TEST=1 to run.

    """

    @unittest.skipUnless(
        os.environ.get("XLSXWRITER_RUN_MEMORY_TEST"),
        "Set XLSXWRITER_RUN_MEMORY_TEST=1 to run the 1M row memory check",
    )
    def test_peak_rss_is_flat(self):
        """Test that 1M rows use at most 4 MiB more peak RSS than 100k rows."""

        import subprocess
        import sys

        package_root = os.path.dirname(
            os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        )
        script = os.path.abspath(__file__)

        tmpdir = tempfile.mkdtemp()

        def run(row_count):
            path = os.path.join(tmpdir, f"mem_{row_count}.xlsx")
            output = subprocess.check_output(
                [
                    sys.executable,
                    "-I",
                    "-c",
                    _MEMORY_MAIN,
                    script,
                    str(row_count),
                    path,
                    package_root,
                ]
            )
            # ru_maxrss is in KiB on Linux.
            return int(output.split()[-1])

        try:
            small = run(100_000)
            large = run(1_000_000)
        finally:
            for filename in os.listdir(tmpdir):
                os.unlink(os.path.join(tmpdir, filename))
            os.rmdir(tmpdir)

        # 4 MiB, in KiB.
        self.assertLessEqual(large - small, 4 * 1024)


# A small runner so the memory check can import this module in a fresh,
# isolated interpreter without importing the test harness.
_MEMORY_MAIN = """
import sys
sys.path.insert(0, sys.argv[4])
import importlib.util
spec = importlib.util.spec_from_file_location('mem_mod', sys.argv[1])
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)
print(mod._memory_worker(int(sys.argv[2]), sys.argv[3]))
"""
