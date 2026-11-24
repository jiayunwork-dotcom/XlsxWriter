###############################################################################
#
# Tests for Worksheet.materialize_autofilter().
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c) 2013-2025, John McNamara, jmcnamara@cpan.org
#
"""
Tests for the materialized autofilter.

The expected values in these tests are written out by hand from the source
data. They are never produced by the filtering code under test. The generated
XLSX file is inspected by independently parsing the worksheet, autoFilter and
chart XML.
"""

import os
import re
import tempfile
import unittest
import zipfile
import xml.etree.ElementTree as ET

from xlsxwriter.workbook import Workbook

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS = {"m": NS_MAIN, "c": NS_CHART}


def _read_members(path):
    members = {}
    with zipfile.ZipFile(path) as zf:
        for name in zf.namelist():
            members[name] = zf.read(name)
    return members


def _sst_strings(sst_bytes):
    root = ET.fromstring(sst_bytes)
    strings = []
    for si in root.findall("m:si", NS):
        runs = si.findall("m:r", NS)
        if runs:
            text = "".join(
                (t.text or "") for r in runs for t in r.findall("m:t", NS)
            )
            strings.append(text)
        else:
            t = si.find("m:t", NS)
            strings.append((t.text or "") if t is not None else "")
    return strings


def _row_hidden(sheet_bytes):
    """Return {row_number (1 based): is_hidden} for every row in sheetData."""
    root = ET.fromstring(sheet_bytes)
    rows = {}
    for row in root.findall(".//m:sheetData/m:row", NS):
        rows[int(row.get("r"))] = row.get("hidden") == "1"
    return rows


def _row_heights(sheet_bytes):
    root = ET.fromstring(sheet_bytes)
    result = {}
    for row in root.findall(".//m:sheetData/m:row", NS):
        result[int(row.get("r"))] = row.get("ht")
    return result


def _sheet_cell_texts(sheet_bytes, sst):
    """Return {(row, col): display text or number string} for every cell."""
    root = ET.fromstring(sheet_bytes)
    cells = {}
    for c in root.findall(".//m:sheetData/m:row/m:c", NS):
        col_letters, row_str = re.match(r"([A-Z]+)(\d+)", c.get("r")).groups()
        col = 0
        for ch in col_letters:
            col = col * 26 + (ord(ch) - ord("A") + 1)
        col -= 1
        row = int(row_str) - 1
        v = c.find("m:v", NS)
        if c.get("t") == "s":
            cells[(row, col)] = sst[int(v.text)]
        elif v is not None:
            cells[(row, col)] = v.text
    return cells


def _chart_cache(chart_bytes):
    """Return a list of series caches with ptCount, points and kinds."""
    root = ET.fromstring(chart_bytes)
    result = []
    for ser in root.findall(".//c:ser", NS):
        entry = {}
        cat = ser.find("c:cat", NS)
        val = ser.find("c:val", NS)
        entry["cat"] = _cache_points(cat)
        entry["val"] = _cache_points(val)
        result.append(entry)
    return result


def _cache_points(element):
    if element is None:
        return None
    for tag in ("c:strRef", "c:numRef"):
        ref = element.find(tag, NS)
        if ref is None:
            continue
        for cache_tag in ("c:strCache", "c:numCache"):
            cache = ref.find(cache_tag, NS)
            if cache is None:
                continue
            count = int(cache.find("c:ptCount", NS).get("val"))
            points = {}
            for pt in cache.findall("c:pt", NS):
                points[int(pt.get("idx"))] = pt.find("c:v", NS).text
            return {"kind": cache_tag, "count": count, "points": points}
    return None


class MaterializeFilterTest(unittest.TestCase):
    def setUp(self):
        self.tmpdir = tempfile.mkdtemp(prefix="xlsxwriter_filter_")
        self.outfile = os.path.join(self.tmpdir, "test.xlsx")

    def tearDown(self):
        for root, dirs, names in os.walk(self.tmpdir, topdown=False):
            for name in names:
                try:
                    os.remove(os.path.join(root, name))
                except OSError:
                    pass
            for name in dirs:
                try:
                    os.rmdir(os.path.join(root, name))
                except OSError:
                    pass
        try:
            os.rmdir(self.tmpdir)
        except OSError:
            pass

    def _workbook(self):
        return Workbook(
            self.outfile,
            {
                "constant_memory": True,
                "shared_strings": "disk",
                "tmpdir": self.tmpdir,
            },
        )

    def test_manual_table_two_column_conditions(self):
        """A hand checked table with two column conditions and manual hides."""
        wb = self._workbook()
        ws = wb.add_worksheet("Data")

        # Header row and five data rows.
        ws.autofilter("A1:C6")
        ws.write_row(0, 0, ["Region", "Qty", "Label"])

        rows = [
            ("East", 10, "alpha"),
            ("West", 20, "Beta"),
            ("east", 30, "gamma"),
            ("North", 0, "ÉAST"),
            ("East", 0, ""),
        ]
        for i, (region, qty, label) in enumerate(rows, start=1):
            # The "North" row (zero based row 4, XML row 5) is manually
            # hidden by the caller. In constant_memory mode a row must be
            # formatted before the next row flushes it.
            if i == 4:
                ws.set_row(4, options={"hidden": True})
            ws.write_string(i, 0, region)
            ws.write_number(i, 1, qty)
            ws.write_string(i, 2, label)

        ws.filter_column("A", "x == East")
        ws.filter_column("B", "x > 0 and x <= 30")
        ws.materialize_autofilter()
        wb.close()

        # Hand evaluated: A matches East (case insensitive): XML rows 2, 4, 6.
        # B matches 0 < qty <= 30: XML rows 2, 3, 4. Intersection: rows 2 and
        # 4. Row 5 was already hidden by the caller and stays hidden; row 6
        # has qty 0 so the filter hides it too.
        hidden = _row_hidden(_read_members(self.outfile)["xl/worksheets/sheet1.xml"])
        self.assertEqual(
            hidden, {1: False, 2: False, 3: True, 4: False, 5: True, 6: True}
        )

    def test_blank_and_rich_text(self):
        """Blanks, empty strings, rich text and hyperlink display text."""
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        bold = wb.add_format({"bold": True})

        ws.autofilter("A1:A7")
        ws.write_string(0, 0, "Value")
        ws.write_string(1, 0, "")  # explicit empty string -> blank
        ws.write_blank(2, 0, None, bold)  # formatted blank -> blank
        # row 3 is never written -> blank
        ws.write_rich_string(4, 0, "ri", bold, "ch")  # display "rich"
        ws.write_url(5, 0, "https://example.com", None, "link text")
        ws.write_string(6, 0, "other")

        ws.filter_column("A", "x == Blanks")
        ws.materialize_autofilter()
        wb.close()

        hidden = _row_hidden(_read_members(self.outfile)["xl/worksheets/sheet1.xml"])
        # Blanks: XML rows 2 and 3 (empty string, formatted blank). XML row 4
        # corresponds to the never-written physical row: it participates in
        # the decision as a blank but has no <row> element to mark.
        self.assertEqual(
            hidden, {1: False, 2: False, 3: False, 5: True, 6: True, 7: True}
        )

    def test_wildcards_escape_and_case(self):
        """Wildcard, tilde escape and case insensitive matching."""
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        ws.autofilter("A1:A6")
        ws.write_string(0, 0, "Value")
        ws.write_string(1, 0, "apple")
        ws.write_string(2, 0, "Apricot")
        ws.write_string(3, 0, "a~b")  # literal tilde
        ws.write_string(4, 0, "axb")
        ws.write_string(5, 0, "banana")

        ws.filter_column("A", "x == a*")
        ws.materialize_autofilter()
        wb.close()

        hidden = _row_hidden(_read_members(self.outfile)["xl/worksheets/sheet1.xml"])
        # a* matches apple, Apricot (case insensitive), a~b, axb; not banana.
        self.assertEqual(
            hidden, {1: False, 2: False, 3: False, 4: False, 5: False, 6: True}
        )

    def test_tilde_escape_literal(self):
        """The tilde escape matches a literal wildcard character."""
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        ws.autofilter("A1:A5")
        ws.write_string(0, 0, "Value")
        ws.write_string(1, 0, "a*b")
        ws.write_string(2, 0, "axb")
        ws.write_string(3, 0, "a?b")
        ws.write_string(4, 0, "a~b")

        ws.filter_column("A", "x == a~*b")
        ws.materialize_autofilter()
        wb.close()

        hidden = _row_hidden(_read_members(self.outfile)["xl/worksheets/sheet1.xml"])
        self.assertEqual(
            hidden, {1: False, 2: False, 3: True, 4: True, 5: True}
        )

    def test_number_zero_and_text_number(self):
        """0 is not blank and the text '12' is not the number 12."""
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        ws.autofilter("A1:A5")
        ws.write_string(0, 0, "Value")
        ws.write_number(1, 0, 0)
        ws.write_string(2, 0, "0")
        ws.write_number(3, 0, 12)
        ws.write_string(4, 0, "12")

        ws.filter_column("A", "x == 0")
        ws.materialize_autofilter()
        wb.close()

        hidden = _row_hidden(_read_members(self.outfile)["xl/worksheets/sheet1.xml"])
        # Only the numeric 0 matches; the text "0" is a string.
        self.assertEqual(
            hidden, {1: False, 2: False, 3: True, 4: True, 5: True}
        )

    def test_list_filter_literal_text(self):
        """filter_column_list() compares text literally and ignores case."""
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        ws.autofilter("A1:A6")
        ws.write_string(0, 0, "Value")
        ws.write_string(1, 0, "East")
        ws.write_string(2, 0, "west")
        ws.write_string(3, 0, "WEST")
        ws.write_string(4, 0, "North")
        ws.write_string(5, 0, "")

        ws.filter_column_list("A", ["East", "West", "Blanks"])
        ws.materialize_autofilter()
        wb.close()

        hidden = _row_hidden(_read_members(self.outfile)["xl/worksheets/sheet1.xml"])
        self.assertEqual(
            hidden, {1: False, 2: False, 3: False, 4: False, 5: True, 6: False}
        )

    def test_list_empty_string_member_matches_blanks(self):
        """An empty-list member behaves like Blanks because empty text is blank."""
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        ws.autofilter("A1:A4")
        ws.write_string(0, 0, "V")
        ws.write_string(1, 0, "a")
        ws.write_string(2, 0, "")
        # row 4 (zero based 3) is never written and is therefore blank.
        ws.filter_column_list("A", ["a", ""])
        ws.materialize_autofilter()
        wb.close()

        hidden = _row_hidden(_read_members(self.outfile)["xl/worksheets/sheet1.xml"])
        self.assertEqual(hidden, {1: False, 2: False, 3: False})

    def test_list_filter_star_is_literal(self):
        """A '*' in a list member is matched literally, not as a wildcard."""
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        ws.autofilter("A1:A4")
        ws.write_string(0, 0, "Value")
        ws.write_string(1, 0, "a*b")
        ws.write_string(2, 0, "axb")
        ws.write_string(3, 0, "ab")

        ws.filter_column_list("A", ["a*b"])
        ws.materialize_autofilter()
        wb.close()

        hidden = _row_hidden(_read_members(self.outfile)["xl/worksheets/sheet1.xml"])
        self.assertEqual(hidden, {1: False, 2: False, 3: True, 4: True})

    def test_numeric_comparisons(self):
        """All six numeric comparisons against a finite threshold."""
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        ws.autofilter("A1:A6")
        ws.write_string(0, 0, "Value")
        for i, value in enumerate([1, 2, 3, 4, 5], start=1):
            ws.write_number(i, 0, value)

        ws.filter_column("A", "x >= 2 and x < 5")
        ws.materialize_autofilter()
        wb.close()

        hidden = _row_hidden(_read_members(self.outfile)["xl/worksheets/sheet1.xml"])
        # 2, 3, 4 match; 1 and 5 don't.
        self.assertEqual(
            hidden, {1: False, 2: True, 3: False, 4: False, 5: False, 6: True}
        )

    def test_number_vs_text_type_mismatch(self):
        """A numeric comparison never matches a text cell and vice versa."""
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        ws.autofilter("A1:A4")
        ws.write_string(0, 0, "Value")
        ws.write_number(1, 0, 5)
        ws.write_string(2, 0, "5")
        ws.write_number(3, 0, 6)

        ws.filter_column("A", "x > 5")
        ws.materialize_autofilter()
        wb.close()

        hidden = _row_hidden(_read_members(self.outfile)["xl/worksheets/sheet1.xml"])
        # Only the number 6 is greater than 5. The text "5" isn't numeric.
        self.assertEqual(hidden, {1: False, 2: True, 3: True, 4: False})

    def test_unfiltered_columns_are_not_restricted(self):
        """Formulas, booleans and errors outside a filtered column are fine."""
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        ws.autofilter("A1:C3")
        ws.write_string(0, 0, "A")
        ws.write_string(0, 1, "B")
        ws.write_string(0, 2, "C")

        ws.write_string(1, 0, "keep")
        ws.write_formula(1, 1, "=1+1", None, 2)
        ws.write_boolean(1, 2, True)

        ws.write_string(2, 0, "drop")
        ws.write_number(2, 1, 1)
        ws.write_number(2, 2, 2)

        ws.filter_column("A", "x == keep")
        ws.materialize_autofilter()
        wb.close()

        hidden = _row_hidden(_read_members(self.outfile)["xl/worksheets/sheet1.xml"])
        self.assertEqual(hidden, {1: False, 2: False, 3: True})

    def test_formula_in_filtered_column_fails(self):
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        ws.autofilter("A1:A2")
        ws.write_string(0, 0, "A")
        ws.write_formula(1, 0, "=1+1", None, 2)
        ws.filter_column("A", "x == 2")
        ws.materialize_autofilter()
        with self.assertRaises(ValueError):
            wb.close()
        self.assertFalse(os.path.exists(self.outfile))

    def test_boolean_in_filtered_column_fails(self):
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        ws.autofilter("A1:A2")
        ws.write_string(0, 0, "A")
        ws.write_boolean(1, 0, True)
        ws.filter_column("A", "x == 1")
        ws.materialize_autofilter()
        with self.assertRaises(ValueError):
            wb.close()

    def test_error_in_filtered_column_fails(self):
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        ws.autofilter("A1:A2")
        ws.write_string(0, 0, "A")
        ws.write_formula(1, 0, "=1/0", None, "#DIV/0!")
        ws.filter_column("A", "x == 1")
        ws.materialize_autofilter()
        with self.assertRaises(ValueError):
            wb.close()

    def test_unsupported_numeric_comparison_fails(self):
        """A text threshold for an ordering operator fails at close."""
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        ws.autofilter("A1:A2")
        ws.write_string(0, 0, "A")
        ws.write_number(1, 0, 1)
        ws.filter_column("A", "x > abc")
        ws.materialize_autofilter()
        with self.assertRaises(ValueError):
            wb.close()

    def test_supported_operator_aliases(self):
        """The parser operator aliases are accepted."""
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        ws.autofilter("A1:A4")
        ws.write_string(0, 0, "Value")
        ws.write_number(1, 0, 1)
        ws.write_number(2, 0, 2)
        ws.write_number(3, 0, 3)

        ws.filter_column("A", "x != 2")
        ws.materialize_autofilter()
        wb.close()

        hidden = _row_hidden(_read_members(self.outfile)["xl/worksheets/sheet1.xml"])
        self.assertEqual(hidden, {1: False, 2: False, 3: True, 4: False})

    def test_materialize_after_data(self):
        """The method can be called before or after writing the rows."""
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        ws.autofilter("A1:A3")
        ws.write_string(0, 0, "V")
        ws.write_string(1, 0, "yes")
        ws.write_string(2, 0, "no")
        ws.filter_column("A", "x == yes")
        ws.materialize_autofilter()
        wb.close()

        hidden = _row_hidden(_read_members(self.outfile)["xl/worksheets/sheet1.xml"])
        self.assertEqual(hidden, {1: False, 2: False, 3: True})

    def test_data_written_after_call_is_included(self):
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        ws.autofilter("A1:A4")
        ws.write_string(0, 0, "V")
        ws.write_number(1, 0, 1)
        ws.write_number(2, 0, 2)
        ws.materialize_autofilter()  # Called before all of the data exists.
        ws.write_number(3, 0, 10)
        ws.write_number(4, 0, 20)
        ws.filter_column("A", "x > 5")
        wb.close()

        hidden = _row_hidden(_read_members(self.outfile)["xl/worksheets/sheet1.xml"])
        self.assertEqual(
            hidden, {1: False, 2: True, 3: True, 4: False, 5: False}
        )

    def test_repeated_call_is_idempotent(self):
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        ws.autofilter("A1:A3")
        ws.write_string(0, 0, "V")
        ws.write_string(1, 0, "yes")
        ws.write_string(2, 0, "no")
        ws.filter_column("A", "x == yes")
        ws.materialize_autofilter()
        ws.materialize_autofilter()
        wb.close()

        hidden = _row_hidden(_read_members(self.outfile)["xl/worksheets/sheet1.xml"])
        self.assertEqual(hidden, {1: False, 2: False, 3: True})

    def test_late_condition_replace_uses_final(self):
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        ws.autofilter("A1:A4")
        ws.write_string(0, 0, "V")
        for i, value in enumerate(["a", "b", "c"], start=1):
            ws.write_string(i, 0, value)
        ws.materialize_autofilter()  # Too early: no condition yet.
        ws.filter_column("A", "x == a")
        ws.filter_column("A", "x == b")  # Replace before close.
        wb.close()

        hidden = _row_hidden(_read_members(self.outfile)["xl/worksheets/sheet1.xml"])
        self.assertEqual(hidden, {1: False, 2: True, 3: False, 4: True})

    def test_conditional_replaced_after_data(self):
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        ws.autofilter("A1:A3")
        ws.write_string(0, 0, "V")
        ws.write_number(1, 0, 1)
        ws.write_number(2, 0, 2)
        ws.filter_column("A", "x > 0")
        ws.materialize_autofilter()
        ws.filter_column("A", "x > 5")  # Too late for row 1 to still match.
        wb.close()

        hidden = _row_hidden(_read_members(self.outfile)["xl/worksheets/sheet1.xml"])
        self.assertEqual(hidden, {1: False, 2: True, 3: True})

    def test_initial_condition_replaced_by_list(self):
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        ws.autofilter("A1:A4")
        ws.write_string(0, 0, "V")
        for i, value in enumerate(["x", "y", "z"], start=1):
            ws.write_string(i, 0, value)
        ws.filter_column("A", "x == x")
        ws.materialize_autofilter()
        ws.filter_column_list("A", ["y", "z"])
        wb.close()

        hidden = _row_hidden(_read_members(self.outfile)["xl/worksheets/sheet1.xml"])
        self.assertEqual(hidden, {1: False, 2: True, 3: False, 4: False})

    def test_all_match(self):
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        ws.autofilter("A1:A3")
        ws.write_string(0, 0, "V")
        ws.write_string(1, 0, "a")
        ws.write_string(2, 0, "a")
        ws.filter_column("A", "x == a")
        ws.materialize_autofilter()
        wb.close()

        hidden = _row_hidden(_read_members(self.outfile)["xl/worksheets/sheet1.xml"])
        self.assertEqual(hidden, {1: False, 2: False, 3: False})

    def test_all_hidden(self):
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        ws.autofilter("A1:A3")
        ws.write_string(0, 0, "V")
        ws.write_string(1, 0, "a")
        ws.write_string(2, 0, "b")
        ws.filter_column("A", "x == zzz")
        ws.materialize_autofilter()
        wb.close()

        hidden = _row_hidden(_read_members(self.outfile)["xl/worksheets/sheet1.xml"])
        self.assertEqual(hidden, {1: False, 2: True, 3: True})

    def test_unwritten_physical_rows_participate(self):
        """Rows in the range that were never written are blanks."""
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        ws.autofilter("A1:A6")
        ws.write_string(0, 0, "V")
        ws.write_string(1, 0, "a")
        # Rows 3, 4, 5 (zero based 2, 3, 4) are never written.
        ws.write_string(5, 0, "a")
        ws.filter_column("A", "x == Blanks")
        ws.materialize_autofilter()
        wb.close()

        hidden = _row_hidden(_read_members(self.outfile)["xl/worksheets/sheet1.xml"])
        # Only physical rows that exist have hidden attributes; rows 2, 3, 4
        # were never written so they don't appear in the XML at all.
        self.assertEqual(hidden, {1: False, 2: True, 6: True})

    def test_header_and_outside_rows_excluded(self):
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        ws.autofilter("A2:A4")
        ws.write_string(0, 0, "Not a header")
        ws.write_string(1, 0, "Header")
        ws.write_string(2, 0, "match")
        ws.write_string(3, 0, "nope")
        ws.write_string(4, 0, "outside")  # Outside the filter range.
        ws.filter_column("A", "x == match")
        ws.materialize_autofilter()
        wb.close()

        hidden = _row_hidden(_read_members(self.outfile)["xl/worksheets/sheet1.xml"])
        # Header row 2 and the outside row 5 aren't evaluated.
        self.assertEqual(hidden, {1: False, 2: False, 3: False, 4: True, 5: False})

    def test_row_properties_preserved(self):
        """Heights, levels and manual hidden flags survive the filter."""
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        ws.autofilter("A1:A4")
        ws.write_string(0, 0, "V")
        ws.set_row(1, 30, None, {"level": 2, "collapsed": True})
        ws.set_row(2, 22, None, {"hidden": True})
        ws.write_string(1, 0, "keep")
        ws.write_string(2, 0, "drop")
        ws.write_string(3, 0, "keep")
        ws.filter_column("A", "x == keep")
        ws.materialize_autofilter()
        wb.close()

        members = _read_members(self.outfile)
        heights = _row_heights(members["xl/worksheets/sheet1.xml"])
        self.assertEqual(heights[2], "30")
        self.assertEqual(heights[3], "22")
        sheet = members["xl/worksheets/sheet1.xml"].decode()
        self.assertIn('r="2"', sheet)
        # The collapsed/outline row keeps its attributes and is not hidden.
        self.assertIn('outlineLevel="2"', sheet)
        self.assertIn('collapsed="1"', sheet)

    def test_manual_hidden_not_expanded(self):
        """A manually hidden row that matches the filter stays hidden."""
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        ws.autofilter("A1:A3")
        ws.write_string(0, 0, "V")
        ws.set_row(1, options={"hidden": True})
        ws.write_string(1, 0, "keep")
        ws.write_string(2, 0, "keep")
        ws.filter_column("A", "x == keep")
        ws.materialize_autofilter()
        wb.close()

        hidden = _row_hidden(_read_members(self.outfile)["xl/worksheets/sheet1.xml"])
        self.assertEqual(hidden, {1: False, 2: True, 3: False})

    def test_multiple_worksheets_independent(self):
        wb = self._workbook()
        ws1 = wb.add_worksheet("One")
        ws2 = wb.add_worksheet("Two")
        ws3 = wb.add_worksheet("Three")

        ws1.autofilter("A1:A3")
        ws1.write_string(0, 0, "V")
        ws1.write_string(1, 0, "a")
        ws1.write_string(2, 0, "b")
        ws1.filter_column("A", "x == a")
        ws1.materialize_autofilter()

        # ws2 doesn't enable materialization even though it has a filter.
        ws2.autofilter("A1:A2")
        ws2.write_string(0, 0, "V")
        ws2.write_string(1, 0, "a")
        ws2.filter_column("A", "x == b")

        # ws3 has no filter at all.
        ws3.write_string(0, 0, "V")
        ws3.write_string(1, 0, "a")

        wb.close()

        members = _read_members(self.outfile)
        hidden1 = _row_hidden(members["xl/worksheets/sheet1.xml"])
        hidden2 = _row_hidden(members["xl/worksheets/sheet2.xml"])
        hidden3 = _row_hidden(members["xl/worksheets/sheet3.xml"])
        self.assertEqual(hidden1, {1: False, 2: False, 3: True})
        # ws2 keeps the old behaviour: no rows hidden by us.
        self.assertEqual(hidden2, {1: False, 2: False})
        self.assertEqual(hidden3, {1: False, 2: False})

    def test_non_disk_workbook_raises(self):
        for options in (
            {},
            {"constant_memory": True},
            {"in_memory": True},
        ):
            wb = Workbook(os.path.join(self.tmpdir, "x.xlsx"), options)
            ws = wb.add_worksheet()
            with self.assertRaises(ValueError):
                ws.materialize_autofilter()
            wb.close()

    def test_blank_not_equal_explicit(self):
        """!= Blanks matches only non-blank cells."""
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        ws.autofilter("A1:A3")
        ws.write_string(0, 0, "V")
        ws.write_string(1, 0, "")
        ws.write_string(2, 0, "x")
        ws.filter_column("A", "x != Blanks")
        ws.materialize_autofilter()
        wb.close()

        hidden = _row_hidden(_read_members(self.outfile)["xl/worksheets/sheet1.xml"])
        self.assertEqual(hidden, {1: False, 2: True, 3: False})

    def test_custom_blank_comparison_with_text(self):
        """A normal comparison never matches a blank or a mismatched type."""
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        ws.autofilter("A1:A4")
        ws.write_string(0, 0, "V")
        ws.write_string(1, 0, "abc")
        ws.write_number(2, 0, 5)
        # row 4 (zero based 3) is blank
        ws.filter_column("A", "x == abc")
        ws.materialize_autofilter()
        wb.close()

        hidden = _row_hidden(_read_members(self.outfile)["xl/worksheets/sheet1.xml"])
        # The never-written row 4 has no <row> element, so only the existing
        # mismatching rows are marked.
        self.assertEqual(hidden, {1: False, 2: False, 3: True})


class MoreMaterializeFilterTest(unittest.TestCase):
    setUp = MaterializeFilterTest.setUp
    tearDown = MaterializeFilterTest.tearDown
    _workbook = MaterializeFilterTest._workbook

    def test_equality_routed_to_list_matches_numbers(self):
        """filter_column('x == 2000') matches the number 2000, not text."""
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        ws.autofilter("A1:A3")
        ws.write_string(0, 0, "V")
        ws.write_number(1, 0, 2000)
        ws.write_string(2, 0, "2000")
        ws.filter_column("A", "x == 2000")
        ws.materialize_autofilter()
        wb.close()

        hidden = _row_hidden(_read_members(self.outfile)["xl/worksheets/sheet1.xml"])
        self.assertEqual(hidden, {1: False, 2: False, 3: True})

    def test_or_equality(self):
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        ws.autofilter("A1:A4")
        ws.write_string(0, 0, "V")
        for i, value in enumerate(["a", "b", "c"], start=1):
            ws.write_string(i, 0, value)
        ws.filter_column("A", "x == a or x == b")
        ws.materialize_autofilter()
        wb.close()

        hidden = _row_hidden(_read_members(self.outfile)["xl/worksheets/sheet1.xml"])
        self.assertEqual(hidden, {1: False, 2: False, 3: False, 4: True})

    def test_not_equal_wildcard(self):
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        ws.autofilter("A1:A4")
        ws.write_string(0, 0, "V")
        for i, value in enumerate(["alpha", "beta", "Alpine"], start=1):
            ws.write_string(i, 0, value)
        ws.filter_column("A", "x != a*")
        ws.materialize_autofilter()
        wb.close()

        hidden = _row_hidden(_read_members(self.outfile)["xl/worksheets/sheet1.xml"])
        # "alpha" and "Alpine" match a* (case insensitive) and are hidden;
        # "beta" doesn't match and remains visible.
        self.assertEqual(hidden, {1: False, 2: True, 3: False, 4: True})

    def test_last_buffered_row_matches(self):
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        ws.autofilter("A1:A2")
        ws.write_string(0, 0, "V")
        ws.write_string(1, 0, "last")  # Still buffered at close.
        ws.filter_column("A", "x == last")
        ws.materialize_autofilter()
        wb.close()

        hidden = _row_hidden(_read_members(self.outfile)["xl/worksheets/sheet1.xml"])
        self.assertEqual(hidden, {1: False, 2: False})

    def test_last_buffered_row_doesnt_match(self):
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        ws.autofilter("A1:A2")
        ws.write_string(0, 0, "V")
        ws.write_string(1, 0, "last")  # Still buffered at close.
        ws.filter_column("A", "x == other")
        ws.materialize_autofilter()
        wb.close()

        hidden = _row_hidden(_read_members(self.outfile)["xl/worksheets/sheet1.xml"])
        self.assertEqual(hidden, {1: False, 2: True})

    def test_buffered_row_overwrite_uses_final_value(self):
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        ws.autofilter("A1:A2")
        ws.write_string(0, 0, "V")
        ws.write_string(1, 0, "old")
        ws.write_string(1, 0, "new")  # Overwrite before flush.
        ws.filter_column("A", "x == new")
        ws.materialize_autofilter()
        wb.close()

        hidden = _row_hidden(_read_members(self.outfile)["xl/worksheets/sheet1.xml"])
        self.assertEqual(hidden, {1: False, 2: False})

    def test_flushed_row_overwrite_uses_final_value(self):
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        ws.autofilter("A1:A3")
        ws.write_string(0, 0, "V")
        ws.write_string(1, 0, "old")
        ws.write_string(1, 0, "new")  # Overwrite, row still buffered.
        ws.write_string(2, 0, "next")  # Flushes row 1 with its final value.
        ws.filter_column("A", "x == new")
        ws.materialize_autofilter()
        wb.close()

        hidden = _row_hidden(_read_members(self.outfile)["xl/worksheets/sheet1.xml"])
        self.assertEqual(hidden, {1: False, 2: False, 3: True})

    def test_top_ten_condition_rejected(self):
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        ws.autofilter("A1:A2")
        ws.write_string(0, 0, "V")
        ws.write_number(1, 0, 1)
        ws.filter_column("A", "Top 10 Items")
        ws.materialize_autofilter()
        with self.assertRaises(ValueError):
            wb.close()

    def test_text_ordering_condition_rejected(self):
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        ws.autofilter("A1:A2")
        ws.write_string(0, 0, "V")
        ws.write_string(1, 0, "a")
        ws.filter_column("A", "x < abc")
        ws.materialize_autofilter()
        with self.assertRaises(ValueError):
            wb.close()

    def test_failure_keeps_existing_output(self):
        original = b"pre-existing, not an xlsx"
        with open(self.outfile, "wb") as fh:
            fh.write(original)

        wb = self._workbook()
        ws = wb.add_worksheet("S")
        ws.autofilter("A1:A2")
        ws.write_string(0, 0, "V")
        ws.write_boolean(1, 0, True)
        ws.filter_column("A", "x == 1")
        ws.materialize_autofilter()
        with self.assertRaises(ValueError):
            wb.close()

        with open(self.outfile, "rb") as fh:
            self.assertEqual(fh.read(), original)
        leftovers = [
            name
            for name in os.listdir(self.tmpdir)
            if name.startswith("xlsxwriter_sst_") or name.startswith(".xlsxwriter_")
        ]
        self.assertEqual(leftovers, [])

    def test_filter_metadata_matches_materialized_conditions(self):
        """The autoFilter element and the hidden rows use the same final data."""
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        ws.autofilter("A1:B3")
        ws.write_row(0, 0, ["R", "V"])
        ws.write_string(1, 0, "a")
        ws.write_number(1, 1, 1)
        ws.write_string(2, 0, "b")
        ws.write_number(2, 1, 2)
        ws.filter_column("A", "x == a")
        ws.materialize_autofilter()
        wb.close()

        members = _read_members(self.outfile)
        sheet = members["xl/worksheets/sheet1.xml"].decode()
        self.assertIn(
            '<autoFilter ref="A1:B3"><filterColumn colId="0"><filters>'
            '<filter val="a"/></filters></filterColumn></autoFilter>',
            sheet,
        )
        hidden = _row_hidden(sheet.encode())
        self.assertEqual(hidden, {1: False, 2: False, 3: True})


class ChartFilterTest(unittest.TestCase):
    def setUp(self):
        self.tmpdir = tempfile.mkdtemp(prefix="xlsxwriter_filter_chart_")
        self.outfile = os.path.join(self.tmpdir, "test.xlsx")

    def tearDown(self):
        for root, dirs, names in os.walk(self.tmpdir, topdown=False):
            for name in names:
                try:
                    os.remove(os.path.join(root, name))
                except OSError:
                    pass
            for name in dirs:
                try:
                    os.rmdir(os.path.join(root, name))
                except OSError:
                    pass
        try:
            os.rmdir(self.tmpdir)
        except OSError:
            pass

    def _build(self, swap_creation):
        wb = Workbook(
            self.outfile,
            {
                "constant_memory": True,
                "shared_strings": "disk",
                "tmpdir": self.tmpdir,
            },
        )
        ws = wb.add_worksheet("Data")
        ws.autofilter("A1:C6")
        ws.write_row(0, 0, ["Region", "Label", "Value"])

        # Hand checked source data.
        rows = [
            ("East", "A", 10),
            ("West", "B", 20),
            ("East", "C", 30),
            ("North", "D", 40),
            ("East", "E", 50),
        ]
        for i, (region, label, value) in enumerate(rows, start=1):
            ws.write_string(i, 0, region)
            ws.write_string(i, 1, label)
            ws.write_number(i, 2, value)

        ws.filter_column("A", "x == East")
        ws.materialize_autofilter()

        def make_visible():
            chart = wb.add_chart({"type": "column"})
            chart.add_series(
                {
                    "categories": "=Data!$B$2:$B$6",
                    "values": "=Data!$C$2:$C$6",
                }
            )
            return chart

        def make_hidden():
            chart = wb.add_chart({"type": "line"})
            chart.add_series(
                {
                    "categories": "=Data!$B$2:$B$6",
                    "values": "=Data!$C$2:$C$6",
                }
            )
            chart.show_hidden_data()
            return chart

        if swap_creation:
            first = make_hidden()
            second = make_visible()
            ws.insert_chart("E2", first)
            ws.insert_chart("E20", second)
        else:
            first = make_visible()
            second = make_hidden()
            ws.insert_chart("E2", first)
            ws.insert_chart("E20", second)
        wb.close()
        return _read_members(self.outfile)

    def _expected_visible(self):
        # Rows 1, 3 and 5 of the range (A, C, E) are visible, at idx 0, 2, 4.
        return {
            "cat": {
                "kind": "c:strCache",
                "count": 5,
                "points": {0: "A", 2: "C", 4: "E"},
            },
            "val": {
                "kind": "c:numCache",
                "count": 5,
                "points": {0: "10", 2: "30", 4: "50"},
            },
        }

    def _expected_hidden(self):
        return {
            "cat": {
                "kind": "c:strCache",
                "count": 5,
                "points": {0: "A", 1: "B", 2: "C", 3: "D", 4: "E"},
            },
            "val": {
                "kind": "c:numCache",
                "count": 5,
                "points": {0: "10", 1: "20", 2: "30", 3: "40", 4: "50"},
            },
        }

    def test_shared_range_two_charts(self):
        members = self._build(swap_creation=False)
        visible = _chart_cache(members["xl/charts/chart1.xml"])[0]
        hidden = _chart_cache(members["xl/charts/chart2.xml"])[0]
        self.assertEqual(visible, self._expected_visible())
        self.assertEqual(hidden, self._expected_hidden())

    def test_shared_range_swapped_creation_order(self):
        """Swapping which chart is created first doesn't mix their modes."""
        members = self._build(swap_creation=True)
        # chart1 is now the show_hidden_data chart because it was created
        # first; it still keeps every point. chart2 is visible-only.
        first = _chart_cache(members["xl/charts/chart1.xml"])[0]
        second = _chart_cache(members["xl/charts/chart2.xml"])[0]
        self.assertEqual(first, self._expected_hidden())
        self.assertEqual(second, self._expected_visible())

    def test_cross_sheet_series(self):
        wb = Workbook(
            self.outfile,
            {
                "constant_memory": True,
                "shared_strings": "disk",
                "tmpdir": self.tmpdir,
            },
        )
        ws1 = wb.add_worksheet("Data")
        ws2 = wb.add_worksheet("Chart Sheet")
        ws1.autofilter("A1:B5")
        ws1.write_row(0, 0, ["Region", "Value"])
        rows = [("East", 1), ("West", 2), ("East", 3), ("East", 4)]
        for i, (region, value) in enumerate(rows, start=1):
            ws1.write_string(i, 0, region)
            ws1.write_number(i, 1, value)
        ws1.filter_column("A", "x == East")
        ws1.materialize_autofilter()

        chart = wb.add_chart({"type": "column"})
        chart.add_series(
            {
                "categories": "=Data!$A$2:$A$5",
                "values": "=Data!$B$2:$B$5",
            }
        )
        ws2.insert_chart("D2", chart)
        wb.close()

        members = _read_members(self.outfile)
        cache = _chart_cache(members["xl/charts/chart1.xml"])[0]
        self.assertEqual(
            cache["cat"]["points"], {0: "East", 2: "East", 3: "East"}
        )
        self.assertEqual(cache["cat"]["count"], 4)
        self.assertEqual(cache["val"]["points"], {0: "1", 2: "3", 3: "4"})

    def test_show_blanks_still_missing(self):
        """Missing cells keep the pre-existing missing value handling."""
        wb = Workbook(
            self.outfile,
            {
                "constant_memory": True,
                "shared_strings": "disk",
                "tmpdir": self.tmpdir,
            },
        )
        ws = wb.add_worksheet("Data")
        ws.autofilter("A1:B4")
        ws.write_row(0, 0, ["Region", "Value"])
        ws.write_string(1, 0, "East")
        ws.write_number(1, 1, 1)
        # Row 3: region present but value missing.
        ws.write_string(2, 0, "East")
        ws.write_string(3, 0, "West")
        ws.write_number(3, 1, 4)
        ws.filter_column("A", "x == East")
        ws.materialize_autofilter()

        chart = wb.add_chart({"type": "column"})
        chart.add_series(
            {
                "categories": "=Data!$A$2:$A$4",
                "values": "=Data!$B$2:$B$4",
            }
        )
        ws.insert_chart("D2", chart)
        wb.close()

        cache = _chart_cache(_read_members(self.outfile)["xl/charts/chart1.xml"])[0]
        # Rows 2 and 3 both have the region East and are visible, so both of
        # their category points remain. Row 3 has no value at all, so it is
        # omitted from the values cache exactly as it would be without a
        # filter. Row 4 (West) is filtered out of both caches.
        self.assertEqual(cache["cat"]["points"], {0: "East", 1: "East"})
        self.assertEqual(cache["val"]["points"], {0: "1"})
        self.assertEqual(cache["val"]["count"], 3)
        self.assertEqual(cache["cat"]["count"], 3)

    def test_single_row_range_all_hidden(self):
        """A one point range that is hidden disappears from the visible cache."""
        wb = Workbook(
            self.outfile,
            {
                "constant_memory": True,
                "shared_strings": "disk",
                "tmpdir": self.tmpdir,
            },
        )
        ws = wb.add_worksheet("S")
        ws.autofilter("A1:B2")
        ws.write_row(0, 0, ["R", "V"])
        ws.write_string(1, 0, "West")
        ws.write_number(1, 1, 5)
        ws.filter_column("A", "x == East")
        ws.materialize_autofilter()

        visible = wb.add_chart({"type": "column"})
        visible.add_series(
            {
                "categories": "=S!$A$2:$A$2",
                "values": "=S!$B$2:$B$2",
            }
        )
        hidden = wb.add_chart({"type": "line"})
        hidden.add_series(
            {
                "categories": "=S!$A$2:$A$2",
                "values": "=S!$B$2:$B$2",
            }
        )
        hidden.show_hidden_data()
        ws.insert_chart("D2", visible)
        ws.insert_chart("D20", hidden)
        wb.close()

        members = _read_members(self.outfile)
        visible_cache = _chart_cache(members["xl/charts/chart1.xml"])[0]
        hidden_cache = _chart_cache(members["xl/charts/chart2.xml"])[0]
        self.assertEqual(visible_cache["val"]["count"], 1)
        self.assertEqual(visible_cache["val"]["points"], {})
        self.assertEqual(hidden_cache["val"]["count"], 1)
        self.assertEqual(hidden_cache["val"]["points"], {0: "5"})

    def test_explicit_data_not_filtered(self):
        """A series with explicit categories_data/values_data is untouched."""
        wb = Workbook(
            self.outfile,
            {
                "constant_memory": True,
                "shared_strings": "disk",
                "tmpdir": self.tmpdir,
            },
        )
        ws = wb.add_worksheet("Data")
        ws.autofilter("A1:A3")
        ws.write_string(0, 0, "Region")
        for i, region in enumerate(["East", "West"], start=1):
            ws.write_string(i, 0, region)
        ws.filter_column("A", "x == East")
        ws.materialize_autofilter()

        chart = wb.add_chart({"type": "column"})
        chart.add_series(
            {
                "values": "=Data!$A$2:$A$3",
                "values_data": [11, 22],
            }
        )
        ws.insert_chart("C2", chart)
        wb.close()

        cache = _chart_cache(_read_members(self.outfile)["xl/charts/chart1.xml"])[0]
        # The explicit values_data is used verbatim and is not filtered.
        self.assertEqual(cache["val"]["points"], {0: "11", 1: "22"})


if __name__ == "__main__":
    unittest.main()
