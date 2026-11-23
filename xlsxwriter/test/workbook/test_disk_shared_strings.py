###############################################################################
#
# Tests for the disk backed shared strings export path.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#
"""
Tests for the ``Workbook(shared_strings='disk')`` option.

The expected values in these tests are recorded from the *source* data
written by the test itself. The generated XLSX file is then inspected by
independently parsing the worksheet, shared strings, relationships and chart
XML. Nothing is compared against an in-memory run of the same code path.
"""

import os
import re
import tempfile
import unittest
import zipfile
import xml.etree.ElementTree as ET

from xlsxwriter.workbook import Workbook

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

NS = {"m": NS_MAIN, "r": NS_REL, "c": NS_CHART, "pr": NS_PKG_REL}


def _read_zip_members(xlsx_path):
    """Return a dict of member name -> raw bytes using streaming reads."""
    members = {}
    with zipfile.ZipFile(xlsx_path) as zf:
        for name in zf.namelist():
            with zf.open(name) as fh:
                members[name] = fh.read()
    return members


def _sst_strings(sst_bytes):
    """Return the list of (is_rich, display_text) in SST order.

    The display text of a rich string is its concatenated fragments. This is
    parsed independently of the writer.
    """
    root = ET.fromstring(sst_bytes)
    strings = []
    for si in root.findall("m:si", NS):
        runs = si.findall("m:r", NS)
        if runs:
            text = "".join(
                (t.text or "")
                for r in runs
                for t in r.findall("m:t", NS)
            )
            strings.append((True, text))
        else:
            t = si.find("m:t", NS)
            strings.append((False, (t.text or "") if t is not None else ""))
    return strings


def _sst_rich_runs(sst_bytes):
    """Return the raw run structure of each rich <si> element."""
    root = ET.fromstring(sst_bytes)
    result = []
    for si in root.findall("m:si", NS):
        runs = []
        for r in si.findall("m:r", NS):
            rpr = r.find("m:rPr", NS)
            fmt = ""
            if rpr is not None:
                fmt = ET.tostring(rpr, encoding="unicode")
            t = r.find("m:t", NS)
            runs.append((fmt, (t.text or "") if t is not None else ""))
        result.append(runs)
    return result


def _sheet_cells(sheet_bytes):
    """Return {(row, col): (cell_type, value)} parsed from a worksheet."""
    root = ET.fromstring(sheet_bytes)
    cells = {}
    for c in root.findall(".//m:sheetData/m:row/m:c", NS):
        ref = c.get("r")
        col_letters = re.match(r"([A-Z]+)(\d+)", ref)
        col_str, row_str = col_letters.groups()
        col = 0
        for ch in col_str:
            col = col * 26 + (ord(ch) - ord("A") + 1)
        col -= 1
        row = int(row_str) - 1
        cell_type = c.get("t", "n")
        v = c.find("m:v", NS)
        is_ = c.find("m:is", NS)
        f = c.find("m:f", NS)
        if cell_type == "s":
            cells[(row, col)] = ("sst", int(v.text))
        elif cell_type == "inlineStr":
            cells[(row, col)] = ("inline", ET.tostring(is_, encoding="unicode"))
        elif f is not None:
            cells[(row, col)] = ("formula", v.text if v is not None else None)
        elif v is not None:
            cells[(row, col)] = (cell_type, v.text)
        else:
            cells[(row, col)] = ("blank", None)
    return cells


def _chart_cache(chart_bytes):
    """Return the first series' category and value caches as a dict."""
    root = ET.fromstring(chart_bytes)
    series = root.findall(".//c:ser", NS)
    result = []
    for ser in series:
        cat = ser.find("c:cat", NS)
        val = ser.find("c:val", NS)
        cat_points = {}
        val_points = {}
        cat_kind = None
        val_kind = None
        if cat is not None:
            str_ref = cat.find("c:strRef", NS)
            num_ref = cat.find("c:numRef", NS)
            if str_ref is not None:
                cat_kind = "str"
                cache = str_ref.find("c:strCache", NS)
                if cache is not None:
                    for pt in cache.findall("c:pt", NS):
                        v = pt.find("c:v", NS)
                        cat_points[int(pt.get("idx"))] = v.text
            elif num_ref is not None:
                cat_kind = "num"
                cache = num_ref.find("c:numCache", NS)
                if cache is not None:
                    for pt in cache.findall("c:pt", NS):
                        v = pt.find("c:v", NS)
                        cat_points[int(pt.get("idx"))] = v.text
        if val is not None:
            num_ref = val.find("c:numRef", NS)
            if num_ref is not None:
                cache = num_ref.find("c:numCache", NS)
                if cache is not None:
                    for pt in cache.findall("c:pt", NS):
                        v = pt.find("c:v", NS)
                        val_points[int(pt.get("idx"))] = v.text
            str_ref = val.find("c:strRef", NS)
            if str_ref is not None:
                val_kind = "str"
                cache = str_ref.find("c:strCache", NS)
                if cache is not None:
                    for pt in cache.findall("c:pt", NS):
                        v = pt.find("c:v", NS)
                        val_points[int(pt.get("idx"))] = v.text
        result.append(
            {
                "cat_kind": cat_kind,
                "cat": cat_points,
                "val_kind": val_kind,
                "val": val_points,
            }
        )
    return result


class DiskSharedStringsTest(unittest.TestCase):
    """Tests for the disk backed shared string export path."""

    def setUp(self):
        self.tmpdir = tempfile.mkdtemp(prefix="xlsxwriter_test_")
        self.outfile = os.path.join(self.tmpdir, "test.xlsx")

    def tearDown(self):
        for name in os.listdir(self.tmpdir):
            path = os.path.join(self.tmpdir, name)
            try:
                if os.path.isdir(path):
                    continue
                os.remove(path)
            except OSError:
                pass
        try:
            os.rmdir(self.tmpdir)
        except OSError:
            pass

    def _workbook(self, **options):
        options.setdefault("constant_memory", True)
        options.setdefault("shared_strings", "disk")
        options.setdefault("tmpdir", self.tmpdir)
        return Workbook(self.outfile, options)

    def test_option_validation(self):
        """Invalid option combinations raise ValueError before temp files."""
        with self.assertRaises(ValueError):
            Workbook(
                self.outfile,
                {"shared_strings": "disk"},
            )

        with self.assertRaises(ValueError):
            Workbook(
                self.outfile,
                {
                    "shared_strings": "disk",
                    "constant_memory": False,
                    "in_memory": False,
                },
            )

        with self.assertRaises(ValueError):
            Workbook(
                self.outfile,
                {
                    "shared_strings": "disk",
                    "constant_memory": True,
                    "in_memory": True,
                },
            )

        with self.assertRaises(ValueError):
            Workbook(
                self.outfile,
                {"shared_strings": "bogus", "constant_memory": True},
            )

        # A valid combination shouldn't raise.
        wb = self._workbook()
        wb.add_worksheet().write(0, 0, "ok")
        wb.close()

    def test_invalid_option_leaves_no_temp_files(self):
        """An invalid combination creates no files in the temp directory."""
        before = set(os.listdir(self.tmpdir))
        with self.assertRaises(ValueError):
            Workbook(
                os.path.join(self.tmpdir, "never.xlsx"),
                {"shared_strings": "disk", "constant_memory": False},
            )
        after = set(os.listdir(self.tmpdir))
        self.assertEqual(before, after)

    def test_no_option_keeps_original_behavior(self):
        """Without the option constant_memory still writes inline strings."""
        wb = Workbook(self.outfile, {"constant_memory": True, "tmpdir": self.tmpdir})
        ws = wb.add_worksheet()
        ws.write_string(0, 0, "inline me")
        wb.close()

        members = _read_zip_members(self.outfile)
        self.assertNotIn("xl/sharedStrings.xml", members)
        sheet = members["xl/worksheets/sheet1.xml"].decode("utf-8")
        self.assertIn('t="inlineStr"', sheet)

    def test_shared_strings_basic_and_dedup(self):
        """Plain strings de-duplicate by text and keep SST indices stable."""
        # Record the source values first.
        values = ["alpha", "beta", "alpha", "gamma", "beta"]
        wb = self._workbook()
        ws = wb.add_worksheet()
        for row, value in enumerate(values):
            ws.write_string(row, 0, value)
        wb.close()

        members = _read_zip_members(self.outfile)
        self.assertIn("xl/sharedStrings.xml", members)

        strings = _sst_strings(members["xl/sharedStrings.xml"])
        self.assertEqual([s[1] for s in strings], ["alpha", "beta", "gamma"])

        cells = _sheet_cells(members["xl/worksheets/sheet1.xml"])
        # Rebuild what the cells should reference from the recorded source.
        first_index = {}
        for i, value in enumerate(["alpha", "beta", "gamma"]):
            first_index[value] = i
        for row, value in enumerate(values):
            self.assertEqual(cells[(row, 0)], ("sst", first_index[value]))

        root = ET.fromstring(members["xl/sharedStrings.xml"])
        self.assertEqual(root.get("count"), str(len(values)))
        self.assertEqual(root.get("uniqueCount"), "3")

    def test_cross_sheet_dedup(self):
        """The same string on two sheets shares one SST entry."""
        wb = self._workbook()
        ws1 = wb.add_worksheet("One")
        ws2 = wb.add_worksheet("Two")
        ws1.write_string(0, 0, "shared")
        ws2.write_string(0, 0, "shared")
        ws1.write_string(1, 0, "only one")
        wb.close()

        members = _read_zip_members(self.outfile)
        strings = _sst_strings(members["xl/sharedStrings.xml"])
        self.assertEqual([s[1] for s in strings], ["shared", "only one"])

        cells1 = _sheet_cells(members["xl/worksheets/sheet1.xml"])
        cells2 = _sheet_cells(members["xl/worksheets/sheet2.xml"])
        self.assertEqual(cells1[(0, 0)], ("sst", 0))
        self.assertEqual(cells2[(0, 0)], ("sst", 0))
        self.assertEqual(cells1[(1, 0)], ("sst", 1))

    def test_overwrite_before_flush(self):
        """A cell overwritten before its row flushes keeps only the final."""
        wb = self._workbook()
        ws = wb.add_worksheet()
        ws.write_string(0, 0, "first")
        ws.write_string(0, 0, "second")
        ws.write_string(1, 0, "kept")
        wb.close()

        members = _read_zip_members(self.outfile)
        cells = _sheet_cells(members["xl/worksheets/sheet1.xml"])
        strings = _sst_strings(members["xl/sharedStrings.xml"])
        # The final cell values are "second" and "kept".
        self.assertEqual(strings[cells[(0, 0)][1]][1], "second")
        self.assertEqual(strings[cells[(1, 0)][1]][1], "kept")

        # "first" is an unreferenced entry: allowed. count must match the
        # two retained string cells.
        root = ET.fromstring(members["xl/sharedStrings.xml"])
        self.assertEqual(root.get("count"), "2")
        self.assertEqual(root.get("uniqueCount"), str(len(strings)))

    def test_rich_string_format_identity(self):
        """Rich strings dedupe on ordered fragments and effective format."""
        wb = self._workbook()
        ws = wb.add_worksheet()
        bold = wb.add_format({"bold": True})
        bold_again = wb.add_format({"bold": True})
        italic = wb.add_format({"italic": True})

        # Record the source writes.
        ws.write_rich_string(0, 0, "a", bold, "b")
        ws.write_rich_string(1, 0, "a", bold_again, "b")  # Equal format.
        ws.write_rich_string(2, 0, "a", italic, "b")  # Different format.
        ws.write_rich_string(3, 0, "a", bold, "c")  # Different text.
        ws.write_rich_string(4, 0, "x", wb.add_format({"font_color": "red"}), "y")
        ws.write_rich_string(5, 0, "x", wb.add_format({"font_color": "red"}), "y")
        ws.write_rich_string(6, 0, "x", wb.add_format({"bg_color": "red"}), "y")
        wb.close()

        members = _read_zip_members(self.outfile)
        cells = _sheet_cells(members["xl/worksheets/sheet1.xml"])
        runs = _sst_rich_runs(members["xl/sharedStrings.xml"])

        idx = {k: v[1] for k, v in cells.items()}
        # Equal formats must share an index.
        self.assertEqual(idx[(0, 0)], idx[(1, 0)])
        # Different formats and text must not.
        self.assertNotEqual(idx[(0, 0)], idx[(2, 0)])
        self.assertNotEqual(idx[(0, 0)], idx[(3, 0)])
        # Equal font color formats share an entry, and differ from a
        # differently formatted but identical-text entry.
        self.assertEqual(idx[(4, 0)], idx[(5, 0)])
        self.assertNotEqual(idx[(4, 0)], idx[(0, 0)])
        # A cell-only background color doesn't affect the rich text runs, so
        # it is a separate entry from the red-font run with the same text.
        self.assertNotEqual(idx[(6, 0)], idx[(4, 0)])
        self.assertNotEqual(idx[(6, 0)], idx[(0, 0)])
        self.assertEqual(len(set(idx.values())), 5)
        self.assertEqual(len(runs), 5)
        # Check the fragment text of each unique entry.
        texts = sorted("".join(t for _, t in r) for r in runs)
        self.assertEqual(texts, ["ab", "ab", "ac", "xy", "xy"])

    def test_plain_text_looking_like_rich(self):
        """Plain text that looks like <r>..</r> stays a plain string."""
        wb = self._workbook()
        ws = wb.add_worksheet()
        ws.write_string(0, 0, "<r>not rich</r>")
        ws.write_string(1, 0, "<r>")
        wb.close()

        members = _read_zip_members(self.outfile)
        strings = _sst_strings(members["xl/sharedStrings.xml"])
        self.assertTrue(all(not is_rich for is_rich, _ in strings))
        self.assertEqual([text for _, text in strings], ["<r>not rich</r>", "<r>"])

        cells = _sheet_cells(members["xl/worksheets/sheet1.xml"])
        self.assertEqual(strings[cells[(0, 0)][1]][1], "<r>not rich</r>")

    def test_special_characters(self):
        """Whitespace, Chinese, controls and Excel escape lookalikes."""
        wb = self._workbook()
        ws = wb.add_worksheet()
        source = {
            0: " leading",
            1: "trailing ",
            2: "  both  ",
            3: "中文文字",
            4: "with\x01control",
            5: "_x0001_",
            6: "_x005F_x0001_",
        }
        for row, value in source.items():
            ws.write_string(row, 0, value)
        wb.close()

        members = _read_zip_members(self.outfile)
        sst_xml = members["xl/sharedStrings.xml"].decode("utf-8")
        root = ET.fromstring(sst_xml)
        si_list = root.findall("m:si", NS)

        # Leading/trailing whitespace is preserved with xml:space.
        for i in (0, 1, 2):
            t = si_list[i].find("m:t", NS)
            self.assertEqual(t.get("{http://www.w3.org/XML/1998/namespace}space"),
                             "preserve")
        self.assertEqual(si_list[3].find("m:t", NS).text, "中文文字")
        # Control chars and Excel escape lookalikes are escaped as before.
        self.assertEqual(si_list[4].find("m:t", NS).text, "with_x0001_control")
        self.assertEqual(si_list[5].find("m:t", NS).text, "_x005F_x0001_")
        self.assertEqual(si_list[6].find("m:t", NS).text, "_x005F_x005F_x0001_")

        # Every cell references the entry recorded for its source value.
        cells = _sheet_cells(members["xl/worksheets/sheet1.xml"])
        for row in source:
            self.assertEqual(cells[(row, 0)], ("sst", row))

    def test_write_api_string_sources(self):
        """write(), write_string, rich strings and hyperlink text all use SST."""
        wb = self._workbook()
        ws = wb.add_worksheet()
        bold = wb.add_format({"bold": True})
        ws.write(0, 0, "via write")
        ws.write_string(1, 0, "via write_string")
        ws.write_rich_string(2, 0, "ri", bold, "ch")
        ws.write_url(3, 0, "https://example.com", None, "link text")
        wb.close()

        members = _read_zip_members(self.outfile)
        strings = _sst_strings(members["xl/sharedStrings.xml"])
        texts = [t for _, t in strings]
        self.assertIn("via write", texts)
        self.assertIn("via write_string", texts)
        self.assertIn("rich", texts)
        self.assertIn("link text", texts)

        cells = _sheet_cells(members["xl/worksheets/sheet1.xml"])
        for row in range(4):
            self.assertEqual(cells[(row, 0)][0], "sst")

    def test_non_string_cells_unchanged(self):
        """Numbers, dates, booleans, blanks and formulas keep their meaning."""
        import datetime

        wb = self._workbook()
        ws = wb.add_worksheet()
        red = wb.add_format({"bg_color": "red"})
        ws.write_number(0, 0, 3.5)
        ws.write_datetime(1, 0, datetime.datetime(2020, 1, 1))
        ws.write_boolean(2, 0, True)
        ws.write_blank(3, 0, None, red)
        ws.write_formula(4, 0, "=1+1", None, 2)
        wb.close()

        members = _read_zip_members(self.outfile)
        cells = _sheet_cells(members["xl/worksheets/sheet1.xml"])
        self.assertEqual(cells[(0, 0)], ("n", "3.5"))
        self.assertEqual(cells[(1, 0)], ("n", "43831"))
        self.assertEqual(cells[(2, 0)], ("b", "1"))
        self.assertEqual(cells[(3, 0)], ("blank", None))
        self.assertEqual(cells[(4, 0)][0], "formula")
        self.assertEqual(cells[(4, 0)][1], "2")

    def test_chart_caches_after_flush(self):
        """Chart caches are correct for rows already flushed to disk."""
        from xlsxwriter import Workbook as WB

        categories = ["cat A", "  padded  ", "", "cat D", "cat E"]
        values = [10, 20, 30, 40, 50]

        wb = WB(
            self.outfile,
            {
                "constant_memory": True,
                "shared_strings": "disk",
                "tmpdir": self.tmpdir,
            },
        )
        ws1 = wb.add_worksheet("Data")
        ws2 = wb.add_worksheet("Other")
        bold = wb.add_format({"bold": True})

        for row in range(5):
            if row == 2:
                # Leave B3 missing to exercise a gap.
                ws1.write_row(row, 0, [categories[row]])
            else:
                ws1.write_row(row, 0, [categories[row], values[row]])

        # Cross-sheet numeric series with a formula cache in between.
        ws2.write_string(0, 0, "x")
        ws2.write_number(0, 1, 1)
        ws2.write_formula(1, 0, "=B1*2", None, 2)
        ws2.write_number(1, 1, 2)
        ws2.write_string(2, 0, "rich category")
        ws2.write_rich_string(3, 0, "ri", bold, "ch")

        # All referenced rows are now flushed out of memory.
        ws2.write_number(10, 0, 0)

        chart = wb.add_chart({"type": "column"})
        chart.add_series(
            {
                "categories": "=Data!$A$1:$A$5",
                "values": "=Data!$B$1:$B$5",
            }
        )
        chart2 = wb.add_chart({"type": "line"})
        chart2.add_series(
            {
                "categories": "=Other!$A$1:$A$4",
                "values": "=Other!$B$1:$B$2",
            }
        )
        ws1.insert_chart("D2", chart)
        ws2.insert_chart("D2", chart2)
        wb.close()

        members = _read_zip_members(self.outfile)
        cs = _chart_cache(members["xl/charts/chart1.xml"])
        self.assertEqual(cs[0]["cat_kind"], "str")
        # The source categories are recorded by the test, expected at their
        # original positions (idx). Row 3 was written without a category so
        # it has no cache point, matching the in-memory behavior.
        expected_cat = {i: categories[i] for i in range(5) if i != 2}
        self.assertEqual(cs[0]["cat"], expected_cat)
        # B3 is missing so it isn't emitted, but the indices are preserved.
        self.assertEqual(cs[0]["val"], {0: "10", 1: "20", 3: "40", 4: "50"})

        c2 = _chart_cache(members["xl/charts/chart2.xml"])
        self.assertEqual(c2[0]["cat"][0], "x")
        self.assertEqual(c2[0]["cat"][2], "rich category")
        self.assertEqual(c2[0]["cat"][3], "rich")
        self.assertEqual(c2[0]["val"], {0: "1", 1: "2"})

    def test_chart_cross_sheet_and_escape_like(self):
        """Cross-sheet string categories and escape-like text in charts."""
        wb = self._workbook()
        ws1 = wb.add_worksheet("S One")
        ws2 = wb.add_worksheet("S Two")
        labels = ["a b", "_x0001_", "pad ", "中文"]
        for i, label in enumerate(labels):
            ws1.write_string(i, 0, label)
            ws1.write_number(i, 1, i + 1)
        ws2.write_number(0, 0, 0)

        chart = wb.add_chart({"type": "column"})
        chart.add_series(
            {
                "categories": "='S One'!$A$1:$A$4",
                "values": "='S One'!$B$1:$B$4",
            }
        )
        ws2.insert_chart("D2", chart)
        wb.close()

        members = _read_zip_members(self.outfile)
        cs = _chart_cache(members["xl/charts/chart1.xml"])
        # The chart cache stores the same escaped text that appears in the
        # shared string table, i.e. the on-disk representation of the source
        # values.
        expected = ["a b", "_x005F_x0001_", "pad ", "中文"]
        self.assertEqual(cs[0]["cat"], {i: expected[i] for i in range(4)})
        self.assertEqual(cs[0]["val"], {0: "1", 1: "2", 2: "3", 3: "4"})

    def test_chart_row_existence_outside_range(self):
        """Rows that exist only outside the column range suppress gaps."""
        wb = self._workbook()
        ws = wb.add_worksheet("S")
        # Rows 0 and 3 only have a cell in column F, outside the chart range.
        ws.write_number(0, 5, 99)
        ws.write_number(1, 0, 5)
        ws.write_number(3, 5, 1)
        ws.write_number(10, 0, 0)

        chart = wb.add_chart({"type": "column"})
        chart.add_series(
            {
                "categories": "=S!$A$1:$A$4",
                "values": "=S!$B$1:$B$4",
            }
        )
        ws.insert_chart("H2", chart)
        wb.close()

        members = _read_zip_members(self.outfile)
        cs = _chart_cache(members["xl/charts/chart1.xml"])
        # Only A2 has a value in the range; rows 1 and 4 "exist" because they
        # had a cell elsewhere, so the in-memory lookup emits None for those
        # cells rather than collapsing the whole row.
        self.assertEqual(cs[0]["cat"], {1: "5"})

    def test_chart_uses_final_buffered_row(self):
        """A chart can see the last row still buffered before close()."""
        wb = self._workbook()
        ws = wb.add_worksheet("Last")
        ws.write_string(0, 0, "never flushed")
        ws.write_number(0, 1, 7)
        chart = wb.add_chart({"type": "column"})
        chart.add_series(
            {
                "categories": "=Last!$A$1",
                "values": "=Last!$B$1",
            }
        )
        ws.insert_chart("D2", chart)
        wb.close()

        members = _read_zip_members(self.outfile)
        cs = _chart_cache(members["xl/charts/chart1.xml"])
        self.assertEqual(cs[0]["cat"], {0: "never flushed"})
        self.assertEqual(cs[0]["val"], {0: "7"})

    def test_relationships_and_part_presence(self):
        """sharedStrings.xml is declared in rels and content types."""
        wb = self._workbook()
        wb.add_worksheet().write_string(0, 0, "x")
        wb.close()

        members = _read_zip_members(self.outfile)
        ct = members["[Content_Types].xml"].decode("utf-8")
        self.assertIn("/sharedStrings", ct)

        rels = members["xl/_rels/workbook.xml.rels"].decode("utf-8")
        self.assertIn("sharedStrings.xml", rels)
        self.assertIn("relationships/sharedStrings", rels)

    def test_rich_and_plain_with_same_display_are_distinct(self):
        """A rich string and plain text with the same display are separate."""
        wb = self._workbook()
        ws1 = wb.add_worksheet("A")
        ws2 = wb.add_worksheet("B")
        ws1.write_rich_string(0, 0, "xy", wb.add_format({"bold": True}), "z")
        ws2.write_string(0, 0, "xyz")
        # Equal rich content with an independently created equal format
        # object must share an entry, even across sheets.
        ws1.write_rich_string(1, 0, "x", wb.add_format({"bold": True}), "y")
        ws2.write_rich_string(0, 1, "x", wb.add_format({"bold": True}), "y")
        wb.close()

        members = _read_zip_members(self.outfile)
        strings = _sst_strings(members["xl/sharedStrings.xml"])
        cells1 = _sheet_cells(members["xl/worksheets/sheet1.xml"])
        cells2 = _sheet_cells(members["xl/worksheets/sheet2.xml"])

        # "xyz" plain and the rich "xyz" are different entries.
        self.assertNotEqual(cells1[(0, 0)][1], cells2[(0, 0)][1])
        self.assertTrue(strings[cells1[(0, 0)][1]][0])  # rich
        self.assertFalse(strings[cells2[(0, 0)][1]][0])  # plain
        # Equal rich content shares an index.
        self.assertEqual(cells1[(1, 0)][1], cells2[(0, 1)][1])

    def test_count_matches_retained_string_cells(self):
        """SST count equals the number of retained string cells."""
        wb = self._workbook()
        ws = wb.add_worksheet()
        red = wb.add_format({"bg_color": "red"})

        # Overwritten before flush: no string cell retained.
        ws.write_string(0, 0, "gone")
        ws.write_number(0, 0, 1)
        ws.write_string(1, 0, "gone2")
        ws.write_blank(1, 0, None, red)
        # Retained string cells.
        ws.write_string(2, 0, "keep")
        ws.write_string(2, 1, "keep")
        ws.write_string(3, 0, "other")
        ws.write_number(5, 0, 0)
        wb.close()

        members = _read_zip_members(self.outfile)
        sheet_xml = members["xl/worksheets/sheet1.xml"].decode("utf-8")
        string_cells = sheet_xml.count('t="s"')
        root = ET.fromstring(members["xl/sharedStrings.xml"])
        self.assertEqual(int(root.get("count")), string_cells)
        self.assertEqual(string_cells, 3)
        # uniqueCount is the real number of si elements.
        self.assertEqual(int(root.get("uniqueCount")), len(_sst_strings(
            members["xl/sharedStrings.xml"])))

    def test_interleaved_workbooks_isolated(self):
        """Two interleaved workbooks keep their own numbering and temps."""
        tmp_a = tempfile.mkdtemp(prefix="xlsxwriter_a_", dir=self.tmpdir)
        tmp_b = tempfile.mkdtemp(prefix="xlsxwriter_b_", dir=self.tmpdir)
        out_a = os.path.join(self.tmpdir, "a.xlsx")
        out_b = os.path.join(self.tmpdir, "b.xlsx")

        wb_a = Workbook(
            out_a,
            {"constant_memory": True, "shared_strings": "disk", "tmpdir": tmp_a},
        )
        wb_b = Workbook(
            out_b,
            {"constant_memory": True, "shared_strings": "disk", "tmpdir": tmp_b},
        )
        ws_a = wb_a.add_worksheet()
        ws_b = wb_b.add_worksheet()

        ws_a.write_string(0, 0, "aaa")
        ws_b.write_string(0, 0, "bbb")
        ws_a.write_string(1, 0, "ccc")
        ws_b.write_string(1, 0, "aaa")
        wb_b.close()
        ws_a.write_string(2, 0, "bbb")
        wb_a.close()

        members_a = _read_zip_members(out_a)
        members_b = _read_zip_members(out_b)
        strings_a = [t for _, t in _sst_strings(members_a["xl/sharedStrings.xml"])]
        strings_b = [t for _, t in _sst_strings(members_b["xl/sharedStrings.xml"])]
        self.assertEqual(strings_a, ["aaa", "ccc", "bbb"])
        self.assertEqual(strings_b, ["bbb", "aaa"])

        # Temp directories are emptied of the store files.
        self.assertEqual(
            [n for n in os.listdir(tmp_a) if n.startswith("xlsxwriter_sst_")], []
        )
        self.assertEqual(
            [n for n in os.listdir(tmp_b) if n.startswith("xlsxwriter_sst_")], []
        )


class DiskSharedStringsFailureTest(unittest.TestCase):
    """Failure injection tests for the disk backed export path."""

    def setUp(self):
        self.tmpdir = tempfile.mkdtemp(prefix="xlsxwriter_fail_")
        self.outfile = os.path.join(self.tmpdir, "existing.xlsx")

    def tearDown(self):
        for root, _dirs, names in os.walk(self.tmpdir):
            for name in names:
                try:
                    os.remove(os.path.join(root, name))
                except OSError:
                    pass
        for root, dirs, _names in os.walk(self.tmpdir, topdown=False):
            for name in dirs:
                try:
                    os.rmdir(os.path.join(root, name))
                except OSError:
                    pass
        try:
            os.rmdir(self.tmpdir)
        except OSError:
            pass

    def test_write_failure_leaves_existing_file(self):
        """A staging write failure leaves the pre-existing output intact."""
        original = b"pre-existing, not an xlsx"
        with open(self.outfile, "wb") as fh:
            fh.write(original)

        wb = Workbook(
            self.outfile,
            {"constant_memory": True, "shared_strings": "disk",
             "tmpdir": self.tmpdir},
        )
        ws = wb.add_worksheet()
        ws.write_string(0, 0, "hello")

        # Make the shared string store unwritable mid-stream.
        wb.str_table.db.execute(
            "CREATE TRIGGER fail_insert BEFORE INSERT ON strings "
            "BEGIN SELECT RAISE(ABORT, 'injected failure'); END"
        )

        # A new string now fails while being staged.
        with self.assertRaises(Exception):
            ws.write_string(5, 0, "a new unique string")

        # close() must not report success after the staging failure.
        with self.assertRaises(Exception):
            wb.close()

        with open(self.outfile, "rb") as fh:
            self.assertEqual(fh.read(), original)

        # No temp store files are left behind.
        leftovers = [
            n for n in os.listdir(self.tmpdir)
            if n.startswith("xlsxwriter_sst_") or n.startswith(".xlsxwriter_")
        ]
        self.assertEqual(leftovers, [])

    def test_zip_failure_leaves_existing_file(self):
        """A ZIP packaging failure leaves the pre-existing output intact."""
        import xlsxwriter.workbook as wbmod

        original = b"pre-existing"
        with open(self.outfile, "wb") as fh:
            fh.write(original)

        real_zipfile = wbmod.ZipFile

        class FailingZipFile(real_zipfile):
            def write(self, filename, arcname=None, compress_type=None):
                raise OSError("injected ZIP failure")

        wbmod.ZipFile = FailingZipFile
        try:
            wb = Workbook(
                self.outfile,
                {"constant_memory": True, "shared_strings": "disk",
                 "tmpdir": self.tmpdir},
            )
            ws = wb.add_worksheet()
            ws.write_string(0, 0, "hello")
            with self.assertRaises(Exception):
                wb.close()
        finally:
            wbmod.ZipFile = real_zipfile

        with open(self.outfile, "rb") as fh:
            self.assertEqual(fh.read(), original)

        leftovers = [
            n for n in os.listdir(self.tmpdir)
            if n.startswith("xlsxwriter_sst_") or n.startswith(".xlsxwriter_")
        ]
        self.assertEqual(leftovers, [])


if __name__ == "__main__":
    unittest.main()
