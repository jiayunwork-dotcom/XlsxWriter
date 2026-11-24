###############################################################################
#
# Memory experiment for Worksheet.materialize_autofilter().
#
# SPDX-License-Identifier: BSD-2-Clause
#
"""
Export a workbook with a materialized autofilter and measure peak RSS from the
start of writing until close() returns.

The data is generated one row at a time and discarded. Each row has:

  * a filter column alternating 'East' / 'West',
  * a fixed 64 character text payload,
  * a finite number used for a numeric range filter.

Both filter conditions are set only after every row has been written. Two
charts reference the same contiguous 100 row range: one plots only visible
data and one uses show_hidden_data().

The check pass stream-parses the worksheet XML in 64 KiB chunks and verifies
every row against the generation rule. It never holds the source table or the
whole extracted XML in memory.

Run a single scale:
    python3 dev/autofilter/bench_materialize_autofilter.py --rows 100000 \
        --out /tmp/small.xlsx
"""

import argparse
import os
import resource
import sys
import time
import zipfile
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..")))

import xlsxwriter  # noqa: E402

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart"

# The preferred contiguous chart range: worksheet rows 3001..3100, i.e. data
# rows 3000..3099. Their values (3001..3100) all fall inside the numeric
# filter, so about half of them remain visible and half are filtered out. For
# smaller smoke runs the range is moved back to fit in the data.
CHART_FIRST_ROW = 3000
CHART_ROWS = 100

PAYLOAD_SIZE = 64

# Text filter: keep 'East'. Numeric filter: keep 2500 <= value < 7500.
TEXT_VALUE = "East"
NUM_LOW = 2500
NUM_HIGH = 7500


def region_for(row):
    # Row 0 is the header, data rows start at 1.
    return "East" if row % 2 == 0 else "West"


def number_for(row):
    return (row % 10000) + 1


def payload_for(row):
    # A deterministic, unique, fixed length ASCII string.
    prefix = f"{row:010d}:"
    filler = "abcdefghijklmnopqrstuvwxyz0123456789"
    return (prefix + filler * PAYLOAD_SIZE)[:PAYLOAD_SIZE]


def expected_hidden(row):
    # Straight from the documented conditions: region must be East and the
    # number must be in [2500, 7500).
    return not (region_for(row) == TEXT_VALUE and NUM_LOW <= number_for(row) < NUM_HIGH)


def chart_first_row(rows):
    # Keep the 100 row chart range inside the data for small smoke runs.
    return min(CHART_FIRST_ROW, max(0, rows - CHART_ROWS - 1))


def build(path, rows, tmpdir):
    start = time.time()
    wb = xlsxwriter.Workbook(
        path,
        {"constant_memory": True, "shared_strings": "disk", "tmpdir": tmpdir},
    )
    ws = wb.add_worksheet("Data")
    ws.autofilter(0, 0, rows, 2)
    ws.write_row(0, 0, ["Region", "Payload", "Value"])

    for row in range(1, rows + 1):
        ws.write_string(row, 0, region_for(row))
        ws.write_string(row, 1, payload_for(row))
        ws.write_number(row, 2, number_for(row))

    # The conditions are set only after all of the data has been written.
    ws.filter_column("A", "x == East")
    ws.filter_column("C", f"x >= {NUM_LOW} and x < {NUM_HIGH}")
    ws.materialize_autofilter()

    first_row = chart_first_row(rows)
    visible = wb.add_chart({"type": "column"})
    visible.add_series(
        {
            "categories": f"=Data!$A${first_row + 1}:"
            f"$A${first_row + CHART_ROWS}",
            "values": f"=Data!$C${first_row + 1}:"
            f"$C${first_row + CHART_ROWS}",
        }
    )
    hidden = wb.add_chart({"type": "line"})
    hidden.add_series(
        {
            "categories": f"=Data!$A${first_row + 1}:"
            f"$A${first_row + CHART_ROWS}",
            "values": f"=Data!$C${first_row + 1}:"
            f"$C${first_row + CHART_ROWS}",
        }
    )
    hidden.show_hidden_data()
    ws.insert_chart("F2", visible)
    ws.insert_chart("F20", hidden)

    wb.close()
    return time.time() - start


def _cell_refs(row_elem):
    # Extract (column letter, value text, type) for a row without keeping it.
    cells = []
    for cell in row_elem:
        if not cell.tag.endswith("}c"):
            continue
        ref = cell.get("r")
        letters = "".join(ch for ch in ref if ch.isalpha())
        v = cell.find(f"{{{NS_MAIN}}}v")
        cells.append((letters, None if v is None else v.text, cell.get("t")))
    return cells


def stream_check(path, rows):
    """Verify every row's identity and hidden flag with a streaming parser.

    The row number, the A/B/C cell coordinates and the generated number in
    column C are checked against the generation rule. The text columns are
    only checked for presence because resolving their shared string indices
    would require the whole SST in memory.
    """
    checked = 0
    hidden_count = 0
    total_bytes = 0

    with zipfile.ZipFile(path) as zf:
        info = zf.getinfo("xl/worksheets/sheet1.xml")
        with zf.open("xl/worksheets/sheet1.xml") as fh:
            parser = ET.XMLPullParser(events=("end",))
            while True:
                chunk = fh.read(1 << 16)
                if not chunk:
                    break
                total_bytes += len(chunk)
                parser.feed(chunk)
                for _event, elem in parser.read_events():
                    if not elem.tag.endswith("}row"):
                        continue
                    row_num = int(elem.get("r"))
                    hidden = elem.get("hidden") == "1"
                    # Row 1 is the header and isn't filtered.
                    if row_num > 1:
                        data_row = row_num - 1
                        expected = expected_hidden(data_row)
                        if hidden != expected:
                            raise AssertionError(
                                f"row {row_num}: hidden={hidden} expected={expected}"
                            )

                        cells = _cell_refs(elem)
                        refs = [cell[0] for cell in cells]
                        if refs != ["A", "B", "C"]:
                            raise AssertionError(
                                f"row {row_num}: unexpected cells {refs}"
                            )
                        value = cells[2][1]
                        expected_value = f"{number_for(data_row):.16G}"
                        if value != expected_value:
                            raise AssertionError(
                                f"row {row_num}: value={value} "
                                f"expected={expected_value}"
                            )
                        if cells[0][2] != "s" or cells[1][2] != "s":
                            raise AssertionError(
                                f"row {row_num}: text cells aren't shared strings"
                            )

                        checked += 1
                        hidden_count += int(hidden)
                    elem.clear()
            parser.close()

    assert total_bytes == info.file_size, (total_bytes, info.file_size)
    # Every physical data row in the filter range must be present.
    assert checked == rows, (checked, rows)
    return checked, hidden_count


def chart_check(path, rows):
    """Verify the visible-only and show_hidden_data chart caches."""
    rows_in_range = CHART_ROWS
    first_row = chart_first_row(rows)
    # The chart range starts at worksheet row first_row + 1, which is zero
    # based data row first_row; the offset is the data row delta.
    visible_rows = [
        offset
        for offset in range(rows_in_range)
        if not expected_hidden(first_row + offset)
    ]
    expected_visible = set(visible_rows)
    expected_all = set(range(rows_in_range))

    with zipfile.ZipFile(path) as zf:
        for name, expected in (
            ("xl/charts/chart1.xml", expected_visible),
            ("xl/charts/chart2.xml", expected_all),
        ):
            root = ET.fromstring(zf.read(name))
            series = root.findall(f".//{{{NS_CHART}}}ser")
            assert len(series) == 1, (name, len(series))
            for tag in ("cat", "val"):
                element = series[0].find(f"{{{NS_CHART}}}{tag}")
                cache = element[0].find(f"{{{NS_CHART}}}"
                                       f"{'strCache' if tag == 'cat' else 'numCache'}")
                count = int(cache.find(f"{{{NS_CHART}}}ptCount").get("val"))
                idx = {int(pt.get("idx")) for pt in cache.findall(
                    f"{{{NS_CHART}}}pt")}
                assert count == rows_in_range, (name, tag, count)
                assert idx == expected, (name, tag, sorted(idx)[:5])
    return len(visible_rows)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--rows", type=int, default=100000)
    parser.add_argument("--out", required=True)
    parser.add_argument("--tmpdir", default=None)
    parser.add_argument("--skip-check", action="store_true")
    args = parser.parse_args()

    elapsed = build(args.out, args.rows, args.tmpdir)
    peak_kib = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    print(
        f"rows={args.rows} elapsed={elapsed:.1f}s "
        f"peak_rss_kib={peak_kib} out={args.out}"
    )

    if not args.skip_check:
        checked, hidden_count = stream_check(args.out, args.rows)
        visible_chart_points = chart_check(args.out, args.rows)
        print(
            f"stream-check rows={checked} hidden={hidden_count} "
            f"visible_chart_points={visible_chart_points} "
            f"file_size={os.path.getsize(args.out)}"
        )


if __name__ == "__main__":
    main()
