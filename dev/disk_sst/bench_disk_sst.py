###############################################################################
#
# Memory experiment for the disk backed shared string export path.
#
# SPDX-License-Identifier: BSD-2-Clause
#
"""
Export two synthetic workbooks and measure peak RSS from the start of writing
to the completion of close().

  * Small: 100,000 distinct 512-character ASCII strings across two sheets.
  * Large: 1,000,000 distinct 512-character ASCII strings across two sheets.

Each run is a separate process. The check pass reads sharedStrings.xml
through a streaming parser so that the source data and the XML are never held
in memory at the same time.

Run:
    python3 dev/disk_sst/bench_disk_sst.py --rows 100000 --out /tmp/small.xlsx
    python3 dev/disk_sst/bench_disk_sst.py --rows 1000000 --out /tmp/large.xlsx

Compare the reported peak_rss_kib for each run. The requirement is that the
large run exceeds the small run by at most 64 MiB.
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


def make_text(seed, size):
    # A deterministic, unique, ASCII string of the requested size.
    prefix = f"{seed:010d}:"
    filler = "abcdefghijklmnopqrstuvwxyz0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    body = (filler * ((size // len(filler)) + 1))[: size - len(prefix)]
    return prefix + body


def stream_check(path, expected_unique):
    """Independently stream-parse the SST to count <si> and <t> elements."""
    si_count = 0
    t_chars = 0
    with zipfile.ZipFile(path) as zf:
        info = zf.getinfo("xl/sharedStrings.xml")
        with zf.open("xl/sharedStrings.xml") as fh:
            # Use an incremental XML parser so only a small buffer is live.
            parser = ET.XMLPullParser(events=("end",))
            total = 0
            while True:
                chunk = fh.read(1 << 16)
                if not chunk:
                    break
                total += len(chunk)
                parser.feed(chunk)
                for _event, elem in parser.read_events():
                    if elem.tag == f"{{{NS_MAIN}}}si":
                        si_count += 1
                    elif elem.tag == f"{{{NS_MAIN}}}t":
                        t_chars += len(elem.text or "")
                        elem.clear()
            parser.close()
    assert total == info.file_size, (total, info.file_size)
    assert si_count == expected_unique, (si_count, expected_unique)
    return si_count, t_chars


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--rows", type=int, default=100000)
    p.add_argument("--size", type=int, default=512)
    p.add_argument("--out", required=True)
    p.add_argument("--tmpdir", default=None)
    p.add_argument("--skip-check", action="store_true")
    args = p.parse_args()

    total = args.rows
    per_sheet = total // 2
    remainder = total - per_sheet * 2

    # Peak RSS from process start, which is before the workbook is created.
    t0 = time.time()
    wb = xlsxwriter.Workbook(
        args.out,
        {
            "constant_memory": True,
            "shared_strings": "disk",
            "tmpdir": args.tmpdir,
        },
    )
    ws1 = wb.add_worksheet("One")
    ws2 = wb.add_worksheet("Two")

    seed = 0
    for row in range(per_sheet):
        # Interleave the two sheets row by row. The strings are all distinct.
        ws1.write_string(row, 0, make_text(seed, args.size))
        ws2.write_string(row, 0, make_text(seed + total, args.size))
        ws1.write_number(row, 1, row)
        ws2.write_number(row, 1, row)
        seed += 1

    for extra in range(remainder):
        ws1.write_string(per_sheet + extra, 0, make_text(total + extra, args.size))

    wb.close()
    elapsed = time.time() - t0

    peak_kib = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    print(
        f"rows={total} size={args.size} elapsed={elapsed:.1f}s "
        f"peak_rss_kib={peak_kib} out={args.out}"
    )

    if not args.skip_check:
        si_count, t_chars = stream_check(args.out, total)
        print(f"stream-check si_count={si_count} t_chars={t_chars}")
        print(f"file_size={os.path.getsize(args.out)}")


if __name__ == "__main__":
    main()
