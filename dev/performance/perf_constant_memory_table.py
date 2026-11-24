#!/usr/bin/env python3
###############################################################################
#
# Measure the peak RSS of writing a single worksheet table in
# 'constant_memory' mode for a given number of rows.
#
# The table has 3 columns, one calculated column and a totals row. Each run
# is an independent process so the reported peak RSS is for that run only.
#
# Usage:
#
#     python3 dev/performance/perf_constant_memory_table.py 100000 1000000
#
# The two (or more) figures can be compared to check that peak memory is
# flat with respect to the number of rows. For example:
#
#     python3 dev/performance/perf_constant_memory_table.py 100000 1000000
#     100000 rows: 18.3 MiB peak RSS
#     1000000 rows: 18.3 MiB peak RSS
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c) 2013-2025, John McNamara, jmcnamara@cpan.org
#

import os
import subprocess
import sys
import tempfile

_SCRIPT = """
import os
import resource
import sys

sys.path.insert(0, sys.argv[3])

import xlsxwriter

row_count = int(sys.argv[1])
path = sys.argv[2]

workbook = xlsxwriter.Workbook(path, {"constant_memory": True, "in_memory": False})
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

print(resource.getrusage(resource.RUSAGE_SELF).ru_maxrss)
"""


def measure(row_count):
    # Run a single measurement in a fresh, isolated process and return the
    # peak RSS in bytes.
    package_root = os.path.dirname(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    )

    with tempfile.TemporaryDirectory() as tmpdir:
        path = os.path.join(tmpdir, "table.xlsx")
        output = subprocess.check_output(
            [
                sys.executable,
                "-I",
                "-c",
                _SCRIPT,
                str(row_count),
                path,
                package_root,
            ]
        )

    # ru_maxrss is in KiB on Linux.
    return int(output.split()[-1]) * 1024


def main(argv):
    row_counts = [int(arg) for arg in argv] or [100_000, 1_000_000]

    for row_count in row_counts:
        peak = measure(row_count)
        print(f"{row_count} rows: {peak / (1024 * 1024):.1f} MiB peak RSS")


if __name__ == "__main__":
    main(sys.argv[1:])
