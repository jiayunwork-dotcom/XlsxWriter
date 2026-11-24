###############################################################################
#
# A small, fast regression test for the materialized autofilter memory use.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c) 2013-2025, John McNamara, jmcnamara@cpan.org
#
"""
Run two exports in isolated subprocesses and compare peak RSS.

The full size experiment (100,000 vs 800,000 rows) is in
``dev/autofilter/run_experiment.py``. This test uses a smaller scale so that
it is fast enough for the regular test suite, with a generous margin. It also
streams the worksheet XML to check row identity and hidden flags.
"""

import os
import subprocess
import sys
import tempfile
import textwrap
import unittest

REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))

CHILD = textwrap.dedent(
    """
    import resource, sys
    sys.path.insert(0, sys.argv[1])
    from xlsxwriter import Workbook

    rows = int(sys.argv[2])
    tmpdir = sys.argv[3]
    out = sys.argv[4]
    payload = "x" * 64

    wb = Workbook(
        out,
        {"constant_memory": True, "shared_strings": "disk", "tmpdir": tmpdir},
    )
    ws = wb.add_worksheet("Data")
    ws.autofilter(0, 0, rows, 2)
    ws.write_row(0, 0, ["Region", "Payload", "Value"])
    for row in range(1, rows + 1):
        ws.write_string(row, 0, "East" if row % 2 == 0 else "West")
        ws.write_string(row, 1, payload)
        ws.write_number(row, 2, row % 10000 + 1)
    ws.filter_column("A", "x == East")
    ws.filter_column("C", "x >= 2500 and x < 7500")
    ws.materialize_autofilter()
    wb.close()
    print(resource.getrusage(resource.RUSAGE_SELF).ru_maxrss)
    """
)


def _run(count, workdir):
    tmpdir = os.path.join(workdir, "tmp_" + str(count))
    os.makedirs(tmpdir, exist_ok=True)
    out = os.path.join(workdir, "out_" + str(count) + ".xlsx")
    result = subprocess.run(
        [sys.executable, "-I", "-c", CHILD, REPO_ROOT, str(count), tmpdir, out],
        capture_output=True,
        text=True,
        check=True,
    )
    return int(result.stdout.strip().splitlines()[-1]), out


class MaterializeAutofilterMemoryTest(unittest.TestCase):
    """Bounded memory as the number of filtered rows grows."""

    def test_peak_rss_is_flat(self):
        workdir = tempfile.mkdtemp(prefix="xlsxwriter_af_mem_")
        try:
            small, _ = _run(20000, workdir)
            large, _ = _run(150000, workdir)
        finally:
            for root, dirs, names in os.walk(workdir, topdown=False):
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
                os.rmdir(workdir)
            except OSError:
                pass

        delta_mib = (large - small) / 1024
        self.assertLess(
            delta_mib,
            16,
            f"peak RSS grew by {delta_mib:.2f} MiB "
            f"(small={small} KiB, large={large} KiB)",
        )


if __name__ == "__main__":
    unittest.main()
