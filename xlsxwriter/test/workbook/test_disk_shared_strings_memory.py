###############################################################################
#
# A small, fast regression test that the disk backed shared string export
# keeps a bounded resident memory footprint as the number of distinct strings
# grows.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#
"""
This runs two exports in isolated subprocesses and compares peak RSS.

The full-size experiment (100k vs 1M strings) is in
``dev/disk_sst/run_experiment.py``. This test uses a smaller scale so it is
fast enough for the regular test suite, and asserts a generous margin.
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
    import io, resource, sys
    sys.path.insert(0, sys.argv[1])
    from xlsxwriter import Workbook

    count = int(sys.argv[2])
    tmpdir = sys.argv[3]
    out = sys.argv[4]

    wb = Workbook(
        out,
        {"constant_memory": True, "shared_strings": "disk", "tmpdir": tmpdir},
    )
    ws1 = wb.add_worksheet("One")
    ws2 = wb.add_worksheet("Two")
    half = count // 2
    for i in range(half):
        ws1.write_string(i, 0, f"{i:010d}" + "x" * 502)
        ws2.write_string(i, 0, f"{i + count:010d}" + "x" * 502)
    wb.close()
    print(resource.getrusage(resource.RUSAGE_SELF).ru_maxrss)
    """
)


def _run(count, workdir):
    tmpdir = os.path.join(workdir, "tmp_" + str(count))
    os.makedirs(tmpdir, exist_ok=True)
    out = os.path.join(workdir, "out_" + str(count) + ".xlsx")
    result = subprocess.run(
        [sys.executable, "-c", CHILD, REPO_ROOT, str(count), tmpdir, out],
        capture_output=True,
        text=True,
        check=True,
    )
    return int(result.stdout.strip().splitlines()[-1])


class DiskSharedStringsMemoryTest(unittest.TestCase):
    """Bounded memory as the number of distinct strings grows."""

    def test_peak_rss_is_flat(self):
        workdir = tempfile.mkdtemp(prefix="xlsxwriter_mem_")
        try:
            small = _run(20000, workdir)
            large = _run(150000, workdir)
        finally:
            for root, _dirs, names in os.walk(workdir, topdown=False):
                for name in names:
                    try:
                        os.remove(os.path.join(root, name))
                    except OSError:
                        pass
                for name in _dirs:
                    try:
                        os.rmdir(os.path.join(root, name))
                    except OSError:
                        pass
            try:
                os.rmdir(workdir)
            except OSError:
                pass

        # 7.5x the strings should not add more than 16 MiB of resident memory.
        delta_mib = (large - small) / 1024
        self.assertLess(
            delta_mib,
            16,
            f"peak RSS grew by {delta_mib:.2f} MiB "
            f"(small={small} KiB, large={large} KiB)",
        )


if __name__ == "__main__":
    unittest.main()
