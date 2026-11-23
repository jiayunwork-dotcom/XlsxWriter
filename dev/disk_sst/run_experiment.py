###############################################################################
#
# Run the disk backed shared strings memory experiment in isolated processes.
#
# SPDX-License-Identifier: BSD-2-Clause
#
"""
Run the small (100,000 strings) and large (1,000,000 strings) exports each in
its own process and compare peak RSS.

The peak RSS of each run is measured by the child process itself (from before
the write loop until after close()). The parent only reads the printed value,
so its own memory use doesn't affect the measurement.
"""

import argparse
import os
import re
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
BENCH = os.path.join(HERE, "bench_disk_sst.py")

PEAK_RE = re.compile(r"peak_rss_kib=(\d+)")


def run(rows, out, tmpdir):
    cmd = [sys.executable, BENCH, "--rows", str(rows), "--out", out]
    if tmpdir:
        cmd += ["--tmpdir", tmpdir]
    result = subprocess.run(cmd, capture_output=True, text=True, check=True)
    print(result.stdout.strip())
    match = PEAK_RE.search(result.stdout)
    if not match:
        raise RuntimeError("Could not parse peak RSS from benchmark output")
    return int(match.group(1))


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--workdir", default=None)
    parser.add_argument("--limit-mib", type=int, default=64)
    args = parser.parse_args()

    workdir = args.workdir or tempfile.mkdtemp(prefix="disk_sst_experiment_")
    os.makedirs(workdir, exist_ok=True)

    small_tmp = os.path.join(workdir, "tmp_small")
    large_tmp = os.path.join(workdir, "tmp_large")
    os.makedirs(small_tmp, exist_ok=True)
    os.makedirs(large_tmp, exist_ok=True)

    print("=== Small: 100,000 distinct strings ===")
    small = run(100000, os.path.join(workdir, "small.xlsx"), small_tmp)

    print("=== Large: 1,000,000 distinct strings ===")
    large = run(1000000, os.path.join(workdir, "large.xlsx"), large_tmp)

    delta_mib = (large - small) / 1024
    print(
        f"\nPeak RSS small={small} KiB large={large} KiB "
        f"delta={delta_mib:.2f} MiB limit={args.limit_mib} MiB"
    )
    if delta_mib <= args.limit_mib:
        print("RESULT: PASS")
    else:
        print("RESULT: FAIL")
        sys.exit(1)


if __name__ == "__main__":
    main()
