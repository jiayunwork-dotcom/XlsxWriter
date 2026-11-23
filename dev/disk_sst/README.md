# Disk backed shared strings: resource experiment

This directory contains the memory experiment for the
`Workbook(shared_strings='disk')` export path.

## What is measured

For each scale, a single process:

1. Creates a workbook with `constant_memory=True` and `shared_strings='disk'`.
2. Writes two worksheets whose combined number of **distinct** 512-character
   ASCII strings is either 100,000 or 1,000,000. Cell formats are fixed (a
   couple of `add_format()` objects are reused).
3. Calls `close()`, which streams `sharedStrings.xml` out of the SQLite store,
   packages every part and writes the ZIP.
4. Reports `resource.getrusage(RUSAGE_SELF).ru_maxrss` (peak RSS in KiB) for
   the whole process lifetime, i.e. from before the write loop starts until
   after `close()` returns.

The two runs are separate processes so their peaks are independent.

The check pass then stream-parses `xl/sharedStrings.xml` with
`xml.etree.ElementTree.XMLPullParser` in 64 KiB chunks, so neither the source
data nor the whole XML/ZIP is ever live in memory during verification.

## Reproduce

From the repository root:

```sh
python3 dev/disk_sst/bench_disk_sst.py --rows 100000  --out /tmp/disk_sst_small.xlsx
python3 dev/disk_sst/bench_disk_sst.py --rows 1000000 --out /tmp/disk_sst_large.xlsx
python3 dev/disk_sst/run_experiment.py
```

`run_experiment.py` runs both scales in isolated subprocesses (optionally with
a dedicated `--tmpdir` per run) and prints the peak RSS delta.

## Recorded result

On the development machine used for this change:

| Scale   | Distinct strings | Elapsed | Peak RSS  |
| ------- | ---------------- | ------- | --------- |
| Small   | 100,000          | ~2.7 s  | 24,068 KiB |
| Large   | 1,000,000        | ~23 s   | 24,204 KiB |

Large minus small: **~136 KiB**, far below the 64 MiB limit. The peak stays
flat because the distinct strings live only in the SQLite file; the only
in-memory string state is a 512-entry LRU cache plus the current worksheet
row.

`stream_check()` confirms the written file contains exactly the expected
number of `<si>` elements (100,000 / 1,000,000) and 512 characters each.

## Notes

* The SQLite database is created under the workbook's `tmpdir` and removed when
  the workbook closes (success or failure).
* The same database stores the flushed cell values used to build chart caches.
  A chart range lookup reads only the requested cells, not the whole sheet.
