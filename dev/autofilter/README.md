# Materialized autofilter: resource experiment

This directory contains the memory and correctness experiment for
`Worksheet.materialize_autofilter()`.

## What the benchmark does

For each scale, a single process:

1. Creates a workbook with `constant_memory=True` and
   `shared_strings='disk'`.
2. Writes one worksheet with the requested number of data rows. Every row has
   a region column alternating `'East'`/`'West'`, a fixed 64 character text
   payload and a finite number, generated one row at a time and discarded.
3. Sets a text list condition (`== East`) and a numeric range
   (`>= 2500 and < 7500`) only **after** all of the data has been written.
4. Calls `materialize_autofilter()`.
5. Adds two charts that both reference the same contiguous 100 row range, one
   plotting only visible data and one using `show_hidden_data()`.
6. Calls `close()` and reports `resource.getrusage(RUSAGE_SELF).ru_maxrss`.

The two runs happen in separate Python 3.11 processes, so their peak RSS
values are independent.

The check pass stream-parses `xl/worksheets/sheet1.xml` in 64 KiB chunks with
`xml.etree.ElementTree.XMLPullParser`, verifying every row's hidden flag
against the generation rule. Neither the source data nor the whole extracted
XML is ever live in memory.

## Reproduce

From the repository root:

```sh
python3 dev/autofilter/bench_materialize_autofilter.py --rows 100000 \
    --out /tmp/af_small.xlsx
python3 dev/autofilter/bench_materialize_autofilter.py --rows 800000 \
    --out /tmp/af_large.xlsx
python3 dev/autofilter/run_experiment.py
```

`run_experiment.py` runs both scales in isolated subprocesses and prints the
peak RSS delta. It exits non-zero if the delta exceeds 48 MiB.

The lightweight regression version (20,000 vs 150,000 rows, 16 MiB margin)
runs as part of the normal test suite:
`xlsxwriter/test/worksheet/test_materialize_autofilter_memory.py`.

## Recorded result

On the development machine for this change (Python 3.11.2, Linux):

| Scale   | Rows    | Elapsed | Peak RSS  | Hidden rows | Visible chart points |
| ------- | ------- | ------- | --------- | ----------- | -------------------- |
| Small   | 100,000 | ~2.4 s  | 24,800 KiB | 75,000     | 50 / 100             |
| Large   | 800,000 | ~16.8 s | 24,800 KiB | 600,000    | 50 / 100             |

Large minus small: **0.00 MiB** in the final run (earlier runs measured
-0.06 to +0.16 MiB), far below the 48 MiB limit. The visible-only chart keeps
50 of the 100 reference points at their original indices; the
`show_hidden_data` chart keeps all 100. The exact figures vary by run; the
point is that the delta stays flat.

## Resource cost of the chosen mechanism

* **Late conditions.** The filter is evaluated once, at close, over the
  flushed cells. Setting or replacing a condition after the data is written
  therefore costs nothing extra. Calling `materialize_autofilter()` early and
  adding a condition later costs the same as setting it last.
* **Sparse matching.** The evaluation streams every physical row in the range
  (one SQL cursor over the filtered columns) and writes only the hidden
  decisions. A mostly-matching table and a mostly-hidden table have the same
  scan cost; the difference is the number of rows written to the
  `filter_hidden` table. There is no early exit per row.
* **Two charts sharing a range.** The chart cache values are materialized
  once per referenced range and the visibility positions live on that shared
  cache object. Each chart applies its own `show_hidden_data` mode when it
  writes its cache XML, so the creation order doesn't matter and the second
  chart doesn't force another worksheet scan. Chart memory is proportional to
  the ranges that are actually referenced, not to the sheet.
* **Temporary disk.** The flushed cell records already existed for chart
  caches. The additions are the `filter_hidden` rows (one small row per
  hidden data row) and `filter_errors` positions (only for error cells). Both
  live in the existing SQLite file under `tmpdir`, use `WITHOUT ROWID`, and
  are removed with the rest of the store on close.

The peak stays flat because the filter evaluation streams the flushed cells
from the SQLite store in one ordered cursor (only one row of comparison values
is live at a time) and then streams the hidden row numbers back out. The
hidden decisions are stored in a `WITHOUT ROWID` table in the same store and
are never all held in memory. The chart caches only materialize the ranges
that the charts actually reference (100 points here).

## Notes

* The SQLite database is created under the workbook's `tmpdir` and removed
  when the workbook closes, on success or failure. A filter that can't be
  evaluated (unsupported condition, or a formula/boolean/error cell in a
  filtered column) raises before any output file is created, so a pre-existing
  target file is left untouched.
* The same database stores the flushed cell values used to build chart
  caches. A chart range lookup reads only the requested cells.
* Because the filter is evaluated at close time, the cost of a late condition
  is one extra sequential pass over the filtered columns. Rows that don't
  match early in the range are still scanned; there is no partial-scan
  optimization. The numeric column filter and the text column filter are
  evaluated per row from the same cursor, so adding another filtered column
  doesn't add another pass.
* Two charts that share a range share the underlying cached values, but the
  visibility mode is applied per chart when the cache XML is written, so a
  `show_hidden_data()` chart and a visible-only chart never leak into each
  other regardless of creation order.
