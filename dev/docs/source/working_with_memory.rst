.. SPDX-License-Identifier: BSD-2-Clause
   Copyright (c) 2013-2025, John McNamara, jmcnamara@cpan.org

.. _memory_perf:

Working with Memory and Performance
===================================

By default XlsxWriter holds all cell data in memory. This is to allow future
features when formatting is applied separately from the data.

The effect of this is that XlsxWriter can consume a lot of memory and it is
possible to run out of memory when creating large files.

Fortunately, this memory usage can be reduced almost completely by using the
:func:`Workbook` ``'constant_memory'`` property::

    workbook = xlsxwriter.Workbook(filename, {'constant_memory': True})

The optimization works by flushing each row after a subsequent row is written.
In this way the largest amount of data held in memory for a worksheet is the
amount of data required to hold a single row of data.

Since each new row flushes the previous row, data must be written in sequential
row order when ``'constant_memory'`` mode is on::

    # Ok. With 'constant_memory' you must write data in row by column order.
    for row in range(0, row_max):
        for col in range(0, col_max):
            worksheet.write(row, col, some_data)

    # Not ok. With 'constant_memory' this will only write the first column of data.
    for col in range(0, col_max):
        for row in range(0, row_max):
            worksheet.write(row, col, some_data)

Another optimization that is used to reduce memory usage is that cell strings
aren't stored in an Excel structure call "shared strings" and instead are
written "in-line". This is a documented Excel feature that is supported by
most spreadsheet applications.

If a receiving application needs the standard ``sharedStrings.xml`` table you
can keep it while still using ``constant_memory`` by adding the
``'shared_strings': 'disk'`` option::

    workbook = xlsxwriter.Workbook(
        filename,
        {'constant_memory': True, 'shared_strings': 'disk'},
    )

In this mode strings are de-duplicated in a temporary SQLite database in the
``tmpdir`` and the shared string table is streamed back out when the workbook
is closed. Peak memory stays bounded regardless of the number of distinct
strings. The same database records the flushed cell values so that chart data
caches remain correct for rows that have already been written to disk.

The trade-off when using ``'constant_memory'`` mode is that you won't be able
to take advantage of any new features that manipulate cell data after it is
written. Currently the :func:`add_table()` method doesn't work in this mode
and :func:`merge_range()` and :func:`set_row()` only work for the current row.

One feature that does work with the disk backed mode is the materialized
autofilter. Calling :func:`materialize_autofilter()` hides the non-matching
rows at close time, so the caller doesn't need to iterate the data and call
:func:`set_row()` for every row::

    workbook = xlsxwriter.Workbook(
        filename,
        {'constant_memory': True, 'shared_strings': 'disk'},
    )
    worksheet = workbook.add_worksheet()
    worksheet.autofilter('A1:D51')
    worksheet.filter_column('A', 'Region == East')
    worksheet.materialize_autofilter()

    for row in data:
        worksheet.write_row(row, 0, row_data)

The matching is streamed from the same SQLite store that holds the flushed
cells, so peak memory stays flat. The decisions are written to that store and
reused by the worksheet XML writer and the chart cache builder, so charts
that reference the range see exactly the rows that the file hides. See
:ref:`working_with_autofilters` for the supported conditions and the
``dev/autofilter`` experiment for the resource measurements.


Performance Figures
-------------------

The performance figures below show execution time and memory usage for
worksheets of size ``N`` rows x 50 columns with a 50/50 mixture of strings and
numbers. The figures are taken from an arbitrary, mid-range, machine. Specific
figures will vary from machine to machine but the trends should be the same.

XlsxWriter in normal operation mode: the execution time and memory usage
increase more or less linearly with the number of rows:

+-------+---------+----------+----------------+
| Rows  | Columns | Time (s) | Memory (bytes) |
+=======+=========+==========+================+
| 200   | 50      | 0.43     | 2346728        |
+-------+---------+----------+----------------+
| 400   | 50      | 0.84     | 4670904        |
+-------+---------+----------+----------------+
| 800   | 50      | 1.68     | 8325928        |
+-------+---------+----------+----------------+
| 1600  | 50      | 3.39     | 17855192       |
+-------+---------+----------+----------------+
| 3200  | 50      | 6.82     | 32279672       |
+-------+---------+----------+----------------+
| 6400  | 50      | 13.66    | 64862232       |
+-------+---------+----------+----------------+
| 12800 | 50      | 27.60    | 128851880      |
+-------+---------+----------+----------------+

XlsxWriter in ``constant_memory`` mode: the execution time still increases
linearly with the number of rows but the memory usage remains small and
constant:

+-------+---------+----------+----------------+
| Rows  | Columns | Time (s) | Memory (bytes) |
+=======+=========+==========+================+
| 200   | 50      | 0.37     | 62208          |
+-------+---------+----------+----------------+
| 400   | 50      | 0.74     | 62208          |
+-------+---------+----------+----------------+
| 800   | 50      | 1.46     | 62208          |
+-------+---------+----------+----------------+
| 1600  | 50      | 2.93     | 62208          |
+-------+---------+----------+----------------+
| 3200  | 50      | 5.90     | 62208          |
+-------+---------+----------+----------------+
| 6400  | 50      | 11.84    | 62208          |
+-------+---------+----------+----------------+
| 12800 | 50      | 23.63    | 62208          |
+-------+---------+----------+----------------+

In ``constant_memory`` mode the performance should be approximately the same
as normal mode.

These figures were generated using programs in the ``dev/performance``
directory of the XlsxWriter repo.
