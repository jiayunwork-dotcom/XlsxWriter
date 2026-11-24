###############################################################################
#
# An example of using Worksheet.materialize_autofilter() to let XlsxWriter
# apply an autofilter to a large data set that is written in constant memory
# mode.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c) 2013-2025, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter

workbook = xlsxwriter.Workbook(
    "autofilter_materialize.xlsx",
    {"constant_memory": True, "shared_strings": "disk"},
)
worksheet = workbook.add_worksheet()

# Set up the autofilter and write the header row.
worksheet.autofilter(0, 0, 6, 2)
worksheet.write_row(0, 0, ["Region", "Product", "Sales"])

# Write the data rows. In constant_memory mode there is no need to keep the
# data around: XlsxWriter can filter it while the workbook is closed.
data = [
    ["East", "Apple", 1000],
    ["West", "Pear", 2000],
    ["East", "Plum", 3000],
    ["North", "Apple", 4000],
    ["East", "Pear", 5000],
    ["West", "Plum", 6000],
]
for row, row_data in enumerate(data, start=1):
    worksheet.write_row(row, 0, row_data)

# Set the filter conditions after the data has been written and ask for them
# to be applied. Only the East rows with sales above 1000 remain visible.
worksheet.filter_column("A", "x == East")
worksheet.filter_column("C", "x > 1000")
worksheet.materialize_autofilter()

workbook.close()
