#######################################################################
#
# An example of using the disk backed shared strings export path with
# 'constant_memory' mode. This keeps a real sharedStrings.xml table while
# still writing the workbook with bounded memory.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c) 2013-2025, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter

workbook = xlsxwriter.Workbook(
    "disk_shared_strings.xlsx",
    {
        "constant_memory": True,
        "shared_strings": "disk",
    },
)
worksheet = workbook.add_worksheet()

bold = workbook.add_format({"bold": True})
red = workbook.add_format({"font_color": "red"})

# Data is written in row order, as required by 'constant_memory' mode.
worksheet.write_string(0, 0, "Repeated category")
worksheet.write_number(0, 1, 10)

worksheet.write_string(1, 0, "Repeated category")
worksheet.write_number(1, 1, 20)

worksheet.write_rich_string(2, 0, "Rich ", bold, "text", red, " fragments")
worksheet.write_number(2, 1, 30)

worksheet.write_url(3, 0, "https://example.com", None, "Link text")
worksheet.write_number(3, 1, 40)

# Rows written above have already been flushed to disk. The chart can still
# use them because their values are recorded alongside the shared strings.
worksheet.write_string(10, 0, "Sentinel")

chart = workbook.add_chart({"type": "column"})
chart.add_series(
    {
        "categories": "=Sheet1!$A$1:$A$4",
        "values": "=Sheet1!$B$1:$B$4",
    }
)
worksheet.insert_chart("D2", chart)

workbook.close()
