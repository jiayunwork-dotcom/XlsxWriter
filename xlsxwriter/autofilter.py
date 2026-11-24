###############################################################################
#
# Autofilter - Compile and evaluate materialized worksheet autofilters.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c) 2013-2025, John McNamara, jmcnamara@cpan.org
#
"""
Autofilter criteria used by Worksheet.materialize_autofilter().

The regular XlsxWriter autofilter only stores the filter metadata in the XLSX
file. This module implements the subset of Excel's filter semantics that can
be applied to the data rows while the workbook is being closed:

* list filters are an OR of the members,
* a custom filter has one or two comparisons joined by AND or OR,
* columns are combined with AND,
* numeric cells support the six comparisons against a finite numeric
  threshold,
* text cells support equality/inequality and list membership, case
  insensitively. Custom equality supports the ``*``, ``?`` and ``~`` wildcard
  syntax; list members are compared literally.

Anything outside that subset raises ``ValueError`` while the plan is compiled
at close time, so a partially filtered file is never produced.
"""

import math
import re
from numbers import Number

from .sharedstrings import (
    CELL_BLANK,
    CELL_BOOLEAN,
    CELL_ERROR,
    CELL_FORMULA,
    CELL_NUMBER,
    CELL_RICH_STRING,
    CELL_STRING,
)
from .utility import xl_col_to_name

# Cell kinds that materialize_autofilter() refuses to filter over. The values
# give the description used in the error message.
UNSUPPORTED_CELL_KINDS = {
    CELL_FORMULA: "formula",
    CELL_BOOLEAN: "boolean",
    CELL_ERROR: "error",
}

# Parser operator ids inherited from Spreadsheet::WriteExcel:
# 1 <, 2 ==, 3 <=, 4 >, 5 !=, 6 >=, 22 wildcard ==.
_ORDERING_OPS = {1, 3, 4, 6}
_EQUALITY_OPS = {2, 5, 22}

# The comparison value used for a missing, blank or empty-string cell.
BLANK_VALUE = ("blank", None)


class FilteredSeriesData(list):
    """Chart cache data with a set of hidden source positions.

    The list still contains every source value so the chart can infer the
    numeric/string type of the whole range. The cache writers omit the points
    whose position is in ``hidden_indices`` but keep iterating over the
    original positions, so ``ptCount`` stays at the range length and the
    remaining ``pt`` elements keep their original ``idx`` values.
    """

    __slots__ = ("hidden_indices",)

    def __init__(self, values, hidden_indices):
        super().__init__(values)
        self.hidden_indices = hidden_indices


class _ListCondition:
    """A filter_column_list() condition: OR of literal members."""

    __slots__ = ("numbers", "texts", "blanks")

    def __init__(self, numbers, texts, blanks):
        self.numbers = numbers
        self.texts = texts
        self.blanks = blanks

    def matches(self, value):
        kind, token = value
        if kind == "num":
            return token in self.numbers
        if kind == "str":
            return token.casefold() in self.texts
        return self.blanks


class _NumberComparison:
    __slots__ = ("op", "threshold")

    def __init__(self, op, threshold):
        self.op = op
        self.threshold = threshold

    def matches(self, value):
        if value[0] != "num":
            return False
        number = value[1]
        op = self.op
        if op == 2:
            return number == self.threshold
        if op == 5:
            return number != self.threshold
        if op == 1:
            return number < self.threshold
        if op == 3:
            return number <= self.threshold
        if op == 4:
            return number > self.threshold
        return number >= self.threshold  # op == 6


class _TextComparison:
    __slots__ = ("regex", "negate")

    def __init__(self, regex, negate):
        self.regex = regex
        self.negate = negate

    def matches(self, value):
        if value[0] != "str":
            return False
        result = self.regex.match(value[1].casefold()) is not None
        return result != self.negate


class _BlankComparison:
    __slots__ = ("equal",)

    def __init__(self, equal):
        self.equal = equal

    def matches(self, value):
        return (value[0] == "blank") == self.equal


class _CustomCondition:
    __slots__ = ("comparisons", "join_and")

    def __init__(self, comparisons, join_and):
        self.comparisons = comparisons
        self.join_and = join_and

    def matches(self, value):
        if self.join_and:
            return all(comp.matches(value) for comp in self.comparisons)
        return any(comp.matches(value) for comp in self.comparisons)


def _finite_number(token):
    # Return a finite float for a token, or None if it isn't a finite number.
    if isinstance(token, bool):
        return None
    if isinstance(token, (int, float)):
        number = float(token)
    else:
        try:
            number = float(token)
        except (TypeError, ValueError):
            return None
    if not math.isfinite(number):
        return None
    return number


def _wildcard_regex(pattern):
    # Convert Excel's filter wildcard syntax into a full-match regular
    # expression. The caller casefolds the pattern and the candidate.
    parts = ["\\A"]
    pos = 0
    length = len(pattern)
    while pos < length:
        char = pattern[pos]
        if char == "~" and pos + 1 < length:
            parts.append(re.escape(pattern[pos + 1]))
            pos += 2
            continue
        if char == "*":
            parts.append(".*")
        elif char == "?":
            parts.append(".")
        else:
            parts.append(re.escape(char))
        pos += 1
    parts.append("\\Z")
    return re.compile("".join(parts), re.DOTALL)


def _blank_comparison(op, token):
    # The parser represents "== Blanks" as [2, "blanks"] and "!= Blanks",
    # "== NonBlanks" as [5, " "]. "x <> NonBlanks" becomes [2, "blanks"].
    if op == 2 and str(token).casefold() == "blanks":
        return _BlankComparison(True)
    if op == 5 and token == " ":
        return _BlankComparison(False)
    return None


def _compile_comparison(op, token, column_name):
    if op not in _ORDERING_OPS and op not in _EQUALITY_OPS:
        raise ValueError(
            f"Unsupported autofilter operator '{op}' in column {column_name}. "
            "materialize_autofilter() supports ==, !=, <, <=, > and >=."
        )

    blank = _blank_comparison(op, token)
    if blank is not None:
        return blank

    if op in _ORDERING_OPS:
        number = _finite_number(token)
        if number is None:
            raise ValueError(
                f"Autofilter comparison in column {column_name} needs a finite "
                f"numeric threshold, got '{token}'."
            )
        return _NumberComparison(op, number)

    number = _finite_number(token)
    if number is not None:
        return _NumberComparison(op, number)

    return _TextComparison(_wildcard_regex(str(token).casefold()), op == 5)


def _compile_list(criteria, from_equality, col_name):
    numbers = set()
    texts = set()
    blanks = False

    for member in criteria:
        if isinstance(member, bool):
            raise ValueError(
                f"Unsupported boolean list member {member!r} in autofilter "
                f"column {col_name}."
            )
        if isinstance(member, str):
            if member.casefold() == "blanks" or member == "":
                # "Blanks" is the documented list member for blank cells. An
                # empty string is also treated as blank because empty text
                # cells are blanks, so a literal "" member would never match
                # a stored string.
                blanks = True
                continue
            if from_equality:
                number = _finite_number(member)
                if number is not None:
                    numbers.add(number)
                    continue
            texts.add(member.casefold())
            continue
        if isinstance(member, Number):
            number = _finite_number(member)
            if number is None:
                raise ValueError(
                    f"Autofilter list member {member!r} in column {col_name} "
                    "must be a finite number or a string."
                )
            numbers.add(number)
            continue
        raise ValueError(
            f"Unsupported autofilter list member {member!r} in column "
            f"{col_name}. Only strings and finite numbers are supported."
        )

    return _ListCondition(numbers, texts, blanks)


def _compile_custom(criteria, col_name):
    if len(criteria) == 2:
        op, token = criteria
        comparisons = [_compile_comparison(op, token, col_name)]
        join_and = True

    elif len(criteria) == 5:
        op_1, token_1, join, op_2, token_2 = criteria
        if join not in (0, 1):
            raise ValueError(
                f"Unsupported autofilter join {join!r} in column {col_name}."
            )
        comparisons = [
            _compile_comparison(op_1, token_1, col_name),
            _compile_comparison(op_2, token_2, col_name),
        ]
        join_and = join == 0

    else:
        raise ValueError(
            f"Unsupported autofilter custom condition in column {col_name}: "
            f"{criteria!r}."
        )

    return _CustomCondition(comparisons, join_and)


def compile_filter_plan(
    filter_cols, filter_type, equality_cols, col_first, col_last
):
    """Compile the final worksheet filter metadata into conditions.

    Returns a dict of absolute column number -> condition. Columns outside the
    final autofilter range are ignored, matching the metadata writer.
    """
    conditions = {}

    for col in sorted(filter_cols):
        if col < col_first or col > col_last:
            continue

        col_name = xl_col_to_name(col) or str(col)
        criteria = filter_cols[col]
        filter_kind = filter_type.get(col)

        if filter_kind == 1:
            from_equality = col in equality_cols
            conditions[col] = _compile_list(criteria, from_equality, col_name)
        elif filter_kind == 0:
            conditions[col] = _compile_custom(criteria, col_name)
        else:
            raise ValueError(
                f"Unsupported autofilter condition type {filter_kind!r} in "
                f"column {col_name}."
            )

    return conditions


def check_cell_kind(kind, sheet_name, row, col_name):
    """Raise ValueError for a cell kind that can't be evaluated."""
    description = UNSUPPORTED_CELL_KINDS.get(kind)
    if description is not None:
        raise ValueError(
            f"Autofilter column {col_name} in worksheet '{sheet_name}' "
            f"contains a {description} cell at row {row + 1}. "
            "materialize_autofilter() cannot evaluate this filter."
        )


def row_matches(plan, values):
    """Return True if a row's filter values satisfy every column condition."""
    for col, condition in plan.items():
        if not condition.matches(values.get(col, BLANK_VALUE)):
            return False
    return True


def filter_value_from_record(kind, value, display):
    """Convert a flushed cell record into a filter comparison value."""
    if kind == CELL_NUMBER:
        return ("num", float(value))
    if kind in (CELL_STRING, CELL_RICH_STRING):
        text = display if display is not None else ""
        return ("str", text) if text else ("blank", None)
    if kind == CELL_BLANK:
        return BLANK_VALUE
    # The caller validates kinds before calling this.
    return BLANK_VALUE
