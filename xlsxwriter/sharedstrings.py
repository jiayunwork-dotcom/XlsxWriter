###############################################################################
#
# SharedStrings - A class for writing the Excel XLSX sharedStrings file.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c) 2013-2025, John McNamara, jmcnamara@cpan.org
#

# Standard packages.
import os
import sqlite3
import tempfile
from collections import OrderedDict

# Package imports.
from . import xmlwriter
from .utility import _preserve_whitespace


class SharedStrings(xmlwriter.XMLwriter):
    """
    A class for writing the Excel XLSX sharedStrings file.

    """

    ###########################################################################
    #
    # Public API.
    #
    ###########################################################################

    def __init__(self) -> None:
        """
        Constructor.

        """

        super().__init__()

        self.string_table = None

    ###########################################################################
    #
    # Private API.
    #
    ###########################################################################

    def _assemble_xml_file(self) -> None:
        # Assemble and write the XML file.

        # Write the XML declaration.
        self._xml_declaration()

        # Write the sst element.
        self._write_sst()

        # Write the sst strings.
        self._write_sst_strings()

        # Close the sst tag.
        self._xml_end_tag("sst")

        # Close the file.
        self._xml_close()

    ###########################################################################
    #
    # XML methods.
    #
    ###########################################################################

    def _write_sst(self) -> None:
        # Write the <sst> element.
        xmlns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

        attributes = [
            ("xmlns", xmlns),
            ("count", self.string_table.count),
            ("uniqueCount", self.string_table.unique_count),
        ]

        self._xml_start_tag("sst", attributes)

    def _write_sst_strings(self) -> None:
        # Write the sst string elements.

        if getattr(self.string_table, "is_disk", False):
            # Stream the strings from the on-disk store. Only one <si> element
            # is held in memory at a time.
            for string, is_rich in self.string_table._iter_strings():
                self._write_si(string, is_rich)
            return

        for string in self.string_table.string_array:
            self._write_si(string)

    def _write_si(self, string, is_rich=None) -> None:
        # Write the <si> element.
        attributes = []

        # Convert control character to a _xHHHH_ escape.
        string = self._escape_control_characters(string)

        # Add attribute to preserve leading or trailing whitespace.
        if _preserve_whitespace(string):
            attributes.append(("xml:space", "preserve"))

        # When the caller doesn't tell us the type we fall back to the
        # historical heuristic of detecting rich strings from their markup.
        # The disk backed store knows the string type explicitly and passes
        # it in so that plain text that happens to look like <r>..</r> is
        # still written as a plain string.
        if is_rich is None:
            is_rich = string.startswith("<r>") and string.endswith("</r>")

        # Write any rich strings without further tags.
        if is_rich:
            self._xml_rich_si_element(string)
        else:
            self._xml_si_element(string, attributes)


# A metadata class to store Excel strings between worksheets.
class SharedStringTable:
    """
    A class to track Excel shared strings between worksheets.

    """

    def __init__(self) -> None:
        self.count = 0
        self.unique_count = 0
        self.string_table = {}
        self.string_array = []

    def _get_shared_string_index(self, string):
        """ " Get the index of the string in the Shared String table."""
        if string not in self.string_table:
            # String isn't already stored in the table so add it.
            index = self.unique_count
            self.string_table[string] = index
            self.count += 1
            self.unique_count += 1
            return index

        # String exists in the table.
        index = self.string_table[string]
        self.count += 1
        return index

    def _get_shared_string(self, index):
        """ " Get a shared string from the index."""
        return self.string_array[index]

    def _sort_string_data(self) -> None:
        """ " Sort the shared string data and convert from dict to list."""
        self.string_array = sorted(self.string_table, key=self.string_table.__getitem__)
        self.string_table = {}

    def _has_shared_strings(self) -> bool:
        # Check if the workbook needs a sharedStrings.xml part.
        return self.count > 0


# The cell record kinds used by the disk backed store.
CELL_NUMBER = 0
CELL_STRING = 1
CELL_RICH_STRING = 2
CELL_FORMULA = 3
CELL_BLANK = 4
CELL_BOOLEAN = 5
CELL_ERROR = 6

# The string record kinds used by the disk backed store.
STRING_PLAIN = 0
STRING_RICH = 1


class DiskSharedStringTable:
    """
    A disk backed SharedStringTable.

    The strings are de-duplicated in an SQLite database in the workbook's
    temp directory. The table records two things:

    1. The shared string catalog, keyed by ``(kind, payload)`` and numbered
       in insertion order.
    2. The cells that have been "flushed" out of the worksheet row buffers.
       These are needed later for chart range lookups.

    Only a bounded number of strings are ever kept in memory (a small LRU
    cache) and the shared strings XML is streamed back from disk when the
    workbook is closed.

    """

    is_disk = True

    # Size of the bounded in-memory de-duplication cache. This is the only
    # string data that stays resident in this class.
    _CACHE_SIZE = 512

    # SQLite page cache size in KiB.
    _CACHE_KIB = 2048

    def __init__(self, tmpdir=None) -> None:
        (fd, filename) = tempfile.mkstemp(
            dir=tmpdir, prefix="xlsxwriter_sst_", suffix=".sqlite3"
        )
        os.close(fd)
        self.filename = filename
        self.closed = False
        self.failed = False
        self._any_strings = False
        self._next_index = 0
        self._lru = OrderedDict()
        self._final_count = 0
        self._final_unique_count = 0

        self.db = sqlite3.connect(filename, isolation_level=None)
        # Keep the memory footprint small and avoid unneeded disk syncs.
        self.db.execute("PRAGMA journal_mode=OFF")
        self.db.execute("PRAGMA synchronous=OFF")
        self.db.execute("PRAGMA locking_mode=EXCLUSIVE")
        self.db.execute("PRAGMA temp_store=FILE")
        self.db.execute(f"PRAGMA cache_size=-{self._CACHE_KIB}")
        self.db.execute(
            "CREATE TABLE strings ("
            "idx INTEGER PRIMARY KEY, "
            "kind INTEGER NOT NULL, "
            "payload TEXT NOT NULL, "
            "display TEXT)"
        )
        self.db.execute(
            "CREATE UNIQUE INDEX strings_dedup ON strings (kind, payload)"
        )
        self.db.execute(
            "CREATE TABLE cells ("
            "sheet INTEGER NOT NULL, "
            "row INTEGER NOT NULL, "
            "col INTEGER NOT NULL, "
            "kind INTEGER NOT NULL, "
            "value TEXT, "
            "sidx INTEGER, "
            "PRIMARY KEY (sheet, row, col)) WITHOUT ROWID"
        )
        self.db.execute(
            "CREATE TABLE filter_hidden ("
            "sheet INTEGER NOT NULL, "
            "row INTEGER NOT NULL, "
            "PRIMARY KEY (sheet, row)) WITHOUT ROWID"
        )
        # Error cells aren't part of the chart cache cell records, but a
        # materialized autofilter must refuse to filter over them. They are
        # rare, so recording them separately keeps the chart path unchanged.
        self.db.execute(
            "CREATE TABLE filter_errors ("
            "sheet INTEGER NOT NULL, "
            "row INTEGER NOT NULL, "
            "col INTEGER NOT NULL, "
            "PRIMARY KEY (sheet, row, col)) WITHOUT ROWID"
        )

    ###########################################################################
    #
    # Shared string catalog.
    #
    ###########################################################################

    def _get_shared_string_index(self, string, kind=STRING_PLAIN, display=None):
        # Return the shared string index for a string, adding it if it is
        # new. Plain strings are de-duplicated by text and rich strings by
        # their canonical (ordered fragments + formats) XML. The index is
        # assigned in insertion order and is independent of the temp file
        # location.
        if display is None:
            display = string

        # For plain strings the display text is the payload, so don't store
        # it twice. Rich strings always store the raw unformatted text as the
        # display value used by chart caches.
        stored_display = None if kind == STRING_PLAIN else display

        key = (kind, string)
        index = self._lru.get(key)

        if index is not None:
            # Move to the end of the LRU.
            self._lru.move_to_end(key)
            return index

        index = self._next_index
        try:
            cursor = self.db.execute(
                "INSERT OR IGNORE INTO strings (idx, kind, payload, display) "
                "VALUES (?, ?, ?, ?)",
                (index, kind, string, stored_display),
            )
            if cursor.rowcount == 0:
                # A duplicate that wasn't in the LRU cache.
                row = self.db.execute(
                    "SELECT idx FROM strings WHERE kind = ? AND payload = ?",
                    (kind, string),
                ).fetchone()
                index = row[0]
            else:
                self._next_index += 1
        except sqlite3.Error:
            # Remember that staging failed so that close() doesn't produce a
            # successful (but incomplete) file.
            self.failed = True
            raise

        self._any_strings = True
        self._lru[key] = index
        if len(self._lru) > self._CACHE_SIZE:
            self._lru.popitem(last=False)

        return index

    def _get_shared_string(self, index):
        # Return the display text of a shared string from its index.
        row = self.db.execute(
            "SELECT COALESCE(display, payload) FROM strings WHERE idx = ?", (index,)
        ).fetchone()
        if row is None:
            return ""
        return row[0]

    @property
    def unique_count(self):
        # The number of <si> elements in the table.
        if self.db is None:
            return self._final_unique_count
        return self.db.execute("SELECT COUNT(*) FROM strings").fetchone()[0]

    @property
    def count(self):
        # The number of retained string cells. Cells that were overwritten
        # before their row was flushed never get a cell record, so this is
        # the number of cells that actually reference a string index.
        if self.db is None:
            return self._final_count
        return self.db.execute(
            "SELECT COUNT(*) FROM cells WHERE kind IN (?, ?)",
            (CELL_STRING, CELL_RICH_STRING),
        ).fetchone()[0]

    def _iter_strings(self):
        # Yield (payload, is_rich) for each string in index order. The
        # cursor is streamed so that only a small number of rows are held
        # in memory at a time. Unreferenced entries are written as well.
        cursor = self.db.execute("SELECT kind, payload FROM strings ORDER BY idx")
        cursor.arraysize = 64
        for kind, payload in cursor:
            yield payload, kind == STRING_RICH

    def _sort_string_data(self) -> None:
        # The disk table is already in index order. No-op.
        pass

    def _has_shared_strings(self) -> bool:
        # Only the very first insert needs to touch the DB for this. After
        # that the in-memory flag is authoritative.
        return self._any_strings

    ###########################################################################
    #
    # Flushed cell records used for chart range lookups.
    #
    ###########################################################################

    def _record_cells(self, records) -> None:
        # Store flushed cell records in a single transaction. Each record is
        # a (sheet, row, col, kind, value, sidx) tuple. The record is keyed
        # by its worksheet position so a range can be looked up without
        # reading the whole table.
        self.db.executemany(
            "INSERT OR REPLACE INTO cells (sheet, row, col, kind, value, sidx) "
            "VALUES (?, ?, ?, ?, ?, ?)",
            records,
        )

    def _get_range_cells(self, sheet, row_start, col_start, row_end, col_end):
        # Return an iterator of (row, col, kind, value) records for a range.
        # String kinds are joined to their display text so that the caller
        # doesn't have to look each one up separately. The identity of a
        # string is its SST index (sidx), so the display text always comes
        # from the same table that produces sharedStrings.xml.
        return self.db.execute(
            "SELECT c.row, c.col, c.kind, c.value, c.sidx, "
            "       COALESCE(s.display, s.payload) "
            "FROM cells c LEFT JOIN strings s ON s.idx = c.sidx "
            "WHERE c.sheet = ? AND c.row BETWEEN ? AND ? "
            "  AND c.col BETWEEN ? AND ?",
            (sheet, row_start, row_end, col_start, col_end),
        )

    def _get_range_rows(self, sheet, row_start, row_end):
        # Return the set of rows that have at least one flushed cell record
        # in the given row range. This mirrors the in-memory lookup which
        # emits a single None for a row that doesn't exist at all.
        cursor = self.db.execute(
            "SELECT DISTINCT row FROM cells "
            "WHERE sheet = ? AND row BETWEEN ? AND ?",
            (sheet, row_start, row_end),
        )
        return {row[0] for row in cursor}

    def _record_error_cells(self, records):
        # Store the positions of flushed error cells for filter validation.
        self.db.executemany(
            "INSERT OR REPLACE INTO filter_errors (sheet, row, col) "
            "VALUES (?, ?, ?)",
            records,
        )

    def _first_filter_error(self, sheet, row_start, row_end, columns):
        # Return the first error cell in the given columns/rows, if any.
        if not columns:
            return None
        placeholders = ",".join("?" for _ in columns)
        params = [sheet, row_start, row_end]
        params.extend(sorted(columns))
        return self.db.execute(
            "SELECT row, col FROM filter_errors "
            f"WHERE sheet = ? AND row BETWEEN ? AND ? "
            f"  AND col IN ({placeholders}) ORDER BY row, col LIMIT 1",
            params,
        ).fetchone()

    def _get_filter_column_cells(self, sheet, row_start, row_end, columns):
        # Return an ordered cursor over the flushed cells in the given
        # columns. Only the cells that can affect an autofilter decision are
        # read, and the caller streams the rows in ascending order. String
        # display text is resolved through the shared string table so rich
        # strings and plain strings compare by what is shown.
        if not columns:
            return iter(())

        placeholders = ",".join("?" for _ in columns)
        params = [sheet, row_start, row_end]
        params.extend(sorted(columns))
        return self.db.execute(
            "SELECT c.row, c.col, c.kind, c.value, c.sidx, "
            "       COALESCE(s.display, s.payload) "
            "FROM cells c LEFT JOIN strings s ON s.idx = c.sidx "
            f"WHERE c.sheet = ? AND c.row BETWEEN ? AND ? "
            f"  AND c.col IN ({placeholders}) "
            "ORDER BY c.row, c.col",
            params,
        )

    def _clear_hidden_rows(self, sheet):
        # Remove any previous materialized decisions for a sheet. This makes
        # a repeated close (for example after a failure) start from scratch.
        self.db.execute("DELETE FROM filter_hidden WHERE sheet = ?", (sheet,))

    def _insert_hidden_rows(self, rows):
        # Insert a batch of (sheet, row) hidden decisions. The caller manages
        # the transaction so an entire sheet's decisions are one disk sync
        # while large decision sets are still written batch by batch.
        self.db.executemany(
            "INSERT OR REPLACE INTO filter_hidden (sheet, row) VALUES (?, ?)",
            rows,
        )

    def _begin_hidden_transaction(self):
        # Use an immediate transaction so the poll of unsupported cells and
        # the decision writes see a consistent database.
        self.db.execute("BEGIN IMMEDIATE")

    def _commit_hidden_transaction(self):
        self.db.execute("COMMIT")

    def _rollback_hidden_transaction(self):
        try:
            self.db.execute("ROLLBACK")
        except sqlite3.Error:
            pass

    def _iter_hidden_rows(self, sheet, row_start, row_end):
        # Yield the hidden row numbers for a sheet in ascending order. Used
        # by the streaming worksheet writer so it can merge them with the row
        # data file without loading either side into memory.
        cursor = self.db.execute(
            "SELECT row FROM filter_hidden "
            "WHERE sheet = ? AND row BETWEEN ? AND ? ORDER BY row",
            (sheet, row_start, row_end),
        )
        cursor.arraysize = 256
        for (row,) in cursor:
            yield row

    def _get_hidden_rows(self, sheet, row_start, row_end):
        # Return the hidden rows in a range as a set. Only chart ranges use
        # this, and a chart's cache is already limited to the ranges it
        # actually references.
        cursor = self.db.execute(
            "SELECT row FROM filter_hidden "
            "WHERE sheet = ? AND row BETWEEN ? AND ?",
            (sheet, row_start, row_end),
        )
        return {row[0] for row in cursor}

    ###########################################################################
    #
    # Cleanup.
    #
    ###########################################################################

    def _close(self) -> None:
        # Close and remove the temporary database. Safe to call more than once.
        if self.closed:
            return

        try:
            if self.db is not None:
                # Remember the final counts in case anything inspects the
                # table after close().
                try:
                    self._final_unique_count = self.db.execute(
                        "SELECT COUNT(*) FROM strings"
                    ).fetchone()[0]
                    self._final_count = self.db.execute(
                        "SELECT COUNT(*) FROM cells WHERE kind IN (?, ?)",
                        (CELL_STRING, CELL_RICH_STRING),
                    ).fetchone()[0]
                except sqlite3.Error:
                    pass
                self.db.close()
        finally:
            self.db = None
            self.closed = True

        for suffix in ("", "-journal", "-wal", "-shm"):
            path = self.filename + suffix
            try:
                os.remove(path)
            except FileNotFoundError:
                pass
            except OSError:
                pass

