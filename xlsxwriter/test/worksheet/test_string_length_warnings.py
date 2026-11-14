###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2026, John McNamara, jmcnamara@cpan.org
#

"""
Tests for the warning emitted when a string exceeds Excel's 32,767
character limit.

The same "string too long" over-limit condition can reach the worksheet
through three public write paths:

1. ``write_string()``  - the string is truncated and -2 is returned.
2. ``write()``         - generic entry, routed to the string path.
3. ``write_rich_string()`` - the input is ignored and -2 is returned.

All three paths must raise a warning, and the warning message must name
the public method that was actually called instead of a hard-coded name.

"""

import unittest
import warnings
from io import StringIO

from xlsxwriter.format import Format
from xlsxwriter.sharedstrings import SharedStringTable
from xlsxwriter.worksheet import Worksheet


class TestStringLengthWarnings(unittest.TestCase):
    """
    Test the 32,767 character over-limit warnings across the write paths.

    """

    def setUp(self):
        self.worksheet = Worksheet()
        self.worksheet.str_table = SharedStringTable()

        # Each emoji is 2 UTF-16 units: 16383 emojis plus an "a" are exactly
        # at the 32,767 limit, one more emoji exceeds it.
        self.limit_string = "\U0001f600" * 16383 + "a"
        self.over_limit_string = "\U0001f600" * 16384

        # A plain ASCII string over the UTF-16 limit.
        self.over_limit_ascii = "a" * 32768

    def _capture_warnings(self, function):
        with warnings.catch_warnings(record=True) as caught:
            warnings.simplefilter("always")
            return_value = function()

        return return_value, caught

    def test_write_string_warns_when_truncated(self):
        """write_string() warns and still returns -2 after truncating."""
        got, caught = self._capture_warnings(
            lambda: self.worksheet.write_string(0, 0, self.over_limit_string)
        )

        self.assertEqual(got, -2)
        self.assertEqual(len(caught), 1)
        self.assertTrue(issubclass(caught[0].category, UserWarning))
        self.assertIn("write_string()", str(caught[0].message))
        self.assertIn("32,767", str(caught[0].message))

        # The truncation behavior itself is unchanged: the cell stores the
        # truncated string, not the original.
        self.assertIn("\U0001f600" * 16383, self.worksheet.str_table.string_table)
        self.assertNotIn(
            self.over_limit_string, self.worksheet.str_table.string_table
        )

    def test_write_string_ascii_warns_when_truncated(self):
        """write_string() warns for an over-limit ASCII string too."""
        got, caught = self._capture_warnings(
            lambda: self.worksheet.write_string(1, 0, self.over_limit_ascii)
        )

        self.assertEqual(got, -2)
        self.assertEqual(len(caught), 1)
        self.assertIn("write_string()", str(caught[0].message))
        self.assertIn("a" * 32767, self.worksheet.str_table.string_table)

    def test_write_warns_with_its_own_entry_name(self):
        """write() warns with a message naming write(), not write_string()."""
        got, caught = self._capture_warnings(
            lambda: self.worksheet.write(2, 0, self.over_limit_string)
        )

        self.assertEqual(got, -2)
        self.assertEqual(len(caught), 1)
        self.assertTrue(issubclass(caught[0].category, UserWarning))
        self.assertIn("write()", str(caught[0].message))
        self.assertNotIn("write_string()", str(caught[0].message))

        # Truncation behavior is unchanged on the generic write() path.
        self.assertIn("\U0001f600" * 16383, self.worksheet.str_table.string_table)

    def test_write_rich_string_warns_with_its_own_entry_name(self):
        """write_rich_string() warns and returns -2 without storing a cell."""
        bold = Format({"bold": 1})

        got, caught = self._capture_warnings(
            lambda: self.worksheet.write_rich_string(
                3, 0, self.over_limit_string, bold, "tail"
            )
        )

        self.assertEqual(got, -2)
        self.assertEqual(len(caught), 1)
        self.assertTrue(issubclass(caught[0].category, UserWarning))
        self.assertIn("write_rich_string()", str(caught[0].message))

        # The rich string input is ignored; no cell is stored.
        self.assertEqual(self.worksheet.table, {})
        self.assertEqual(self.worksheet.str_table.string_table, {})

    def test_rich_string_warning_message_is_not_hard_coded(self):
        """The same over-limit text reports each path's real entry name."""
        bold = Format({"bold": 1})

        _, rich_caught = self._capture_warnings(
            lambda: self.worksheet.write_rich_string(
                4, 0, self.over_limit_string, bold, "tail"
            )
        )
        _, plain_caught = self._capture_warnings(
            lambda: self.worksheet.write_string(5, 0, self.over_limit_string)
        )

        self.assertIn("write_rich_string()", str(rich_caught[0].message))
        self.assertIn("write_string()", str(plain_caught[0].message))
        self.assertNotEqual(
            str(rich_caught[0].message), str(plain_caught[0].message)
        )

    def test_no_warning_at_or_below_the_limit(self):
        """Strings within the limit keep returning 0 without any warning."""
        got, caught = self._capture_warnings(
            lambda: self.worksheet.write_string(6, 0, self.limit_string)
        )

        self.assertEqual(got, 0)
        self.assertEqual(caught, [])

        got, caught = self._capture_warnings(
            lambda: self.worksheet.write(7, 0, "a" * 32767)
        )

        self.assertEqual(got, 0)
        self.assertEqual(caught, [])

    def test_rich_string_within_limit_does_not_warn(self):
        """A rich string at exactly the limit succeeds without a warning."""
        bold = Format({"bold": 1})

        got, caught = self._capture_warnings(
            lambda: self.worksheet.write_rich_string(
                8, 0, self.limit_string[:-1], bold, "a"
            )
        )

        self.assertEqual(got, 0)
        self.assertEqual(caught, [])

    def test_constant_memory_mode_warns_on_both_plain_paths(self):
        """The in-line string (constant_memory) batch path also warns."""
        worksheet = Worksheet()
        worksheet.constant_memory = 1
        worksheet._set_filehandle(StringIO())

        with warnings.catch_warnings(record=True) as caught:
            warnings.simplefilter("always")
            got = worksheet.write_string(0, 0, self.over_limit_ascii)

        self.assertEqual(got, -2)
        self.assertEqual(len(caught), 1)
        self.assertIn("write_string()", str(caught[0].message))

        with warnings.catch_warnings(record=True) as caught:
            warnings.simplefilter("always")
            got = worksheet.write(1, 0, "b" * 32768)

        self.assertEqual(got, -2)
        self.assertEqual(len(caught), 1)
        self.assertIn("write()", str(caught[0].message))

    def test_other_rich_string_warnings_are_unchanged(self):
        """The -3/-4/-5 rich string warnings keep their existing behavior."""
        bold = Format({"bold": 1})
        italic = Format({"italic": 1})

        # -5: insufficient fragments.
        got, caught = self._capture_warnings(
            lambda: self.worksheet.write_rich_string(9, 0, "only one")
        )
        self.assertEqual(got, -5)
        self.assertEqual(len(caught), 1)
        self.assertIn("more than 2 format/fragments", str(caught[0].message))

        # -3: two consecutive formats.
        got, caught = self._capture_warnings(
            lambda: self.worksheet.write_rich_string(
                10, 0, bold, italic, "text"
            )
        )
        self.assertEqual(got, -3)
        self.assertEqual(len(caught), 1)
        self.assertIn("2 consecutive formats", str(caught[0].message))

        # -4: empty string fragment.
        got, caught = self._capture_warnings(
            lambda: self.worksheet.write_rich_string(
                11, 0, "text", bold, ""
            )
        )
        self.assertEqual(got, -4)
        self.assertEqual(len(caught), 1)
        self.assertIn("empty strings", str(caught[0].message))

    def test_write_comment_over_limit_returns_minus_two_silently(self):
        """write_comment() keeps its existing silent -2 behavior."""
        got, caught = self._capture_warnings(
            lambda: self.worksheet.write_comment(12, 0, self.over_limit_ascii)
        )

        self.assertEqual(got, -2)
        self.assertEqual(caught, [])

    def test_write_url_over_limit_warns_as_before(self):
        """The URL length warning behavior is unchanged."""
        long_url = "http://example.com/" + "a" * 2100

        got, caught = self._capture_warnings(
            lambda: self.worksheet.write_url(13, 0, long_url)
        )

        self.assertEqual(got, -3)
        self.assertEqual(len(caught), 1)
        self.assertTrue(issubclass(caught[0].category, UserWarning))
        self.assertIn("limit for URLs", str(caught[0].message))
