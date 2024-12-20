# Exceptions used in Google module
#
# Author: Indrajit Ghosh
#
# Date: Nov 16, 2022
#

class GoogleSheetError(Exception):
    """A base class for google sheet exceptions."""

class SpreadsheetNotFound(GoogleSheetError):
    """Trying to open non-existent or inaccessible spreadsheet."""

class WorksheetNotFound(GoogleSheetError):
    """Trying to open non-existent or inaccessible worksheet."""


class CellNotFound(GoogleSheetError):
    """Cell lookup exception."""


def main():
    print('\nExceptions for Google module!')


if __name__ == '__main__':
    main()