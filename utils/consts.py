from typing import Final

# from enum import EnumType
from datetime import date

BAD_CHARS: Final[str] = r'\/:*?"<>|'
MIN_DATE: Final[date] = date(1900, 1, 1)
ROMAN: Final[dict[str, int]] = {'I': 1, 'V': 5, 'X': 10, 'L': 50, 'C': 100, 'D': 500, 'M': 1000}

# class Slice(EnumType):
#     REVERSED: slice = slice(None, None, -1)
#     ALL_BUT_FIRST: slice = slice(1, None)
#     EVEN: slice = slice(None, None, 2)
#     ODD: slice = slice(1, None, 2)
