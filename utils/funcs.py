from typing import Any, Optional, Literal
from collections.abc import Callable, Sequence, Iterable, Iterator

from utils.consts import ROMAN

from datetime import date, datetime
from re import compile as recompile, escape
from functools import cache


def str2date(day: str, /) -> date:
    # date(*tuple(int(x) for x in day.split('.')[::-1]))
    return datetime.strptime(day, '%d.%m.%Y').date()

def date2str(day: date, /) -> str:
    return day.strftime('%d.%m.%Y')

@cache
def roman2int(number: str, /) -> int:
    try:
        return int(number, base=10)
    except ValueError:
        str_num: str = number.upper()
        if not set(ROMAN.keys()).issuperset(str_num):
            return 0
        rest: int = 0
        for char in reversed(str_num):
            num: int = ROMAN.get(char, 0)
            rest: int = rest - num if 3 * num < rest else rest + num
        return rest
    return 0

def point2tuple(point: str, /) -> tuple[int, ...]:
    return tuple(roman2int(n) for n in point.split('.'))

def sort2return[T](iterable: Iterable[T], /, *, key: Callable[[T], Any]=None, reverse: bool=False) -> list[T]:
    if isinstance(iterable, list):
        iterable.sort(key=key, reverse=reverse)
        return iterable
    return sorted(iterable, key=key, reverse=reverse)

def get_state(state: bool, /) -> Literal['disabled', 'normal']:
    return 'normal' if state else 'disabled'

def extract_dates(day: date, dates: Sequence[tuple[date, date]], /) -> Optional[tuple[date, date]]:
    # check which timeframe is correct #
    for beg, end in dates:
        if beg <= day <= end:
            return beg, end
    return None

def connect_dates(beg: date, end: date, /) -> str:
    return ' - '.join(date2str(d) for d in (beg, end))

def get_timespan_desc(beg: date, end: date, /):
    if beg.year != end.year:
        return f'{beg.month}_{beg:%y}-{end.month}_{end:%y}'
    if beg.month != end.month:
        return f'{beg.month}-{end.month}_{beg:%y}'
    if beg.day != end.day:
        return f'{beg.day}-{end.day}_{beg.month}_{beg:%y}'
    return f'{beg.day}_{beg.month}_{beg:%y}'

def pair_cross[T](seq: Sequence[T], /) -> Iterator[tuple[T, T], ...]:
    if not (isinstance(seq, list) or isinstance(seq, tuple)):
        seq = tuple(seq)
    return zip(seq, seq[1:])

def pair_up[T](seq: Sequence[T], /) -> Iterator[tuple[T, T], ...]:
    if not (isinstance(seq, list) or isinstance(seq, tuple)):
        seq = tuple(seq)
    return zip(seq[::2], seq[1::2])

def flatten[T](seq: Sequence[tuple[T, ...]], /) -> tuple[T, ...]:
    return sum(seq, ())

def replace_mutiple(text: str, trans_table: dict[str, str], /) -> str:
    transpose: dict[str, str] = {escape(key): value for key, value in trans_table.items()}
    return recompile('|'.join(transpose.keys())).sub(lambda x: transpose[escape(x.group(0))], text)
