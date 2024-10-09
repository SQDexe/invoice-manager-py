from typing import Any, Final
from collections.abc import Callable, Hashable, Sequence, Iterable, ItemsView
from tkinter import Widget

from abc import ABC, abstractmethod

from sys import version_info, executable
from os.path import normpath, dirname
from datetime import date, datetime
from re import compile as recompile, escape
from functools import cache

from tkinter import Tk, PhotoImage, Menu
from tkinter.ttk import Style, Frame
from tkinter.messagebox import showwarning, showerror, showinfo
from tkinter.filedialog import askopenfilename
from tktooltip import ToolTip

type Function = Callable[[...], Any]

MIN_DATE: Final[date] = date(1900, 1, 1)
ROMAN: Final[dict[str, int]] = {'I': 1, 'V': 5, 'X': 10, 'L': 50, 'C': 100, 'D': 500, 'M': 1000}

# namespace object for better variable handling #
class Namespace:
    def __init__(self, other: dict[Hashable, Any]=None, /, **kwargs: Any) -> None:
        self.update(other, **kwargs)

    def __getitem__(self, key: Hashable) -> Any:
        return self.__dict__[key]

    def __setitem__(self, key: Hashable, value: Any) -> None:
        self.__dict__[key] = value

    def items(self) -> ItemsView[Hashable, Any]:
        return self.__dict__.items()

    def update(self, other: dict[Hashable, Any]=None, /, **kwargs: Any) -> None:
        if not isinstance(other, dict):
            other = {}
        self.__dict__.update(other | kwargs)

class WindowApp(ABC):
    def __init__(self) -> None:
        # root declaration #
        self.root: Tk = Tk()

        # declearing variables, and elements #
        self.vars: Namespace = Namespace(
          title = None,
          size = Namespace(
            min = (None, None),
            max = (None, None)
            ),
          ver = Namespace(
            py = '{}.{}.{}'.format(version_info.major, version_info.minor, version_info.micro),
            tk = self.root.tk.call('info', 'patchlevel'),
            pyinst = '6.8.0'
            ),
          patterns = Namespace(
            json = recompile(r'\.json$')
            ),
          errors = {
            # 0XX - file errors
            1: 'Brak podstawowego pliku danych',
            2: 'Niedozwolony znak w nazwie pliku',
            3: 'Brak nazwy pliku',
            4: 'Złe rozszerzenie pliku',
            5: 'Plik o tej nazwie już istnieje',
            # 1XX - program errors
            101: 'Funkcja na niedozwolonym elemencie',
            # 2XX - general errors
            201: 'Nie wybrano elementu',
            # 5XX - dp errors
            501: 'Nie podano nazwy',
            502: 'Nazwa zajęta',
            503: 'Zły okres',
            504: 'Okres zajęty',
            505: 'Złe formatowanie tekstu',
            506: 'Zły format nazwy punktu',
            # 6XX - tp errors
            601: 'Wybrany został projekt',
            602: 'Punkt(y) już został wybrany',
            603: 'Data poza zasięgiem'
            },
          var = Namespace(),
          pad = 5,
          bad_chars = r'\/:*?"<>|',
          style = Style()
          )
        self.elem: Namespace = Namespace()
        self.tooltips: list[ToolTip] = []
        self.grid: Namespace = Namespace(
          row = {'default': {'weight': 1, 'minsize': 40} },
          col = {'default': {'weight': 1, 'minsize': 50} }    
          )
        self.menus: Namespace = Namespace()
        self.binds: list[tuple[str, tuple[str, Function]]] = []

        # pre-generation funtions #
        self.pre()

        # window's settings #
        width, height = self.vars.size.min
        self.root.minsize(width, height)
        self.root.maxsize(*self.vars.size.max)
        self.root.resizable(True, True)
        self.root.geometry('{}x{}+{}+{}'.format(
          width,
          height,
          (self.root.winfo_screenwidth() - width) // 2,
          (self.root.winfo_screenheight() - height) // 2
          ))
        self.root.title(self.vars.title)
        self.root.iconbitmap(executable)

        # protocols #
        self.root.protocol('WM_DELETE_WINDOW', self.close)

        # prepare elements #
        self.prep_elems()

        # post-generation funtions #
        self.post()

        # start program #
        self.root.mainloop()

    def prep_elems(self) -> None:
        main: Final[str] = 'main'

        # mainframe #
        self.elem[main]: Frame = Frame(self.root)
        self.elem[main].pack(fill='both', expand=True)

        # create, and configure elements #
        for key, data in tuple((key, data) for key, data in self.elem.items() if key != main):
            self.elem[key]: Widget = data['type'](self.elem[main], **data.get('args', {}))
            options: dict[str, int] = {'padx': self.vars.pad, 'pady': self.vars.pad}
            if nopad := data.get('nopad'):
                if 'x' in nopad:
                    options['padx'] = 0
                if 'y' in nopad:
                    options['pady'] = 0
            if stick := data.get('sticky'):
                options['sticky']: str = stick
            self.elem[key].grid(**data['grid'], **options)
            if border := data.get('borderfull'):
                self.elem[key].config(**border)
            if text := data.get('tooltip'):
                self.tooltips.append(ToolTip(self.elem[key], msg=text, delay=0.25))

        # grid settings #
        cols, rows = self.elem[main].grid_size()
        for i in range(rows):
            self.elem[main].rowconfigure(i, **self.grid.row['default'] | self.grid.row.get(i, {}))
        for i in range(cols):
            self.elem[main].columnconfigure(i, **self.grid.col['default'] | self.grid.col.get(i, {}))

        # create menus #
        for key, data in self.menus.items():
            self.elem[key]: Menu = Menu(self.root, tearoff=0)
            if data.get('main'):
                self.root.config(menu=self.elem[key])
        for key, data in self.menus.items():
            for elem, args in data['elements']:
                match elem:
                    case 'menu':
                        self.elem[key].add_cascade(**args | {'menu': self.elem[args['menu']]})
                    case 'command':
                        self.elem[key].add_command(**args)
                    case 'separator' | _:
                        self.elem[key].add_separator()

        # event binds #
        for elem, (act, cmd) in self.binds:
            self.elem[elem].bind(act, cmd)

    @abstractmethod
    def pre(self) -> None:
        pass

    @abstractmethod
    def post(self) -> None:
        pass

    @abstractmethod
    def close(self) -> None:
        pass

    def throw_error(self, error_code: int, /, message: str='') -> None:
        if msg := self.vars.errors.get(error_code):
            showwarning(title='Błąd', message=msg)
        else:
            showerror(title='Nieznany błąd', message=message)


class PrinterApp(WindowApp):
    def select_file(self) -> None:
        # get new path #
        path: str = normpath(askopenfilename(
          title = 'Wybierz plik',
          initialdir = dirname(self.vars.file),
          filetypes = (('Plik JSON', '*.json'), ),
          multiple = False
          ))

        if not path:
            return

        # check if extension correct #
        if not self.vars.patterns.json.search(path):
            self.throw_error(4)
            return

        # set new file #
        self.vars.file = path

    def show_help(self) -> None:
        msg: str = (
          'Program do drukowania opisów.\n'
          '\n'
          '\u2022 Większość elemntów wyświetla opisy po najechaniu.\n'
          '\u2022 Po kolumnach można poruszać się za pomocą strzałek.\n'
          '\u2022 Elementy wybieramy na liście, a następnie klikamy\n'
          'w odpowiedni guzik w celu podjęcia akcji.\n'
          '\u2022 W nazwach punktów wolno używać jedynie\n'
          'liczb arabskich, liczb rzmyskich oraz kropek\n'
          '\u2022 Dane podstawowo zapisane są w pliku "data.json",\n'
          'można je przeładować, bądź wybrać inny plik danych.\n'
          '\u2022 Aplikacja ma powolny proces uruchamiania\n'
          'spowodowany trybem kompilacji.\n'
          '\n'
          'Python {}\n'
          'TKinter {}\n'
          'PyInstaller {}'
          ).format(self.vars.ver.py, self.vars.ver.tk, self.vars.ver.pyinst)
        showinfo(title='Pomoc', message=msg)

    def show_format(self) -> None:
        msg: str = (
          'Formatowanie opisu:\n'
          '\n'
          '<b> \u2026 </b> - pogrubienie\n'
          '<i> \u2026 </i> - kursywa\n'
          '<u> \u2026 </u> - podkreślenie\n'
          '<s> \u2026 </s> - przekreślenie\n'
          '\n'
          '<d> - okres (tylko w opisie projektu)\n'
          '<t> - podpis projektu (niedostępny)\n'
          '<p> - nazwa punktu (tylko w szablonie podpunktu)\n'
          '<o> - treść punktu (tylko w szablonie podpunktu)\n'
          '\n'
          '<br> - nowa linia\n'
          '\n'
          '{} - niedozwolone znaki nazwy pliku'
          ).format(' '.join(self.vars.bad_chars))
        showinfo(title='Formatowanie', message=msg)

    # other functions #
    @staticmethod
    def str2date(day: str, /) -> date:
        # date(*tuple(int(x) for x in day.split('.')[::-1]))
        return datetime.strptime(day, '%d.%m.%Y').date()

    @staticmethod
    def date2str(day: date, /) -> str:
        return day.strftime('%d.%m.%Y')

    @staticmethod
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

    @staticmethod
    def point2tuple(point: str, /) -> tuple[int, ...]:
        return tuple(PrinterApp.roman2int(n) for n in point.split('.'))

    @staticmethod
    def sort2return(iterable: Iterable[Any], /, *, key: Callable[[Any], Any]=None, reverse: bool=False) -> list[Any]:
        if isinstance(iterable, list):
            iterable.sort(key=key, reverse=reverse)
            return iterable
        return sorted(iterable, key=key, reverse=reverse)

    @staticmethod
    def get_state(state: bool, /) -> str:
        return 'normal' if state else 'disabled'

    @staticmethod
    def extract_dates(day: date, dates: Sequence[tuple[date, date]], /) -> tuple[date, date] | None:
        # check which timeframe is correct #
        for beg, end in dates:
            if beg <= day <= end:
                return beg, end
        return None

    @staticmethod
    def connect_dates(beg: date, end: date, /) -> str:
        return ' - '.join(PrinterApp.date2str(d) for d in (beg, end))

    @staticmethod
    def pair_up[T](seq: Sequence[T], /) -> tuple[tuple[T, T], ...]:
        return tuple(zip(seq[::2], seq[1::2]))

    @staticmethod
    def flatten[T](seq: Sequence[tuple[T, ...]], /) -> tuple[T, ...]:
        return sum(seq, ())

    @staticmethod
    def replace_mutiple(text: str, trans_table: dict[str, str], /) -> str:
        transpose: dict[str, str] = {escape(key): value for key, value in trans_table.items()}
        return recompile('|'.join(transpose.keys())).sub(lambda x: transpose[escape(x.group(0))], text)
