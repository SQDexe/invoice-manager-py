from abc import ABC, abstractmethod

from sys import version_info, executable
from datetime import date

from tkinter import Tk, PhotoImage, Menu
from tkinter.ttk import Style, Frame
from tkinter.messagebox import showerror, showinfo
from tktooltip import ToolTip

# from accessify impor protected

class WindowApp(ABC):
    def __init__(self, title, icon, minSize, maxSize):
        # root declaration #
        self.root = Tk()

        # declearing variables, and elements #
        self.vars = {
            'title': None,
            'size': {
                'min': (None, None),
                'max': (None, None)
                },
            'ver': {
                'py': '{}.{}.{}'.format(version_info.major, version_info.minor, version_info.micro),
                'tk': self.root.tk.call('info', 'patchlevel'),
                'pyinst': '6.8.0'
                },
            'errors': {
                # 0XX - file errors
                0: 'Błąd',
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
                # 6XX - tp errors
                601: 'Wybrany został projekt',
                602: 'Punkt(y) już został wybrany',
                603: 'Data poza zasięgiem'
                },
            'var': {},
            'pad': 5,
            'tags': ('b', 'i', 'u', 's'),
            'bad-chars': r'\/:*?"<>|',
            'style': Style()
            }
        self.elem = {}
        self.tooltips = []
        self.grid = {
            'row': { 'default': {'weight': 1, 'minsize': 40} },
            'col': { 'default': {'weight': 1, 'minsize': 50} }    
            }
        self.menus = {}
        self.binds = []

        # pre-generation funtions #
        self.pre()

        # window's settings #
        width, height = self.vars['size']['min']
        self.root.minsize(width, height)
        self.root.maxsize(*self.vars['size']['max'])
        self.root.resizable(True, True)
        self.root.geometry('{}x{}+{}+{}'.format(
            width,
            height,
            int((self.root.winfo_screenwidth() - width) / 2),
            int((self.root.winfo_screenheight() - height) / 2)
            ))
        self.root.title(self.vars['title'])
        self.root.iconbitmap(executable)

        # protocols #
        self.root.protocol('WM_DELETE_WINDOW', self.close)

        # prepare elements #
        self.prep_elems()

        # post-generation funtions #
        self.post()

        # start program #
        self.root.mainloop()

    def prep_elems(self):
        main = 'main'

        # mainframe #
        self.elem[main] = Frame(self.root)
        self.elem[main].pack(fill='both', expand=True)

        # create, and configure elements #
        for key, data in ((x, y) for x, y in self.elem.items() if x != main):
            self.elem[key] = data['type'](self.elem[main], **data.get('args', {}))
            options = {'padx': self.vars['pad'], 'pady': self.vars['pad']}
            if nopad := data.get('nopad'):
                if 'x' in nopad:
                    options['padx'] = 0
                if 'y' in nopad:
                    options['pady'] = 0
            if stick := data.get('sticky'):
                options['sticky'] = stick
            self.elem[key].grid(**data['grid'], **options)
            if border := data.get('borderfull'):
                self.elem[key].config(**border)
            if text := data.get('tooltip'):
                self.tooltips.append(ToolTip(self.elem[key], msg=text, delay=0.25))

        # grid settings #
        cols, rows = self.elem[main].grid_size()
        for i in range(rows):
            self.elem[main].rowconfigure(i, **self.grid['row']['default'] | self.grid['row'].get(i, {}))
        for i in range(cols):
            self.elem[main].columnconfigure(i, **self.grid['col']['default'] | self.grid['col'].get(i, {}))

        # create menus #
        for key, data in self.menus.items():
            self.elem[key] = Menu(self.root, tearoff=0)
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
            self.elem.get(elem).bind(act, cmd)

    @abstractmethod
    def pre(self):
        pass

    @abstractmethod
    def post(self):
        pass

    def show_help(self):
        msg = (
            'Program do drukowania opisów.\n'
            '\n'
            '\u2022 Większość elemntów wyświetla opisy po najechaniu.\n'
            '\u2022 Po kolumnach można poruszać się za pomocą strzałek.\n'
            '\u2022 Elementy wybieramy na liście, a następnie klikamy\n'
            'w odpowiedni guzik w celu podjęcia akcji.\n'
            '\u2022 W nazwach punktów wolno używać jedynie\n'
            'liczb arabskich, liczb rzymskich oraz kropek\n'
            '\u2022 Dane podstawowo zapisane są w pliku "data.json",\n'
            'można je przeładować, bądź wybrać inny plik danych.\n'
            '\u2022 Aplikacja ma powolny proces uruchamiania\n'
            'spowodowany trybem kompilacji.\n'
            '\n'
            'Python {}\n'
            'TKinter {}\n'
            'PyInstaller {}'
            ).format(self.vars['ver']['py'], self.vars['ver']['tk'], self.vars['ver']['pyinst'])
        showinfo(title='Pomoc', message=msg)

    @abstractmethod
    def close(self):
        pass

    def throw_error(self, errorCode: int):
        txt = 'Błąd' if self.vars['errors'].get(errorCode) else 'Nieznany błąd'
        msg = self.vars['errors'].get(errorCode, errorCode)
        showerror(title='Błąd', message=msg)

    # other functions #
    @staticmethod
    def str2date(d: str) -> date:
        return date(*reversed(tuple(int(x) for x in d.split('.'))))

    @staticmethod
    def date2str(d: date) -> str:
        return ('{:02d}.{:02d}.{:04d}'.format(d.day, d.month, d.year)

    @staticmethod
    def roman2int(n: str) -> int:
        try:
            return int(n)
        except Exception:
            roman, rest, n = {'I': 1, 'V': 5, 'X': 10, 'L': 50, 'C': 100, 'D': 500, 'M': 1000}, 0, n.upper()
            if not set(roman.keys()).issuperset(n):
                return 0
            for i in range(len(n) - 1, -1, -1):
                num = roman.get(n[i], 0)
                rest = rest - num if 3 * num < rest else rest + num
            return rest

    @staticmethod
    def get_state(state: bool) -> str:
        return 'normal' if state else 'disabled'

    @staticmethod
    def get_date(date: str, dates: list[date] | tuple[date], mode: str='raw') -> str | tuple[date] | tuple[date] | None:
        # check which timeframe is correct #
        date, *dates = tuple(WindowApp.str2date(x) for x in (date, *dates))
        if chosen := tuple((dates[i], dates[i + 1]) for i in range(0, len(dates), 2) if dates[i] <= date <= dates[i + 1]):
            chosen_data = chosen[0]
            match mode:
                case 'string':
                    return ' - '.join(WindowApp.date2str(x) for x in chosen_data)
                case 'int':
                    return tuple((x.day, x.month, x.year) for x in chosen_data)
                case 'raw' | _:
                    return chosen_data
        else:
            return None
