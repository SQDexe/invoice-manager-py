from os import getcwd
from os.path import isfile
from json import loads, dumps
from datetime import date

# from accessify import private
from base64 import b64decode

from tkinter import Tk, StringVar as StrVar, PhotoImage, Menu, Text
from tkinter.ttk import Style, Frame, Entry, Button, Treeview, Scrollbar, Label
from tkinter.messagebox import showerror, showinfo, askokcancel
from tkinter.filedialog import askopenfilename
from tkcalendar import DateEntry
from tktooltip import ToolTip

class DataEditor:
    def __init__(self):
        # root declaration #
        self.root = Tk()

        # declearing variables, and elements #
        self.vars = {
            'title': 'Data Editor',
            'icon': 'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQBAMAAADt3eJSAAAAIGNIUk0AAHomAACAhAAA+gAAAIDoAAB1MAAA6mAAADqYAAAXcJy6UTwAAAAVUExURQAAAP///9vb2wAAAP/vlOr/Sv+UlPDI24oAAAABdFJOUwBA5thmAAAAAWJLR0QB/wIt3gAAAAd0SU1FB+cLCBAqAAa89zwAAAAtSURBVAjXY2DABIyCggoQhrGxEoMBRERJwQEiYiwQABERhDBIFlGAMRKw2A0ARWIITkVLLqYAAAAldEVYdGRhdGU6Y3JlYXRlADIwMjMtMTEtMDhUMTY6NDE6NTkrMDA6MDBL+4JDAAAAJXRFWHRkYXRlOm1vZGlmeQAyMDIzLTExLTA4VDE2OjQxOjU5KzAwOjAwOqY6/wAAACh0RVh0ZGF0ZTp0aW1lc3RhbXAAMjAyMy0xMS0wOFQxNjo0MjowMCswMDowMFv865QAAAAASUVORK5CYII=',
            'size': {
                'min': (400, 450),
                'max': (650, 650)
                },
            'file': '{}\\{}'.format(getcwd(), 'data.json'),
            'pad': 5,
            'unsaved': False,
            'var': {
                'name': StrVar(),
                'date-beg': StrVar(),
                'date-end': StrVar()
                },
            'style': Style(),
            'tags': ('b', 'i', 'u', 's')
            }
        self.elem = {
            'tree-points': {
                'type': Treeview,
                'args': {'selectmode': 'browse', 'show': 'tree'},
                'grid': {'row': 0, 'column': 0, 'rowspan': 6},
                'sticky': 'NWES'
                },
            'scroll-points': {
                'type': Scrollbar,
                'args': {'orient': 'vertical'},
                'grid': {'row': 0, 'column': 0, 'rowspan': 6},
                'sticky': 'NES'
                },
            'tree-dates': {
                'type': Treeview,
                'args': {'selectmode': 'browse', 'show': 'tree'},
                'grid': {'row': 0, 'column': 1, 'columnspan': 5},
                'sticky': 'NWES'
                },
            'scroll-dates': {
                'type': Scrollbar,
                'args': {'orient': 'vertical'},
                'grid': {'row': 0, 'column': 5},
                'sticky': 'NES'
                },
            'cal-dates-beg': {
                'type': DateEntry,
                'args': {'date_pattern': 'd.M.y', 'locale': 'pl_PL', 'textvariable': self.vars.get('var').get('date-beg')},
                'grid': {'row': 1, 'column': 1, 'columnspan': 2},
                'tooltip': 'Data początku okresu'
                },
            'cal-dates-end': {
                'type': DateEntry,
                'args': {'date_pattern': 'd.M.y', 'locale': 'pl_PL', 'textvariable': self.vars.get('var').get('date-end')},
                'grid': {'row': 2, 'column': 1, 'columnspan': 2},
                'tooltip': 'Data końca okresu'
                },
            'btn-dates-add': {
                'type': Button,
                'args': {'text': '\ud83d\udfa4', 'command': self.add_date},
                'grid': {'row': 1, 'column': 3},
                'tooltip': 'Dodaj wskazane daty'
                },
            'btn-dates-del': {
                'type': Button,
                'args': {'text': '\u232b', 'command': self.delete_date},
                'grid': {'row': 1, 'column': 5},
                'tooltip': 'Usuń daty'
                },
            'blank': {
                'type': Label,
                'args': {'text': '', 'background': 'white'},
                'grid': {'row': 3, 'column': 4}
                },
            'entry-points': {
                'type': Entry,
                'args': {'textvariable': self.vars.get('var').get('name')},
                'grid': {'row': 4, 'column': 1, 'columnspan': 5},
                'sticky': 'NWES'
                },
            'btn-points-add-project': {
                'type': Button,
                'args': {'text': '\ud83d\uddbf', 'command': self.add_catalogue},
                'grid': {'row': 5, 'column': 1},
                'tooltip': 'Dodaj projekt'
                },
            'btn-points-add-entry': {
                'type': Button,
                'args': {'text': '\ud83d\udfa4', 'command': self.add_item},
                'grid': {'row': 5, 'column': 2},
                'tooltip': 'Dodaj punkt kosztorysu'
                },
            'btn-points-change': {
                'type': Button,
                'args': {'text':'\u270e','command': self.change_item},
                'grid': {'row': 5, 'column': 4},
                'tooltip': 'Zmień nazwę elementu'
                },
            'btn-points-del': {
                'type': Button,
                'args': {'text':'\u232b','command': self.delete_item},
                'grid': {'row': 5, 'column': 5},
                'tooltip': 'Usuń element'
                },
            'txt-field': {
                'type': Text,
                'args': {'wrap': 'word'},
                'grid': {'row': 6, 'column': 0, 'columnspan': 6},
                'sticky': 'NWES',
                'borderfull': {'highlightthickness': 1, 'highlightbackground': 'gray'}
                },
            'btn-text': {
                'type': Button,
                'args': {'text': '\u166d', 'state': 'disabled', 'command': self.save_text},
                'grid': {'row': 7, 'column': 1},
                'tooltip': 'Zapisz tekst'
                },
            'btn-save': {
                'type': Button,
                'args': {'text': 'Zapisz', 'command': self.save_file, 'state': 'disabled'},
                'grid': {'row': 7, 'column': 4, 'columnspan': 2},
                'tooltip': 'Zapisz zmainy w pliku'
                }
            }
        self.tooltips = []

        # window's settings #
        width, height = self.vars.get('size').get('min')
        self.root.minsize(width, height)
        self.root.maxsize(*self.vars.get('size').get('max'))
        self.root.resizable(True, True)
        self.root.geometry('{}x{}+{}+{}'.format(
            width,
            height,
            int((self.root.winfo_screenwidth() - width) / 2),
            int((self.root.winfo_screenheight() - height) / 2)
            ))
        self.root.title(self.vars.get('title'))
        self.root.iconphoto(False, PhotoImage(data=b64decode(self.vars.get('icon'))))

        # protocols #
        self.root.protocol('WM_DELETE_WINDOW', self.close)

        # prepare elements #
        self.prep_elems()
        self.set_data()

        # start program #
        self.root.mainloop()

    # decorator for changes #
    def safecheck(f):
        def wrapper(self):
            f(self)
            self.elem.get('btn-save').config(state='normal')
            self.vars.update({'unsaved': True})
        return wrapper

    def prep_elems(self):
        main = 'main'
        grid = {
            'row': {
                'default': {'weight': 1, 'minsize': 40},
                0: {'weight': 3, 'minsize': 100},
                6: {'weight': 3, 'minsize': 100}
                },
            'col': {
                'default': {'weight': 1, 'minsize': 50},
                0: {'weight': 3, 'minsize': 150}    
                }
            }
        menus = {
            'menu-main': {
                'main': True,
                'elements': [
                    ('menu', {'label': 'Plik', 'menu': 'menu-file'}),
                    ('command', {'label': 'Formatowanie', 'command': self.show_format}),
                    ('command', {'label': 'Pomoc', 'command': self.show_help})
                    ]
                },
            'menu-file': {
                'elements': [
                    ('command', {'label': 'Wybierz...', 'command': self.select_file}),
                    ('command', {'label': 'Przeładuj', 'command': self.reload}),
                    ('separator', None),
                    ('command', {'label': 'Wyjdź', 'command': self.close})
                    ]
                }
            }
        binds = [
            ('tree-points', ('<<TreeviewSelect>>', self.points_select)),
            ('txt-field', ('<Key>', self.text_change))
            ]

        # mainframe #
        self.elem.update({main: Frame(self.root)})
        self.elem.get(main).pack(fill = 'both', expand = True)

        # create, and configure elements #
        for key, data in ((x, y) for x, y in self.elem.items() if x != main):
            self.elem.update({key: data.get('type')(self.elem.get(main), **data.get('args', {}))})
            options = {'padx': self.vars.get('pad'), 'pady': self.vars.get('pad')}
            if nopad := data.get('nopad'):
                if 'x' in nopad:
                    options.update({'padx': 0})
                if 'y' in nopad:
                    options.update({'pady': 0})
            if stick := data.get('sticky'):
                options.update({'sticky': stick})
            self.elem.get(key).grid(**data.get('grid'), **options)
            if border := data.get('borderfull'):
                self.elem.get(key).config(**border)
            if text := data.get('tooltip'):
                self.tooltips.append(ToolTip(self.elem.get(key), msg=text, delay=0.25))

        # grid settings #
        cols, rows = self.elem.get(main).grid_size()
        for i in range(rows):
            self.elem.get(main).rowconfigure(i, **grid.get('row').get('default') | grid.get('row').get(i, {}))
        for i in range(cols):
            self.elem.get(main).columnconfigure(i, **grid.get('col').get('default') | grid.get('col').get(i, {}))

        # create menus #
        for key, data in menus.items():
            self.elem.update({key: Menu(self.root, tearoff=0)})
            if data.get('main'):
                self.root.config(menu=self.elem.get(key))
        for key, data in menus.items():
            for elem, args in data.get('elements'):
                match elem:
                    case 'menu':
                        self.elem.get(key).add_cascade(**args | {'menu': self.elem.get(args.get('menu'))})
                    case 'command':
                        self.elem.get(key).add_command(**args)
                    case 'separator' | _:
                        self.elem.get(key).add_separator()

        # event binds #
        for elem, (act, cmnd) in binds:
            self.elem.get(elem).bind(act, cmnd)

        # styles, other settings, and actions #
        self.vars.get('style').configure('TFrame', background='white')
        self.elem.get('tree-points').configure(yscrollcommand=self.elem.get('scroll-points').set)
        self.elem.get('tree-dates').configure(yscrollcommand=self.elem.get('scroll-dates').set)
        self.elem.get('scroll-points').configure(command=self.elem.get('tree-points').yview)
        self.elem.get('scroll-dates').configure(command=self.elem.get('tree-dates').yview)

    def set_data(self):
        # check if file exists #
        if not isfile(self.vars.get('file')):
            self.throw_error(6)
            return

        # try to read data #
        try:
            with open(self.vars.get('file'), 'rt', encoding='utf-8') as f:
                data = loads(f.read())
                for i, project in enumerate(data):
                    vals = [project.get('description')]
                    for x in project.get('dates'):
                        vals.extend((x.get('from'), x.get('to')))
                    self.elem.get('tree-points').insert('', 'end', i, text=project.get('name'), values=vals, open=False, tags=['catalogue'])
                    for point in project.get('points'):
                        self.elem.get('tree-points').insert(i, 'end', text=point.get('point'), values=[point.get('text')])

        except Exception as e:
            self.throw_error(e)

    def clear_data(self):
        # clear data #
        self.elem.get('tree-points').delete(*self.elem.get('tree-points').get_children())
        self.elem.get('tree-dates').delete(*self.elem.get('tree-dates').get_children())
        self.elem.get('txt-field').delete('1.0', 'end')
        
    def points_select(self, event):
        iid = self.elem.get('tree-points').focus()

        # check wherever iid correct #
        if not iid:
            return

        # set texts #
        self.vars.get('var').get('name').set(self.elem.get('tree-points').item(iid, 'text'))
        self.elem.get('txt-field').delete('1.0', 'end')
        self.elem.get('txt-field').insert('1.0', self.elem.get('tree-points').item(iid, 'values')[0])

        # get catalogue #
        if not self.elem.get('tree-points').tag_has('catalogue', iid):
            iid = self.elem.get('tree-points').parent(iid)

        # set dates #	
        vals = self.elem.get('tree-points').item(iid, 'values')
        for item in self.elem.get('tree-dates').get_children():
            self.elem.get('tree-dates').delete(item)
        for i in range(1, len(vals), 2):
            self.elem.get('tree-dates').insert('', 'end', text='{} - {}'.format(vals[i], vals[i + 1]), values=[iid, vals[i], vals[i + 1]])

    def text_change(self, event):
        # check wherever iid correct #
        if not self.elem.get('tree-points').focus():
            return

        # set button #
        self.elem.get('btn-text').config(text='\u2713', state='normal')

    @safecheck
    def save_text(self):
        iid = self.elem.get('tree-points').focus()

        # check wherever element is focused #
        if not iid:
            self.throw_error(0)
            return

        # assign text #
        text = self.elem.get('txt-field').get('1.0', 'end-1c')
        vals = list(self.elem.get('tree-points').item(iid, 'values'))
        vals[0] = text.strip()
        self.elem.get('tree-points').item(iid, values=vals)

        # set button #
        self.elem.get('btn-text').config(text='\u166d', state='disabled')

    @safecheck
    def delete_item(self):
        iid = self.elem.get('tree-points').focus()

        # check wherever element is focused #
        if not iid:
            self.throw_error(0)
            return

        # delete item #
        self.elem.get('txt-field').delete('1.0', 'end')
        self.elem.get('tree-points').delete(iid)

        # clear view #
        self.elem.get('tree-dates').delete(*self.elem.get('tree-dates').get_children())

    @safecheck
    def add_catalogue(self):
        name = self.vars.get('var').get('name').get()

        # check if name was given #
        if not name:
            self.throw_error(1)
            return

        # check if name not taken #
        if name in tuple(self.elem.get('tree-points').item(x, 'text') for x in self.elem.get('tree-points').get_children()):
            self.throw_error(2)
            return

        # add catalogue #
        focus = self.elem.get('tree-points').insert('', 'end', text=name, values=[''], tags=['catalogue'])
        self.elem.get('tree-points').selection_set(focus)
        self.elem.get('tree-points').focus(focus)

    @safecheck
    def add_item(self):
        iid, name, index = self.elem.get('tree-points').focus(), self.vars.get('var').get('name').get(), 'end'

        # check wherever element is focused #
        if not iid:
            self.throw_error(0)
            return

        # check if name was given #
        if not name:
            self.throw_error(1)
            return

        # get catalogue #
        if not self.elem.get('tree-points').tag_has('catalogue', iid):
            index, iid = self.elem.get('tree-points').index(iid), self.elem.get('tree-points').parent(iid)

        # check if name not taken #	
        if name in tuple(self.elem.get('tree-points').item(x, 'text') for x in self.elem.get('tree-points').get_children(iid)):
            self.throw_error(2)
            return

        # insert element #
        focus = self.elem.get('tree-points').insert(iid, index, text=name, values=[''])
        self.elem.get('tree-points').selection_set(focus)
        self.elem.get('tree-points').focus(focus)

    @safecheck
    def change_item(self):
        iid, name = self.elem.get('tree-points').focus(), self.vars.get('var').get('name').get()

        # check wherever element is focused #
        if not iid:
            self.throw_error(0)
            return

        # check if name was given #
        if not name:
            self.throw_error(1)
            return

        # get catalogue #
        parent = '' if self.elem.get('tree-points').tag_has('catalogue', iid) else self.elem.get('tree-points').parent(iid)

        # check if name not taken #
        if name in tuple(self.elem.get('tree-points').item(x, 'text') for x in self.elem.get('tree-points').get_children(parent)):
            self.throw_error(2)
            return

        # set new name #
        self.elem.get('tree-points').item(iid, text=name)

    @safecheck
    def delete_date(self):
        iid = self.elem.get('tree-dates').focus()

        # check wherever element is focused #
        if not iid:
            self.throw_error(0)
            return

        # update values #
        parent_iid = self.elem.get('tree-dates').item(iid, 'values')[0]
        vals = [self.elem.get('tree-points').item(parent_iid, 'values')[0]]
        for date in [self.elem.get('tree-dates').item(x, 'values')[1:3] for x in self.elem.get('tree-dates').get_children() if x != iid]: 
            vals.extend(date)
        self.elem.get('tree-points').item(parent_iid, values=vals)

        # delete item #
        self.elem.get('tree-dates').delete(iid)

    @safecheck
    def add_date(self):
        # get parent #
        parent_iid = self.elem.get('tree-points').focus()
        if not self.elem.get('tree-points').tag_has('catalogue', parent_iid):
            parent_iid = self.elem.get('tree-points').parent(parent_iid)

        # check wherever element is focused #
        if not parent_iid:
            self.throw_error(0)
            return

        # check if range is correct #
        dates = tuple(self.str2date(x) for x in (self.vars.get('var').get('date-beg').get(), self.vars.get('var').get('date-end').get()))
        if dates[1] <= dates[0]:
            self.throw_error(4)
            return

        # check for other dates #
        index = 'end'
        if self.elem.get('tree-dates').get_children():
            # check if date not taken #
            current_dates = tuple((iid, tuple(self.str2date(z) for z in (x, y))) for iid, (x, y) in tuple((iid, self.elem.get('tree-dates').item(iid, 'values')[1:3]) for iid in self.elem.get('tree-dates').get_children()))
            if not all(dates[1] < start or end < dates[0] for _, (start, end) in current_dates):
                self.throw_error(5)
                return

            # get right index #
            ommit = (0, len(current_dates) - 1)
            if dates[1] < current_dates[ommit[0]][1][0]:
                index = 0
            elif current_dates[ommit[1]][1][0] < dates[0]:
                index = ommit[1] + 1
            else:
                for i, (iid, _) in enumerate(current_dates):
                    if i in ommit:
                        continue
                    if current_dates[i - 1][1][1]  < dates[0] and dates[1] < current_dates[i][1][0]:
                        index = i

        # insert in right postion #
        dates = tuple(self.date2str(x) for x in dates)
        self.elem.get('tree-dates').insert('', index, text=' - '.join(dates), values=[parent_iid, dates[0], dates[1]])

        # update values #
        vals = [self.elem.get('tree-points').item(parent_iid, 'values')[0]]
        for date in tuple(self.elem.get('tree-dates').item(x, 'values')[1:3] for x in self.elem.get('tree-dates').get_children()): 
            vals.extend(date)
        self.elem.get('tree-points').item(parent_iid, values=vals)

    def save_file(self):
        file = []

        # get iterable values, and make obj #
        for project in self.elem.get('tree-points').get_children():
            vals, dates, points = self.elem.get('tree-points').item(project, 'values'), [], []
            for i in range(1, len(vals), 2):
                dates.append({'from': vals[i], 'to': vals[i + 1]})
            for item in self.elem.get('tree-points').get_children(project):
                points.append({
                    'point': self.elem.get('tree-points').item(item, 'text'),
                    'text': self.elem.get('tree-points').item(item, 'values')[0]
                    })
            points.sort(key=lambda x: tuple(self.exint(y) for y in x.get('point').split('.')))
            file.append({
                'name': self.elem.get('tree-points').item(project, 'text'),
                'description': vals[0],
                'dates': dates,
                'points': points
                })
        file.sort(key=lambda x: x.get('name'))

        # try to save file #
        try:
            with open(self.vars.get('file'), 'wt', encoding='utf-8') as f:
                f.write(dumps(file))
            showinfo(title='Zapisywanie', message='Sukces')
            self.elem.get('btn-save').config(state='disabled')
            self.vars.update({'unsaved': False})

        except Exception as e:
            self.throw_error(e)

    def close(self):
        if self.vars.get('unsaved'):
            if self.ask_changes():
                self.root.destroy()	
        else:
            self.root.destroy()

    def show_format(self):
        msg = \
            'Formatowanie opisu:\n' \
            '\n' \
            '<b> ... </b> - pogrubienie\n' \
            '<i> ... </i> - kursywa\n' \
            '<u> ... </u> - podkreślenie\n' \
            '<s> ... </s> - przekreślenie\n' \
            '\n' \
            '<d> - okres (tylko w opisie projektu)\n' \
            '\n' \
            '<br> - nowa linia'
        showinfo(title='Formatowanie', message=msg)

    def show_help(self):
        msg = \
            'Program do edytowania danych opisów.\n' \
            '\n' \
            '\u2022 Większość elemntów wyświetla opisy po najechaniu.\n' \
            '\u2022 Po kolumnach można poruszać się za pomocą strzałek.\n' \
            '\u2022 Elementy wybieramy na liście, a następnie klikamy\n' \
            'w odpowiedni guzik w celu podjęcia akcji.\n' \
            '\u2022 Dane podstawowo zapisane są w pliku "data.json",\n' \
            'można je przeładować, bądź wybrać inny plik danych.\n' \
            '\u2022 Aplikacja uruchamia się stosunkowo powoli.\n' \
            '\n' \
            'Program napisany w Python, z pomocą TKinter.'
        showinfo(title='Pomoc', message=msg)

    def select_file(self):
        # get new path #
        path = askopenfilename(title='Wybierz plik', initialdir='\\'.join(self.vars.get('file').split('\\')[:-1]), filetypes=(('Plik JSON', '.json'), ), multiple=False).replace('/', '\\')
        if not path:
            return

        # check if extension correct #
        if '.json' not in path[-5:]:
            self.throw_error(7)
            return

        # set new file #
        self.vars.update({'file': path})

    def reload(self):
        # check for changes #
        if self.vars.get('unsaved'):
            if not self.ask_changes():
                return

        # reload data #
        self.clear_data()
        self.set_data()
        self.elem.get('btn-save').config(state='disabled')
        self.vars.update({'unsaved': False})

    def ask_changes(self):
        return askokcancel(title='Niezapisane zmiany', message='Czy chcesz kontynuować?')

    def throw_error(self, error):
        match error:
            case 0:
                msg = 'Nie wybrano elementu'
            case 1:
                msg = 'Nie podano nazwy'
            case 2:
                msg = 'Nazwa zajęta'
            case 3:
                msg = 'Złe formatowanie tekstu'
            case 4:
                msg = 'Zły okres'
            case 5:
                msg = 'Okres zajęty'
            case 6:
                msg = 'Brak podstawowego pliku danych'
            case 7:
                msg = 'Złe rozszerzenie pliku'
            case _:
                msg = error
        showerror(title='Błąd', message=msg)


    # other functions #
    @staticmethod
    def str2date(d):
        return date(*reversed(tuple(int(x) for x in d.split('.'))))

    @staticmethod
    def date2str(d):
        return '.'.join(*(d.day, d.month, d.year))

    @staticmethod
    def exint(n):
        try:
            return int(n)
        except Exception:
            roman, rest, n = {'I': 1, 'V': 5, 'X': 10, 'L': 50, 'C': 100, 'D': 500, 'M': 1000}, 0, n.upper()
            for i in range(len(n) - 1, -1, -1):
                num = roman.get(n[i], 0)
                rest = rest - num if 3 * num < rest else rest + num
            return rest

    @staticmethod
    def get_state(state):
        return 'normal' if state else 'disabled'


if __name__ == '__main__':
	DataEditor()
