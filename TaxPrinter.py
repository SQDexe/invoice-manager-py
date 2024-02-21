from os import getcwd
from os.path import isfile
from io import StringIO
from json import loads
from datetime import date
from re import findall, split, sub

from base64 import b64decode
from unidecode import unidecode

from docx import Document
from docx.shared import Pt

from tkinter import Tk, StringVar as StrVar, BooleanVar as BoolVar, PhotoImage, Menu, Text
from tkinter.messagebox import showerror, showinfo, askokcancel
from tkinter.filedialog import askopenfilename, askdirectory
from tkinter.ttk import Style, Frame, Entry, Button, Treeview, Scrollbar, Checkbutton, Radiobutton
from tkcalendar import DateEntry
from tktooltip import ToolTip

class TaxPrinter:
    def __init__(self):
        # root declaration #
        self.root = Tk()

        # declearing variables, and elements #
        self.vars = {
            'title': 'Tax Printer',
            'icon': 'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQBAMAAADt3eJSAAAAIGNIUk0AAHomAACAhAAA+gAAAIDoAAB1MAAA6mAAADqYAAAXcJy6UTwAAAAeUExURQAAAP///9vb2wAAAI+Pj/+FhYv/hf//a3BwcPDw8ChUvUYAAAABdFJOUwBA5thmAAAAAWJLR0QB/wIt3gAAAAd0SU1FB+cLCBAqL61tymUAAAA9SURBVAjXY2BAAEFBIQhD2NhIASaiBBUxNoSKCApiF3EBAyDDNQVIlyCLuBgDAZihBAQgRsdMIOhAMRACAAYpDjSL+1GnAAAAJXRFWHRkYXRlOmNyZWF0ZQAyMDIzLTExLTA4VDE2OjQyOjQ2KzAwOjAwmi5JNwAAACV0RVh0ZGF0ZTptb2RpZnkAMjAyMy0xMS0wOFQxNjo0Mjo0NiswMDowMOtz8YsAAAAodEVYdGRhdGU6dGltZXN0YW1wADIwMjMtMTEtMDhUMTY6NDI6NDcrMDA6MDAaEdvgAAAAAElFTkSuQmCC',
            'size': {
                'min': (400, 450),
                'max': (650, 650)
                },
            'file': '{}\\{}'.format(getcwd(), 'data.json'),
            'pad': 5,
            'var': {
                'date': StrVar(),
                'filename': StrVar(value='plik'),
                'filepath': StrVar(value=getcwd()),
                'title-text': StrVar(value='<b><t></b>'),
                'point-text': StrVar(value='<b>Punkt <p></b> „<o>”'),
                'cash': BoolVar(),
                'addons': BoolVar(),
                'cash-text': StrVar(value='Suma <b>............... ,-</b>, '),
                'addons-text': StrVar(value=' - <br>'),
                'cash-mode': StrVar(),
                'opening-mode': StrVar(),
                'facture-text': StrVar(value='Akapit przykładowy 1.\nTutaj wpisać tekst na, który ma pojawić się na początku dokumentu'),
                'contract-text': StrVar(value='Akapit przykładowy 2.\nTutaj wpisać tekst na, który ma pojawić się na początku dokumentu')
                },
            'style': Style(),
            'tags': ('b', 'i', 'u', 's'),
	        'badChars': r'\/:*?"<>|'
            }
        self.elem = {
            'tree-all': {
                'type': Treeview,
                'args': {'selectmode': 'extended', 'show': 'tree'},
                'grid': {'row': 0, 'column': 0, 'rowspan': 5, 'columnspan': 2},
                'sticky': 'NWES'
                },
            'scroll-all': {
                'type': Scrollbar,
                'args': {'orient': 'vertical'},
                'grid': {'row': 0, 'column': 1, 'rowspan': 5},
                'sticky': 'NES'
                },
            'btn-add': {
                'type': Button,
                'args': {'text': '\u25b6', 'command': self.__add},
                'grid': {'row': 0, 'column': 2},
                'tooltip': 'Dodaj element(y)'
                },
            'btn-add-all': {
                'type': Button,
                'args': {'text': '\u25b6\u25b6', 'command': self.__add_all},
                'grid': {'row': 1, 'column': 2},
                'tooltip': 'Dodaj wszystkie elementy'
                },
            'btn-remove': {
                'type': Button,
                'args': {'text': '\u25c0', 'command': self.__remove},
                'grid': {'row': 3, 'column': 2},
                'tooltip': 'Usuń element(y)'
                },
            'btn-remove-all': {
                'type': Button,
                'args': {'text': '\u25c0\u25c0', 'command': self.__remove_all},
                'grid': {'row': 4, 'column': 2},
                'tooltip': 'Usuń wszystkie elementy'
                },
            'tree-selected': {
                'type': Treeview,
                'args': {'selectmode': 'extended', 'show': 'headings', 'columns': ('project', 'point')},
                'grid': {'row': 0, 'column': 3, 'rowspan': 5, 'columnspan': 2},
                'sticky': 'NWES'
                },
            'txt-opening': {
                'type': Text,
                'args': {'wrap': 'word'},
                'grid': {'row': 5, 'column': 0, 'columnspan': 5},
                'sticky': 'NWES',
                'borderfull': {'highlightthickness': 1, 'highlightbackground': 'gray'}
                },
            'radbtn-facture': {
                'type': Radiobutton,
                'args': {'text': 'Faktura', 'variable': self.vars.get('var').get('opening-mode'), 'command': self.__set_text, 'value': 'facture'},
                'grid': {'row': 6, 'column': 3},
                'sticky': 'W'
                },
            'radbtn-contract': {
                'type': Radiobutton,
                'args': {'text': 'Umowa', 'variable': self.vars.get('var').get('opening-mode'), 'command': self.__set_text, 'value': 'contract'},
                'grid': {'row': 6, 'column': 4},
                'sticky': 'W'
                },
            'entry-point': {
                'type': Entry,
                'args': {'textvariable': self.vars.get('var').get('point-text')},
                'grid': {'row': 6, 'column': 0, 'columnspan': 2},
                'tooltip': 'Tekst podpunktu',
                'sticky': 'NWES'
                },
            'chkbtn-cash': {
                'type': Checkbutton,
                'args': {'text': 'Dodaj pole kwoty', 'variable': self.vars.get('var').get('cash'), 'command': lambda: self.__toggle('cash')},
                'grid': {'row': 7, 'column': 0},
                'sticky': 'W'
                },
            'entry-cash': {
                'type': Entry,
                'args': {'textvariable': self.vars.get('var').get('cash-text'), 'state': 'disabled'},
                'grid': {'row': 7, 'column': 1},
                'tooltip': 'Tekst kwoty',
                'sticky': 'NWES'
                },
            'radbtn-auto': {
                'type': Radiobutton,
                'args': {'text': 'Auto', 'variable': self.vars.get('var').get('cash-mode'), 'value': 'auto'},
                'grid': {'row': 8, 'column': 0},
                'sticky': 'W'
                },
            'radbtn-all': {
                'type': Radiobutton,
                'args': {'text': 'Wszystko', 'variable': self.vars.get('var').get('cash-mode'), 'value': 'all'},
                'grid': {'row': 8, 'column': 1},
                'sticky': 'W'
                },
            'chkbtn-addons': {
                'type': Checkbutton,
                'args': {'text': 'Dodaj myślniki', 'variable': self.vars.get('var').get('addons'), 'command': lambda: self.__toggle('addons')},
                'grid': {'row': 9, 'column': 0},
                'sticky': 'W'
                },
            'entry-addons': {
                'type': Entry,
                'args': {'textvariable': self.vars.get('var').get('addons-text'), 'state': 'disabled'},
                'grid': {'row': 9, 'column': 1},
                'tooltip': 'Tekst notatki',
                'sticky': 'NWES'
                },
            'cal-select': {
                'type': DateEntry,
                'args': {'date_pattern': 'd.M.y', 'locale': 'pl_PL', 'textvariable': self.vars.get('var').get('date')},
                'grid': {'row': 7, 'column': 3},
                'tooltip': 'Data wystawienia opisu'
                },
            'btn-name': {
                'type': Button,
                'args': {'text': '\u270e', 'state': 'disabled', 'command': self.__make_name},
                'grid': {'row': 7, 'column': 4},
                'tooltip': 'Wygeneruj nazwę do jednego podpunktu'
                },
            'entry-filename': {
                'type': Entry,
                'args': {'textvariable': self.vars.get('var').get('filename')},
                'grid': {'row': 8, 'column': 3, 'columnspan': 2},
                'tooltip': 'Nazwa pliku',
                'sticky': 'NWES'
                },
            'btn-filepath': {
                'type': Button,
                'args': {'text': '...', 'command': self.__set_path},
                'grid': {'row': 9, 'column': 3},
                'tooltip': 'Wybierz ścieżkę docelową',
                'sticky': 'W'
                },
            'btn-print': {
                'type': Button,
                'args': {'text': '\ud83d\uddb6', 'state': 'disabled', 'command': self.__print},
                'grid': {'row': 9, 'column': 4},
                'tooltip': 'Drukuj opis'
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

        # prepare elements #
        self.__prep_elems()
        self.__set_data()

        # start program #
        self.root.mainloop()

    # decorators #
    def __check(f):
        def wrapper(self, *args):
            f(self, *args)

            childs = self.elem.get('tree-selected').get_children()
            self.elem.get('btn-print').config(state='normal' if childs else 'disabled')
            self.elem.get('btn-name').config(state='normal' if len(childs) == 1 else 'disabled')

        return wrapper

    def __sort(f):
        def wrapper(self, *args):
            f(self, *args)

            # sort by point, then by project #
            vals = [(iid, self.elem.get('tree-selected').set(iid, 'point'), self.elem.get('tree-selected').set(iid, 'project')) for iid in self.elem.get('tree-selected').get_children()]
            vals.sort(key=lambda x: (x[2], tuple(self.__exint(y) for y in x[1].split('.'))))
            for i, (iid, *_) in enumerate(vals):
                self.elem.get('tree-selected').move(iid, '', i)

        return wrapper

    def __prep_elems(self):
        main = 'main'
        grid = {
            'row': {
                'default': {'weight': 1, 'minsize': 40},
                5: {'minsize': 80}    
                },
            'col': {
                'default': {'weight': 1, 'minsize': 50}
                }
            }
        menus = {
            'menu-main': {
                'main': True,
                'elements': [
                    ('menu', {'label': 'Plik', 'menu': 'menu-file'}),
                    ('command', {'label': 'Formatowanie', 'command': self.__show_format}),
                    ('command', {'label': 'Pomoc', 'command': self.__show_help})
                    ]
                },
            'menu-file': {
                'elements': [
                    ('command', {'label': 'Wybierz...', 'command': self.__select_file}),
                    ('command', {'label': 'Przeładuj', 'command': self.__reload}),
                    ('separator', None),
                    ('command', {'label': 'Wyjdź', 'command': self.root.destroy})
                    ]
                }
            }
        binds = [
            ('tree-all', ('<Return>', self.__add_by_btn)),
            ('tree-all', ('<Right>', self.__add_by_btn)),
            ('tree-all', ('<Double-Button-1>', self.__add_by_btn)),
            ('tree-selected', ('<Return>', self.__remove_by_btn)),
            ('tree-selected', ('<Left>', self.__remove_by_btn)),
            ('tree-selected', ('<Double-Button-1>', self.__remove_by_btn))
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
        self.vars.get('style').configure('TCheckbutton', background='white')
        self.vars.get('style').configure('TRadiobutton', background='white')
        self.elem.get('tree-selected').heading('project', text='Projekt')
        self.elem.get('tree-selected').heading('point', text='Punkt')
        self.elem.get('tree-selected').column('project', width=40, minwidth=40)
        self.elem.get('tree-selected').column('point', width=40, minwidth=40)
        self.elem.get('tree-all').configure(yscrollcommand=self.elem.get('scroll-all').set)
        self.elem.get('scroll-all').configure(command=self.elem.get('tree-all').yview)
        self.elem.get('radbtn-facture').invoke()
        self.elem.get('chkbtn-cash').invoke()
        self.elem.get('radbtn-auto').invoke()

    def __set_data(self):
        # check if file exists #
        if not isfile(self.vars.get('file')):
            self.__throw_error(5)
            return

        # try to read data #
        try:
            with open(self.vars.get('file'), 'rt', encoding='utf-8') as f:
                data = loads(f.read())
                for i, project in enumerate(data):
                    vals = [project.get('description')]
                    for x in project.get('dates'):
                        vals.extend((x.get('from'), x.get('to')))
                    self.elem.get('tree-all').insert('', 'end', i, text=project.get('name'), values=vals, open=False, tags=['catalogue'])
                    for point in project.get('points'):
                        self.elem.get('tree-all').insert(i, 'end', text=point.get('point'), values=[point.get('text')])

        except Exception as e:
            self.__throw_error(e)

    @__check
    def __clear_data(self):
        # clear data #
        self.elem.get('tree-all').delete(*self.elem.get('tree-all').get_children())
        self.elem.get('tree-selected').delete(*self.elem.get('tree-selected').get_children())

    def __get_state(self, state):
        return 'normal' if state else 'disabled'

    def __toggle(self, elem):
        match elem:
            case 'cash':
                state = self.__get_state(self.vars.get('var').get('cash').get())
                self.elem.get('entry-cash').config(state=state)
                self.elem.get('radbtn-auto').config(state=state)
                self.elem.get('radbtn-all').config(state=state)
            case 'addons':
                self.elem.get('entry-addons').config(state=self.__get_state(self.vars.get('var').get('addons').get()))
            case _:
                self.__throw_error(8)

    def __set_text(self):
        self.elem.get('txt-opening').delete('1.0', 'end')
        self.elem.get('txt-opening').insert('1.0', self.vars.get('var').get('{}-text'.format(self.vars.get('var').get('opening-mode').get())).get())

    def __get_date(self, date, dates, mode='raw'):
        # check which timeframe is correct #
        date, *dates = tuple(self.__str2date(x) for x in (date, *dates))
        if chosen := tuple((dates[i], dates[i + 1]) for i in range(0, len(dates), 2) if dates[i] <= date <= dates[i + 1]):
            match mode:
                case 'string':
                    return ' - '.join(tuple('{:02d}.{:02d}.{:04d}'.format(x.day, x.month, x.year) for x in chosen[0]))
                case 'int':
                    return tuple((x.day, x.month, x.year) for x in chosen[0])
                case 'raw' | _:
                    return chosen[0]
        else:
            return None

    def __set_path(self):
        # get path #
        if path := askdirectory(title='Wybierz folder', initialdir=self.vars.get('var').get('filepath').get()):
            self.vars.get('var').get('filepath').set(path)

    def __make_name(self):
        # get values #
        iid, name = self.elem.get('tree-selected').get_children()[0], []
        project, point, _, parent = self.elem.get('tree-selected').item(iid, 'values')

        # make new filename #
        if self.vars.get('var').get('opening-mode').get() == 'contract':
            name.append('u')
        name.append(sub('[^a-z0-9]+', '', unidecode(project).lower()))
        name.append(''.join(point.split('.')))
        if time := self.__get_date(self.vars.get('var').get('date').get(), self.elem.get('tree-all').item(parent, 'values')[1:]):
            start, end = time
            if start.year == end.year:
                if start.month == end.month:
                    name.append('{}-{}_{}_{:02d}'.format(start.day, end.day, start.month, start.year % 100))
                else:
                    name.append('{}-{}_{:02d}'.format(start.month, end.month, start.year % 100))
            else:
                name.append('{}_{:02d}-{}_{:02d}'.format(start.month, start.year % 100, end.month, end.year % 100))

        # set the filename #
        self.vars.get('var').get('filename').set('_'.join(name))

    @__sort
    @__check
    def __add_by_btn(self, event):
        iid = self.elem.get('tree-all').focus()

        # check wherever iid correct #
        if not iid:
            return

        # check wherever catalogue was selected #
        if self.elem.get('tree-all').tag_has('catalogue', iid):
            return

        # check wherever already not selected #
        if iid in tuple(self.elem.get('tree-selected').item(x, 'values')[2] for x in self.elem.get('tree-selected').get_children()):
            return

        # add element to table #
        parent = self.elem.get('tree-all').parent(iid)
        vals = (self.elem.get('tree-all').item(parent, 'text'), self.elem.get('tree-all').item(iid, 'text'), iid, parent)
        self.elem.get('tree-selected').insert('', 'end', values=vals)

    @__check
    def __remove_by_btn(self, event):
        iid = self.elem.get('tree-selected').focus()

        # check wherever iids correct #
        if not iid:
            return

        # remove selected elements #
        self.elem.get('tree-selected').delete(iid)

    @__sort
    @__check
    def __add(self):
        # get all iids, and leave out folders #
        iids = tuple((x, self.elem.get('tree-all').parent(x)) for x in self.elem.get('tree-all').selection() if not self.elem.get('tree-all').tag_has('catalogue', x))

        # check wherever iids correct #
        if not iids:
            self.__throw_error(0)
            return

        # check wherever already not selected #
        if any(x in tuple(y for y, _ in iids) for x in tuple(self.elem.get('tree-selected').item(x, 'values')[2] for x in self.elem.get('tree-selected').get_children())):
            self.__throw_error(2)
            return

        # add elements to table #
        for iid, parent in iids:
            vals = (self.elem.get('tree-all').item(parent, 'text'), self.elem.get('tree-all').item(iid, 'text'), iid, parent)
            self.elem.get('tree-selected').insert('', 'end', values=vals)

    @__sort
    @__check
    def __add_all(self):
        # clear table #
        self.__remove_all()

        # refill table #
        for parent in self.elem.get('tree-all').get_children():
            for iid in self.elem.get('tree-all').get_children(parent):
                vals = (self.elem.get('tree-all').item(parent, 'text'), self.elem.get('tree-all').item(iid, 'text'), iid, parent)
                self.elem.get('tree-selected').insert('', 'end', values=vals)

    @__check
    def __remove(self):
        iids = tuple(x for x in self.elem.get('tree-selected').selection())

        # check wherever iids correct #
        if not iids:
            self.__throw_error(0)
            return

        # remove selected elements #
        self.elem.get('tree-selected').delete(*iids)

    @__check
    def __remove_all(self):
        # remove all elements #
        self.elem.get('tree-selected').delete(*self.elem.get('tree-selected').get_children())

    def __print(self):
        name = self.vars.get('var').get('filename').get()

        # check if filename correct #
        if any(char in name for char in self.vars.get('badChars')):
            self.__throw_error(7)
            return

        path = '{}\\{}.docx'.format(self.vars.get('var').get('filepath').get(), name)

        # check wherever file exists #
        if isfile(path):
            if not askokcancel(title='Plik już istnieje', message='Czy chcesz kontynuować?'):
                self.__throw_error(3)
                return

        # get projects, and prepare text #
        txt = ''
        with StringIO('', newline='') as txt_file:
            if beg := self.elem.get('txt-opening').get('1.0', 'end-1c'):
                txt_file.write(beg + '<br><br>')

            items = self.elem.get('tree-selected').get_children()
            for project, parent in sorted({self.elem.get('tree-selected').item(iid, 'values')[0::3] for iid in items}, key=lambda x: x[0]):

                # perpare variables #
                desc, *dates = self.elem.get('tree-all').item(parent, 'values')
                time = self.__get_date(self.vars.get('var').get('date').get(), vals[1:], 'string')
                if not time:
                    self.__throw_error(4)
                    return

                # write text #
                txt_file.write(self.vars.get('var').get('title-text').get().replace('<t>', project) + '<br>')
                txt_file.write(desc.replace('<d>', time) + '<br>')

                # get data for point, and write them #
                collected_points = tuple(self.elem.get('tree-selected').item(iid, 'values')[1:3] for iid in items if self.elem.get('tree-selected').set(iid, 'project') == project)
                print_cash = (self.vars.get('var').get('cash-mode').get() == 'auto' and 1 < len(collected_points)) or self.vars.get('var').get('cash-mode').get() == 'all'
                for point, iid in collected_points:
                    if self.vars.get('var').get('cash').get() and print_cash:
                        txt_file.write(self.vars.get('var').get('cash-text').get())
                    txt_file.write(self.vars.get('var').get('point-text').get().replace('<p>', point).replace('<o>', *self.elem.get('tree-all').item(iid, 'values')))
                    if self.vars.get('var').get('addons').get():
                        txt_file.write(self.vars.get('var').get('addons-text').get())
                    txt_file.write('<br>')
                txt_file.write('<br>')

            txt = txt_file.getvalue().replace('<br>', '\n').strip()
        
        try:
            # create file #
            document = Document()

            # define style #
            style = document.styles['Normal']
            style.font.name = 'Calibri'
            style.font.size = Pt(10)

            # regex for tags #
            regex = '</?[{}]>'.format(''.join(self.vars.get('tags')))

            # regex matching #
            tags = findall(regex, txt)
            txt = split(regex, txt)

            # write data to docx #
            par = document.add_paragraph(txt.pop(0), style)
            for tag, part in zip(tags, txt):
                run = par.add_run(part)

                # set formatting #
                match tag.lower():
                    case '<b>':
                        run.font.bold = True
                    case '</b>':
                        run.font.bold = False
                    case '<i>':
                        run.font.italic = True
                    case '</i>':
                        run.font.italic = False
                    case '<u>':
                        run.font.underline = True
                    case '</u>':
                        run.font.underline = False
                    case '<s>':
                        run.font.strike = True
                    case '</s>':
                        run.font.strike = False

            # save file #
            document.save(path)

            showinfo(title='Zapisywanie', message='Sukces')

        except Exception as e:
            self.__throw_error(e)

    def __show_format(self):
        msg = \
            'Formatowanie opisu:\n' \
            '\n' \
            '<b> ... </b> - pogrubienie\n' \
            '<i> ... </i> - kursywa\n' \
            '<u> ... </u> - podkreślenie\n' \
            '<s> ... </s> - przekreślenie\n' \
            '\n' \
            '<t> - podpis projektu (niedostępny)\n' \
            '<p> - nazwa punktu (tylko w szablonie podpunktu)\n' \
            '<o> - treść punktu (tylko w szablonie podpunktu)\n' \
            '\n' \
            '<br> - nowa linia'
        showinfo(title='Formatowanie', message=msg)

    def __show_help(self):
        msg = \
            'Program do drukowania opisów.\n' \
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

    def __select_file(self):
        # get new path #
        path = askopenfilename(title='Wybierz plik', initialdir='\\'.join(self.vars.get('file').split('\\')[:-1]), filetypes=(('Plik JSON', '.json'), ), multiple=False).replace('/', '\\')
        if not path:
            return

        # check if extension correct #
        if '.json' not in path[-5:]:
            self.__throw_error(6)
            return

        # set new file #
        self.vars.update({'file': path})

    def __reload(self):
        # reload data #
        self.__clear_data()
        self.__set_data()

    def __throw_error(self, error):
        match error:
            case 0:
                msg = 'Nic nie zostało wybrane'
            # old #
            case 1:
                msg = 'Wybrany został projekt'
            case 2:
                msg = 'Punkt(y) już został wybrany'
            case 3:
                msg = 'Plik o tej nazwie już istnieje'
            case 4:
                msg = 'Data poza zasięgiem'
            case 5:
                msg = 'Brak podstawowego pliku danych'
            case 6:
                msg = 'Złe rozszerzenie pliku'
            case 7:
                msg = 'Niedozwolony znak w nazwie ( {} )'.format(' '.join(self.vars.get('badChars')))
            case 8:
                msg = 'Funkcja na niedozwolonym elemencie'
            case _:
                msg = error
        showerror(title='Błąd', message=msg)


    # other functions #
    def __str2date(self, d):
        return date(*reversed(tuple(int(x) for x in d.split('.'))))

    def __date2str(self, d):
        return '{}.{}.{}'.format(d.day, d.month, d.year)

    def __exint(self, n):
        try:
            return int(n)
        except Exception:
            roman, rest, n = {'I': 1, 'V': 5, 'X': 10, 'L': 50, 'C': 100, 'D': 500, 'M': 1000}, 0, n.upper()
            for i in range(len(n) - 1, -1, -1):
                num = roman.get(n[i], 0)
                rest = rest - num if 3 * num < rest else rest + num
            return rest


if __name__ == '__main__':
	TaxPrinter()
