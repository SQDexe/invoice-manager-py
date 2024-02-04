from os import getcwd
from os.path import isfile
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
from tkinter.ttk import Style, Frame, Entry, Button, Treeview, Checkbutton, Radiobutton
from tkcalendar import DateEntry
from tktooltip import ToolTip

class TaxPrinter:
    def __init__(self):
        # root declaration #
        self.root = Tk()

        # declearing variables, and elements #
        self.vars = {
            'title': 'Tax Printer',
            'minWidth': 400,
            'minHeight': 450,
            'maxSize': 650,
            'icon': 'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQBAMAAADt3eJSAAAAIGNIUk0AAHomAACAhAAA+gAAAIDoAAB1MAAA6mAAADqYAAAXcJy6UTwAAAAeUExURQAAAP///9vb2wAAAI+Pj/+FhYv/hf//a3BwcPDw8ChUvUYAAAABdFJOUwBA5thmAAAAAWJLR0QB/wIt3gAAAAd0SU1FB+cLCBAqL61tymUAAAA9SURBVAjXY2BAAEFBIQhD2NhIASaiBBUxNoSKCApiF3EBAyDDNQVIlyCLuBgDAZihBAQgRsdMIOhAMRACAAYpDjSL+1GnAAAAJXRFWHRkYXRlOmNyZWF0ZQAyMDIzLTExLTA4VDE2OjQyOjQ2KzAwOjAwmi5JNwAAACV0RVh0ZGF0ZTptb2RpZnkAMjAyMy0xMS0wOFQxNjo0Mjo0NiswMDowMOtz8YsAAAAodEVYdGRhdGU6dGltZXN0YW1wADIwMjMtMTEtMDhUMTY6NDI6NDcrMDA6MDAaEdvgAAAAAElFTkSuQmCC',
            'file': '{}\\{}'.format(getcwd(), 'data.json'),
            'pad': 5,
            'var-date': StrVar(),
            'var-filename': StrVar(value='plik'),
                    'var-filepath': StrVar(value=getcwd()),
            'var-point-text': StrVar(value='<b>Punkt <p></b> „<o>”'),
            'var-cash': BoolVar(),
            'var-addons': BoolVar(),
            'var-cash-text': StrVar(value='Suma <b>............... ,-</b>, '),
            'var-addons-text': StrVar(value=' - <br>'),
            'var-opening-text': StrVar(),
            'style': Style(),
            'tags': ('b', 'i', 'u', 's'),
	    'badChars': '\\/:*?"<>|'
            }
        self.elem = {
            'tree-all': {
                'type': Treeview,
                'args': {'selectmode': 'extended', 'show': 'tree'},
                'grid': {'row': 0, 'column': 0, 'rowspan': 5, 'columnspan': 2},
                'sticky': 'NWES'
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
            'entry-point': {
                'type': Entry,
                'args': {'textvariable': self.vars.get('var-point-text')},
                'grid': {'row': 5, 'column': 0, 'columnspan': 2},
                'tooltip': 'Tekst podpunktu',
                'sticky': 'NWES'
                },
            'chkbtn-cash': {
                'type': Checkbutton,
                'args': {'text': 'Dodaj pole kwoty', 'variable': self.vars.get('var-cash'), 'command': lambda: self.__toggle('cash')},
                'grid': {'row': 6, 'column': 0},
                'sticky': 'W'
                },
            'entry-cash': {
                'type': Entry,
                'args': {'textvariable': self.vars.get('var-cash-text'), 'state': 'disabled'},
                'grid': {'row': 6, 'column': 1},
                'tooltip': 'Tekst kwoty',
                'sticky': 'NWES'
                },
            'chkbtn-addons': {
                'type': Checkbutton,
                'args': {'text': 'Dodaj myślniki', 'variable': self.vars.get('var-addons'), 'command': lambda: self.__toggle('addons')},
                'grid': {'row': 7, 'column': 0},
                'sticky': 'W'
                },
            'entry-addons': {
                'type': Entry,
                'args': {'textvariable': self.vars.get('var-addons-text'), 'state': 'disabled'},
                'grid': {'row': 7, 'column': 1},
                'tooltip': 'Tekst notatki',
                'sticky': 'NWES'
                },
            'radbtn-first': {
                'type': Radiobutton,
                'args': {'text': 'Faktura', 'variable': self.vars.get('var-opening-text'), 'command': self.__set_text, 'value': 'Akapit przykładowy 1.\nTutaj wpisać tekst na, który ma pojawić się na początku dokumentu'},
                'grid': {'row': 8, 'column': 0},
                'sticky': 'W'
                },
            'radbtn-second': {
                'type': Radiobutton,
                'args': {'text': 'Umowa', 'variable': self.vars.get('var-opening-text'), 'command': self.__set_text, 'value': 'Akapit przykładowy 2.\nTutaj wpisać tekst na, który ma pojawić się na początku dokumentu'},
                'grid': {'row': 8, 'column': 1},
                'sticky': 'W'
                },
            'cal-select': {
                'type': DateEntry,
                'args': {'date_pattern': 'd.M.y', 'locale': 'pl_PL', 'textvariable': self.vars.get('var-date')},
                'grid': {'row': 6, 'column': 3},
                'tooltip': 'Data wystawienia opisu'
                },
            'btn-name': {
                'type': Button,
                'args': {'text': '\u270e', 'state': 'disabled', 'command': self.__make_name},
                'grid': {'row': 6, 'column': 4},
                'tooltip': 'Wygeneruj nazwę do jednego podpunktu'
                },
            'entry-filename': {
                'type': Entry,
                'args': {'textvariable': self.vars.get('var-filename')},
                'grid': {'row': 7, 'column': 3, 'columnspan': 2},
                'tooltip': 'Nazwa pliku',
                'sticky': 'NWES'
                },
            'btn-filepath': {
                'type': Button,
                'args': {'text': '...', 'command': self.__set_path},
                'grid': {'row': 8, 'column': 3},
                'tooltip': 'Wybierz ścieżkę docelową',
                        'sticky': 'W'
                },
            'btn-print': {
                'type': Button,
                'args': {'text': '\ud83d\uddb6', 'state': 'disabled', 'command': self.__print},
                'grid': {'row': 8, 'column': 4},
                'tooltip': 'Drukuj opis'
                },
            'txt-opening': {
                'type': Text,
                'args': {'wrap': 'word'},
                'grid': {'row': 9, 'column': 0, 'columnspan': 5},
                'sticky': 'NWES',
                'borderfull': {'highlightthickness': 1, 'highlightbackground': 'gray'}
                }
            }
        self.tooltips = []

        # window's settings #
        width, height = self.vars.get('minWidth'), self.vars.get('minHeight')
        self.root.geometry('{}x{}+{}+{}'.format(
            width,
            height,
            int((self.root.winfo_screenwidth() - width) / 2),
            int((self.root.winfo_screenheight() - height) / 2)
            ))
        self.root.title(self.vars.get('title'))
        self.root.minsize(width, height)
        self.root.maxsize(self.vars.get('maxSize'), self.vars.get('maxSize'))
        self.root.resizable(True, True)
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
            vals.sort(key=lambda x: (x[2], list(map(self.__exint, x[1].split('.')))))
            for i, (iid, *_) in enumerate(vals):
                self.elem.get('tree-selected').move(iid, '', i)

        return wrapper

    def __prep_elems(self):
        main = 'main'
	grid = {
	    'col': {
		0: {'minsize': 130}    
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
        for key, data in [(x, y) for x, y in self.elem.items() if x != main]:
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

        # create menus #
        self.elem.update({'menu-main': Menu(self.root, tearoff=0)})
        self.elem.update({'menu-file': Menu(self.elem.get('menu-main'), tearoff=0)})
        self.root.config(menu=self.elem.get('menu-main'))
        self.elem.get('menu-main').add_cascade(label='Plik', menu=self.elem.get('menu-file'))
        self.elem.get('menu-main').add_command(label='Formatowanie', command=self.__show_format)
        self.elem.get('menu-main').add_command(label='Pomoc', command=self.__show_help)
        self.elem.get('menu-file').add_command(label='Wybierz...', command=self.__select_file)
        self.elem.get('menu-file').add_command(label='Przeładuj', command=self.__reload)
        self.elem.get('menu-file').add_separator()
        self.elem.get('menu-file').add_command(label='Wyjdź', command=self.root.destroy)

        # grid settings #
        cols, rows = self.elem.get(main).grid_size()
        for i in range(rows):
            self.elem.get(main).rowconfigure(i, **{'weight': 1, 'minsize': 40} | grid.get('row', {}).get(i, {}))
        for i in range(cols):
            self.elem.get(main).columnconfigure(i, **{'weight': 1, 'minsize': 50} | grid.get('col', {}).get(i, {}))

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
        self.elem.get('radbtn-first').invoke()

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

    def __toggle(self, elem):
        self.elem.get('entry-{}'.format(elem)).config(state='normal' if self.vars.get('var-{}'.format(elem)).get() else 'disabled')

    def __set_text(self):
        self.elem.get('txt-opening').delete('1.0', 'end')
        self.elem.get('txt-opening').insert('1.0', self.vars.get('var-opening-text').get())

    def __get_date(self, date, dates, raw=False):
        # check which timeframe is correct #
        date, *dates = list(map(self.__str2date, [date, *dates]))
        if chosen := [[dates[i], dates[i + 1]] for i in range(0, len(dates), 2) if dates[i] <= date <= dates[i + 1]]:
            return ' - '.join(list(map(lambda x: '{:02d}.{:02d}.{:04d}'.format(x.day, x.month, x.year), chosen[0]))) if not raw else chosen[0]
        else:
            return None

    def __set_path(self):
        # get path #
        if path := askdirectory(title='Wybierz folder', initialdir=self.vars.get('var-filepath').get()):
            self.vars.get('var-filepath').set(path)

    def __make_name(self):
        # get values #
        iid, name = self.elem.get('tree-selected').get_children()[0], []
        project, point, _, parent = self.elem.get('tree-selected').item(iid, 'values')

        # make new filename #
        name.append(sub('[^a-z0-9]+', '', unidecode(project).lower()))
        name.append(''.join(point.split('.')))
        if time := self.__get_date(self.vars.get('var-date').get(), self.elem.get('tree-all').item(parent, 'values')[1:], True):
            start, end = time
	    datestamp = ''
	    if start.year == end.year:
		if start.month == end.month:
		    datestamp = '{}-{}_{}_{:02d}'.format(start.day, end.day, start.month, start.year % 100)
		else:
		    datestamp = '{}-{}_{:02d}'.format(start.month, end.month, start.year % 100)
	    else:
	    	datestamp = '{}_{:02d}-{}_{:02d}'.format(start.month, start.year % 100, end.month, end.year % 100)

            name.append(datestamp)

        # set the filename #
        self.vars.get('var-filename').set('_'.join(name))

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
        if iid in [self.elem.get('tree-selected').item(x, 'values')[2] for x in self.elem.get('tree-selected').get_children()]:
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
        iids = [(x, self.elem.get('tree-all').parent(x)) for x in self.elem.get('tree-all').selection() if not self.elem.get('tree-all').tag_has('catalogue', x)]

        # check wherever iids correct #
        if not iids:
            self.__throw_error(0)
            return

        # check wherever already not selected #
        if any(x in [y for y, _ in iids] for x in [self.elem.get('tree-selected').item(x, 'values')[2] for x in self.elem.get('tree-selected').get_children()]):
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
        iids = [x for x in self.elem.get('tree-selected').selection()]

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
        name = self.vars.get('var-filename').get()

        # check if filename correct #
        if any(char in name for char in self.vars.get('badChars')):
            self.__throw_error(7)
            return

        path = '{}\\{}.docx'.format(self.vars.get('var-filepath').get(), name)

        # check wherever file exists #
        if isfile(path):
            if not askokcancel(title='Plik już istnieje', message='Czy chcesz kontynuować?'):
                self.__throw_error(3)
                return

        # get projects, and prepare text #
        txt = ''
        if beg := self.elem.get('txt-opening').get('1.0', 'end-1c'):
            txt = beg + '\n'
        items = self.elem.get('tree-selected').get_children()
        for project, parent in sorted({self.elem.get('tree-selected').item(iid, 'values')[0::3] for iid in items}, key=lambda x: x[0]):

            # perpare variables #
            vals = self.elem.get('tree-all').item(parent, 'values')
            time = self.__get_date(self.vars.get('var-date').get(), vals[1:])
            if not time:
                self.__throw_error(4)
                return
            desc = vals[0].replace('<d>', time)

            # write text #
            txt += '<b>{}</b>\n'.format(project)
            txt += desc + '\n'

            # get data for point, and write them #
            for point, iid in [self.elem.get('tree-selected').item(iid, 'values')[1:3] for iid in items if self.elem.get('tree-selected').set(iid, 'project') == project]:
                if self.vars.get('var-cash').get():
                    txt += self.vars.get('var-cash-text').get()
                desc = self.vars.get('var-point-text').get()
                desc = desc.replace('<p>', point)
                desc = desc.replace('<o>', *self.elem.get('tree-all').item(iid, 'values'))
                txt += desc
                if self.vars.get('var-addons').get():
                    txt += self.vars.get('var-addons-text').get()
                txt += '\n'
            txt = txt.replace('<br>', '\n')
            txt += '\n'

        txt = txt.strip()
        
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
            case _:
                msg = error
        showerror(title='Błąd', message=msg)


    # other functions #
    def __str2date(self, d):
        return date(*reversed(list(map(int, d.split('.')))))

    def __date2str(self, d):
        return '{}.{}.{}'.format(d.day, d.month, d.year)

    def __exint(self, n):
        try:
            return int(n)
        except:
            roman, rest, n = {'I': 1, 'V': 5, 'X': 10, 'L': 50, 'C': 100, 'D': 500, 'M': 1000}, 0, n.upper()
            for i in range(len(n) - 1, -1, -1):
                num = roman.get(n[i], 0)
                rest = rest - num if 3 * num < rest else rest + num
            return rest


if __name__ == '__main__':
	TaxPrinter()
