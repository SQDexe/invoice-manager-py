from utilities import WindowApp

from os import getcwd
from os.path import isfile
from io import StringIO
from json import loads
from re import findall, split, sub
from functools import wraps

from unidecode import unidecode

from docx import Document
from docx.shared import Pt

from tkinter import StringVar as StrVar, BooleanVar as BoolVar, Text
from tkinter.ttk import Entry, Button, Treeview, Scrollbar, Checkbutton, Radiobutton
from tkinter.messagebox import showinfo, askokcancel
from tkinter.filedialog import askopenfilename, askdirectory
from tkcalendar import DateEntry

# from accessify import private

class TaxPrinter(WindowApp):
	# decorators #
	def check(f):
	    @wraps(f)
	    def wrapper(self, *args, **kwargs):
	        f(self)

	        childs = bool(self.elem['tree-selected'].get_children())
	        self.elem['btn-print'].config(state=self.get_state(childs))
	        self.elem['btn-name'].config(state=self.get_state(childs))

	    return wrapper

	def sort(f):
	    @wraps(f)
	    def wrapper(self, *args, **kwargs):
	        f(self)

	        # sort by point, then by project #
	        vals = [(iid, self.elem['tree-selected'].set(iid, 'point'), self.elem['tree-selected'].set(iid, 'project')) for iid in self.elem['tree-selected'].get_children()]
	        vals.sort(key=lambda x: (x[2], tuple(self.roman2int(y) for y in x[1].split('.'))))
	        for i, (iid, *_) in enumerate(vals):
	            self.elem['tree-selected'].move(iid, '', i)

	    return wrapper

	def set_data(self):
	    # check if file exists #
	    if not isfile(self.vars['file']):
	        self.throw_error(1)
	        return

	    # try to read data #
	    try:
	        with open(self.vars['file'], 'rt', encoding='utf-8') as f:
	            data = loads(f.read())
	            for i, project in enumerate(data):
	                vals = [project['description']]
	                for x in project['dates']:
	                    vals.extend((x['from'], x['to']))
	                self.elem['tree-all'].insert('', 'end', i, text=project['name'], values=vals, open=False, tags=['catalogue'])
	                for point in project['points']:
	                    self.elem['tree-all'].insert(i, 'end', text=point['point'], values=[point['text']])

	    except Exception as e:
	        self.throw_error(e)

	@check
	def clear_data(self):
	    # clear data #
	    self.elem['tree-all'].delete(*self.elem['tree-all'].get_children())
	    self.elem['tree-selected'].delete(*self.elem['tree-selected'].get_children())

	def toggle(self, elem):
	    states = (self.vars['var']['cash'].get(), self.vars['var']['addons'].get())
	    cash_state, addons_state = tuple(self.get_state(x) for x in states)
	    self.elem['entry-cash'].config(state=cash_state)
	    self.elem['radbtn-auto'].config(state=cash_state)
	    self.elem['radbtn-all'].config(state=cash_state)
	    self.elem['entry-addons'].config(state=addons_state)

	def set_text(self):
	    self.elem['txt-opening'].delete('1.0', 'end')
	    self.elem['txt-opening'].insert('1.0', self.vars['var']['{}-text'.format(self.vars['var']['opening-mode'].get())].get())

	def get_date(self, date, dates, mode='raw'):
	    # check which timeframe is correct #
	    date, *dates = tuple(self.str2date(x) for x in (date, *dates))
	    if chosen := tuple((dates[i], dates[i + 1]) for i in range(0, len(dates), 2) if dates[i] <= date <= dates[i + 1]):
		chosen_data = chosen[0]
	        match mode:
	            case 'string':
	                return ' - '.join(('{:02d}.{:02d}.{:04d}'.format(x.day, x.month, x.year) for x in chosen_data))
	            case 'int':
	                return tuple((x.day, x.month, x.year) for x in chosen_data)
	            case 'raw' | _:
	                return chosen_data
	    else:
	        return None

	def set_path(self):
	    # get path #
	    if path := askdirectory(title='Wybierz folder', initialdir=self.vars['var']['filepath'].get()):
	        self.vars['var']['filepath'].set(path)

	def make_name(self):
	    # get values #
	    iids, name = self.elem['tree-selected'].get_children(), []

	    # make new filename #
	    if self.vars['var']['opening-mode'].get() == 'contract':
	        name.append('u')

	    if len(iids) != 1:
	        name.extend(sorted(set(sub('[^a-z]+', '', unidecode(self.elem['tree-all'].item(iid, 'values')[0]).lower()) for iid in iids)))
	    else:
		project, point, _, parent = self.elem['tree-selected'].item(iid, 'values')
		name.append(sub('[^a-z0-9]+', '', unidecode(project).lower()))
		name.append(''.join(point.split('.')))

	    	if time := self.get_date(self.vars['var']['date'].get(), self.elem['tree-all'].item(*iids, 'values')[1:]):
	            start, end = time
		    name.append(( \
			    '{}-{}_{}_{:02d}'.format(start.day, end.day, start.month, start.year % 100) \
			    if start.month == end.month else \
			    '{}-{}_{:02d}'.format(start.month, end.month, start.year % 100)) \
			if start.year == end.year else \
		    	'{}_{:02d}-{}_{:02d}'.format(start.month, start.year % 100, end.month, end.year % 100))

	    # set the filename #
	    self.vars['var']['filename'].set('_'.join(name))

	@sort
	@check
	def add(self):
	    # get all iids and leave out folders #
	    iids = tuple((x, self.elem['tree-all'].parent(x)) for x in self.elem['tree-all'].selection() if not self.elem['tree-all'].tag_has('catalogue', x))

	    # check wherever iids correct #
	    if not iids:
	        self.throw_error(201)
	        return

	    # check wherever already not selected #
	    if any(x in tuple(y for y, _ in iids) for x in tuple(self.elem['tree-selected'].item(x, 'values')[2] for x in self.elem['tree-selected'].get_children())):
	        self.throw_error(602)
	        return

	    # add elements to table #
	    for iid, parent in iids:
	        vals = (self.elem['tree-all'].item(parent, 'text'), self.elem['tree-all'].item(iid, 'text'), iid, parent)
	        self.elem['tree-selected'].insert('', 'end', values=vals)

	@sort
	@check
	def add_all(self):
	    # clear table #
	    self.remove_all()

	    # refill table #
	    for parent in self.elem['tree-all'].get_children():
	        for iid in self.elem['tree-all'].get_children(parent):
	            vals = (self.elem['tree-all'].item(parent, 'text'), self.elem['tree-all'].item(iid, 'text'), iid, parent)
	            self.elem['tree-selected'].insert('', 'end', values=vals)

	@sort
	@check
	def add_by_btn(self):
	    iid = self.elem['tree-all'].focus()

	    # check wherever iid correct #
	    if not iid:
	        return

	    # check wherever catalogue was selected #
	    if self.elem['tree-all'].tag_has('catalogue', iid):
	        return

	    # check wherever already not selected #
	    if iid in tuple(self.elem['tree-selected'].item(x, 'values')[2] for x in self.elem['tree-selected'].get_children()):
	        return

	    # add element to table #
	    parent = self.elem['tree-all'].parent(iid)
	    vals = (self.elem['tree-all'].item(parent, 'text'), self.elem['tree-all'].item(iid, 'text'), iid, parent)
	    self.elem['tree-selected'].insert('', 'end', values=vals)

	@check
	def remove(self):
	    iids = tuple(x for x in self.elem['tree-selected'].selection())

	    # check wherever iids correct #
	    if not iids:
	        self.throw_error(201)
	        return

	    # remove selected elements #
	    self.elem['tree-selected'].delete(*iids)

	@check
	def remove_all(self):
	    # remove all elements #
	    self.elem['tree-selected'].delete(*self.elem['tree-selected'].get_children())

	@check
	def remove_by_btn(self):
	    iid = self.elem['tree-selected'].focus()

	    # check wherever iids correct #
	    if not iid:
	        return

	    # remove selected elements #
	    self.elem['tree-selected'].delete(iid)

	def print(self):
	    name = self.vars['var']['filename'].get()

	    # check if filename not empty #
	    if not name:
	        self.throw_error(3)
	        return

	    # check if filename correct #
	    if any(char in name for char in self.vars['bad-chars']):
	        self.throw_error(2)
	        return

	    path = '{}\\{}.docx'.format(self.vars['var']['filepath'].get(), name)

	    # check wherever file exists #
	    if isfile(path):
	        if not askokcancel(title='Plik już istnieje', message='Czy chcesz kontynuować?'):
	            self.throw_error(5)
	            return

	    # get projects, and prepare text #
	    txt = ''
	    with StringIO('', newline='') as txt_file:
	        if beg := self.elem['txt-opening'].get('1.0', 'end-1c'):
	            txt_file.write(beg + '<br><br>')

	        items = self.elem['tree-selected'].get_children()
	        for project, parent in sorted({self.elem['tree-selected'].item(iid, 'values')[0::3] for iid in items}, key=lambda x: x[0]):

	            # perpare variables #
	            desc, *dates = self.elem['tree-all'].item(parent, 'values')
	            time = self.get_date(self.vars['var']['date'].get(), dates, 'string')
	            if not time:
	                self.throw_error(603)
	                return

	            # write text #
	            txt_file.write(self.vars['var']['title-text'].get().replace('<t>', project) + '<br>')
	            txt_file.write(desc.replace('<d>', time) + '<br>')

	            # get data for point, and write them #
	            collected_points = tuple(self.elem['tree-selected'].item(iid, 'values')[1:3] for iid in items if self.elem['tree-selected'].set(iid, 'project') == project)
	            print_cash = (self.vars['var']['cash-mode'].get() == 'auto' and 1 < len(collected_points)) or self.vars['var']['cash-mode'].get() == 'all'
	            for point, iid in collected_points:
	                if self.vars['var']['cash'].get() and print_cash:
	                    txt_file.write(self.vars['var']['cash-text'].get())
	                txt_file.write(self.vars['var']['point-text'].get().replace('<p>', point).replace('<o>', *self.elem['tree-all'].item(iid, 'values')))
	                if self.vars['var']['addons'].get():
	                    txt_file.write(self.vars['var']['addons-text'].get())
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
	        regex = '</?[{}]>'.format(''.join(self.vars['tags']))

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
	        self.throw_error(e)

	def show_format(self):
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
	        '<br> - nowa linia\n' \
	        '\n' \
	        '{} - niedozwolone znaki nazwy pliku' \
	        .format(' '.join(self.vars['bad-chars']))
	    showinfo(title='Formatowanie', message=msg)

	def select_file(self):
	    # get new path #
	    path = askopenfilename(title='Wybierz plik', initialdir='\\'.join(self.vars['file'].split('\\')[:-1]), filetypes=(('Plik JSON', '.json'), ), multiple=False).replace('/', '\\')
	    if not path:
	        return

	    # check if extension correct #
	    if '.json' not in path[-5:]:
	        self.throw_error(4)
	        return

	    # set new file #
	    self.vars['file'] = path

	def reload(self):
	    # reload data #
	    self.clear_data()
	    self.set_data()

	# overridden #
	def pre(self):
	    # important #
		self.vars.update({
	        'title': 'Tax Printer',
	        'size': {
	            'min': (400, 450),
	            'max': (650, 650)
	            }
	        })

		# add, and set variables #
		self.vars.update({
	        'file': '{}\\{}'.format(getcwd(), 'data.json'),
	        'tags': ('b', 'i', 'u', 's'),
	        'bad-chars': r'\/:*?"<>|'
	        })
		self.vars['var'].update({
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
            })
		self.elem.update({
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
	            'args': {'text': '\u25b6', 'command': self.add},
	            'grid': {'row': 0, 'column': 2},
	            'tooltip': 'Dodaj element(y)'
	            },
	        'btn-add-all': {
	            'type': Button,
	            'args': {'text': '\u25b6\u25b6', 'command': self.add_all},
	            'grid': {'row': 1, 'column': 2},
	            'tooltip': 'Dodaj wszystkie elementy'
	            },
	        'btn-remove': {
	            'type': Button,
	            'args': {'text': '\u25c0', 'command': self.remove},
	            'grid': {'row': 3, 'column': 2},
	            'tooltip': 'Usuń element(y)'
	            },
	        'btn-remove-all': {
	            'type': Button,
	            'args': {'text': '\u25c0\u25c0', 'command': self.remove_all},
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
	            'args': {'text': 'Faktura', 'variable': self.vars['var']['opening-mode'], 'command': self.set_text, 'value': 'facture'},
	            'grid': {'row': 6, 'column': 3},
	            'sticky': 'W'
	            },
	        'radbtn-contract': {
	            'type': Radiobutton,
	            'args': {'text': 'Umowa', 'variable': self.vars['var']['opening-mode'], 'command': self.set_text, 'value': 'contract'},
	            'grid': {'row': 6, 'column': 4},
	            'sticky': 'W'
	            },
	        'entry-point': {
	            'type': Entry,
	            'args': {'textvariable': self.vars['var']['point-text']},
	            'grid': {'row': 6, 'column': 0, 'columnspan': 2},
	            'tooltip': 'Szablon podpunktu',
	            'sticky': 'NWES'
	            },
	        'chkbtn-cash': {
	            'type': Checkbutton,
	            'args': {'text': 'Dodaj pole kwoty', 'variable': self.vars['var']['cash'], 'command': self.toggle},
	            'grid': {'row': 7, 'column': 0},
	            'sticky': 'W'
	            },
	        'entry-cash': {
	            'type': Entry,
	            'args': {'textvariable': self.vars['var']['cash-text'], 'state': 'disabled'},
	            'grid': {'row': 7, 'column': 1},
	            'tooltip': 'Tekst kwoty',
	            'sticky': 'NWES'
	            },
	        'radbtn-auto': {
	            'type': Radiobutton,
	            'args': {'text': 'Auto', 'variable': self.vars['var']['cash-mode'], 'value': 'auto'},
	            'grid': {'row': 8, 'column': 0},
	            'sticky': 'W'
	            },
	        'radbtn-all': {
	            'type': Radiobutton,
	            'args': {'text': 'Wszystko', 'variable': self.vars['var']['cash-mode'], 'value': 'all'},
	            'grid': {'row': 8, 'column': 1},
	            'sticky': 'W'
	            },
	        'chkbtn-addons': {
	            'type': Checkbutton,
	            'args': {'text': 'Dodaj myślniki', 'variable': self.vars['var']['addons'], 'command': self.toggle},
	            'grid': {'row': 9, 'column': 0},
	            'sticky': 'W'
	            },
	        'entry-addons': {
	            'type': Entry,
	            'args': {'textvariable': self.vars['var']['addons-text'], 'state': 'disabled'},
	            'grid': {'row': 9, 'column': 1},
	            'tooltip': 'Tekst notatki',
	            'sticky': 'NWES'
	            },
	        'cal-select': {
	            'type': DateEntry,
	            'args': {'date_pattern': 'd.M.y', 'locale': 'pl_PL', 'textvariable': self.vars['var']['date']},
	            'grid': {'row': 7, 'column': 3},
	            'tooltip': 'Data wystawienia opisu'
	            },
	        'btn-name': {
	            'type': Button,
	            'args': {'text': '\u270e', 'state': 'disabled', 'command': self.make_name},
	            'grid': {'row': 7, 'column': 4},
	            'tooltip': 'Wygeneruj nazwę do jednego podpunktu'
	            },
	        'entry-filename': {
	            'type': Entry,
	            'args': {'textvariable': self.vars['var']['filename']},
	            'grid': {'row': 8, 'column': 3, 'columnspan': 2},
	            'tooltip': 'Nazwa pliku',
	            'sticky': 'NWES'
	            },
	        'btn-filepath': {
	            'type': Button,
	            'args': {'text': '...', 'command': self.set_path},
	            'grid': {'row': 9, 'column': 3},
	            'tooltip': 'Wybierz ścieżkę docelową',
	            'sticky': 'W'
	            },
	        'btn-print': {
	            'type': Button,
	            'args': {'text': '\ud83d\uddb6', 'state': 'disabled', 'command': self.print},
	            'grid': {'row': 9, 'column': 4},
	            'tooltip': 'Drukuj opis'
	            }
	        })
		self.grid['row'].update({
	        5: {'minsize': 80}
	        })
	    # self.grid['col'].update()
		self.menus.update({
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
	        })
		self.binds.extend([
	        ('tree-all', ('<Return>', self.add_by_btn)),
	        ('tree-all', ('<Right>', self.add_by_btn)),
	        ('tree-all', ('<Double-Button-1>', self.add_by_btn)),
	        ('tree-selected', ('<Return>', self.remove_by_btn)),
	        ('tree-selected', ('<Left>', self.remove_by_btn)),
	        ('tree-selected', ('<Double-Button-1>', self.remove_by_btn))
	        ])

	def post(self):
	    # styles, other settings, and actions #
	    self.vars['style'].configure('TFrame', background='white')
	    self.vars['style'].configure('TCheckbutton', background='white')
	    self.vars['style'].configure('TRadiobutton', background='white')
	    self.elem['tree-selected'].heading('project', text='Projekt')
	    self.elem['tree-selected'].heading('point', text='Punkt')
	    self.elem['tree-selected'].column('project', width=40, minwidth=40)
	    self.elem['tree-selected'].column('point', width=40, minwidth=40)
	    self.elem['tree-all'].configure(yscrollcommand=self.elem['scroll-all'].set)
	    self.elem['scroll-all'].configure(command=self.elem['tree-all'].yview)
	    self.elem['radbtn-facture'].invoke()
	    self.elem['chkbtn-cash'].invoke()
	    self.elem['radbtn-auto'].invoke()

	    self.set_data()

	def close(self):
	    self.root.destroy()


if __name__ == '__main__':
    TaxPrinter()
