from os import getcwd, remove
from os.path import isfile
from json import loads
from datetime import date
from time import time
from re import findall, split, sub

from base64 import b64decode
from docx import Document
from docx.shared import Pt

from tkinter import Tk, StringVar as StrVar, BooleanVar as BoolVar, Frame, Text
from tkinter.messagebox import showerror, showinfo, askokcancel
from tkinter.ttk import Style, Entry, Button, Treeview, Checkbutton, Label
from tkcalendar import DateEntry
from tktooltip import ToolTip

class TaxPrinter:
	def __init__(self):
		# root declaration #
		self.root = Tk()

		# declearing variables, and elements #
		self.vars = {
			'title': 'Tax Printer',
			'minWidth': 450,
			'minHeight': 450,
			'maxSize': 650,
			'workDir': getcwd(),
			'workFile': 'data.json',
			'pad': 5,
			'var-date': StrVar(),
			'var-filename': StrVar(value='plik'),
			'var-point-text': StrVar(value='<b>Punkt <p></b> „<o>”'),
			'var-cash': BoolVar(),
			'var-addons': BoolVar(),
			'var-cash-text': StrVar(value='Suma <b>............... ,-</b>, '),
			'var-addons-text': StrVar(value=' - <br>'),
			'opening-text': 'Akapit przykładowy\nTutaj wpisać tekst na, który ma pojawić się na początku dokumentu',
			'style': Style(),
			'tags': ('b', 'i', 'u', 's')
			}
		self.elem = {
			'tree-all': {
				'type': Treeview,
				'args': {'selectmode': 'extended', 'show': 'tree'},
				'grid': {'row': 0, 'column': 0, 'rowspan': 5, 'columnspan': 2},
				'sticky': {'sticky': 'NWES'}
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
				'sticky': {'sticky': 'NWES'}
				},
			'entry-point': {
				'type': Entry,
				'args': {'textvariable': self.vars.get('var-point-text')},
				'grid': {'row': 5, 'column': 0, 'columnspan': 2},
				'tooltip': 'Tekst podpunktu',
				'sticky': {'sticky': 'NWES'}
				},
			'chkbtn-cash': {
				'type': Checkbutton,
				'args': {'text': 'Dodaj pole kwoty', 'variable': self.vars.get('var-cash'), 'command': lambda: self.__toggle('cash')},
				'grid': {'row': 6, 'column': 0},
				'sticky': {'sticky': 'W'}
				},
			'entry-cash': {
				'type': Entry,
				'args': {'textvariable': self.vars.get('var-cash-text'), 'state': 'disabled'},
				'grid': {'row': 6, 'column': 1},
				'tooltip': 'Tekst kwoty',
				'sticky': {'sticky': 'NWES'}
				},
			'chkbtn-addons': {
				'type': Checkbutton,
				'args': {'text': 'Dodaj myślniki', 'variable': self.vars.get('var-addons'), 'command': lambda: self.__toggle('addons')},
				'grid': {'row': 7, 'column': 0},
				'sticky': {'sticky': 'W'}
				},
			'entry-addons': {
				'type': Entry,
				'args': {'textvariable': self.vars.get('var-addons-text'), 'state': 'disabled'},
				'grid': {'row': 7, 'column': 1},
				'tooltip': 'Tekst notatki',
				'sticky': {'sticky': 'NWES'}
				},
			'label-tip': {
				'type': Label,
				'args': {'text': '?', 'background': 'white'},
				'grid': {'row': 5, 'column': 2},
				'tooltip': 'Formatowanie opisów:\n<b> ... </b> - pogrubienie\n<i> ... </i> - kursywa\n<u> ... </u> - podkreślenie\n<s> ... </s> - przekreślenie\n<br> - nowa linia\n<p> - nazwa punktu\n<o> - treść punktu'
				},
			'btn-name': {
				'type': Button,
				'args': {'text': '\u270e', 'state': 'disabled', 'command': self.__make_name},
				'grid': {'row': 5, 'column': 4},
				'tooltip': 'Wygeneruj nazwę do jednego podpunktu'
				},
			'entry-filename': {
				'type': Entry,
				'args': {'textvariable': self.vars.get('var-filename')},
				'grid': {'row': 6, 'column': 3, 'columnspan': 2},
				'tooltip': 'Nazwa pliku',
				'sticky': {'sticky': 'NWES'}
				},
			'cal-select': {
				'type': DateEntry,
				'args': {'date_pattern': 'd.M.y', 'locale': 'pl_PL', 'textvariable': self.vars.get('var-date')},
				'grid': {'row': 7, 'column': 3},
				'tooltip': 'Data wystawienia opisu'
				},
			'btn-print': {
				'type': Button,
				'args': {'text': '\ud83d\uddb6', 'state': 'disabled', 'command': self.__print},
				'grid': {'row': 7, 'column': 4},
				'tooltip': 'Drukuj opis'
				},
			'text-opening': {
				'type': Text,
				'args': {},
				'grid': {'row': 8, 'column': 0, 'columnspan': 5},
				'sticky': {'sticky': 'NWES'},
				'borderfull': {'highlightthickness': 1, 'highlightbackground': 'gray'}
				}
			}
		self.tooltips = []

		# window's settings #
		width, height, file = self.vars.get('minWidth'), self.vars.get('minHeight'), '%s\\tmp.ico' % self.vars.get('workDir')
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

		# prepare elements #
		self.__prep_elems()
		self.__set_data()

		# set icon #
		file = self.__get_file('tmp.ico')
		while isfile(file):
			file = self.__get_file('{}.ico'.format(time()))
		with open(file, 'wb') as f:
			f.write(b64decode('AAABAAEAEBAQAAEABAAoAQAAFgAAACgAAAAQAAAAIAAAAAEABAAAAAAAgAAAAAAAAAAAAAAAEAAAAAAAAAAAAAAA////AHBwcACF/4sA8PDwAGv//wDb29sAj4+PAIWF/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABEREREAAAAiRERERCIAAHdmZmZmdwAAdwAAAAB3AAB3d3d3d3cAAHg3d3d3VwAAd3d3d3d3AAAAEREREQAAAAAQAAABAAAAABEREREAAAAAEAAAAQAAAAARERZmAAAAABAABmAAAAAAEREWAAAAAAAAAAAAAAD//wAA8A8AAMADAADAAwAAwAMAAMADAADAAwAAwAMAAPAPAADwDwAA8A8AAPAPAADwDwAA8B8AAPA/AAD//wAA'))
		self.root.iconbitmap(file)
		remove(file)

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
			vals.sort(key=lambda x: list(map(self.__exint, x[1].split('.'))))
			vals.sort(key=lambda x: x[2])
			for i, (iid, *_) in enumerate(vals):
				self.elem.get('tree-selected').move(iid, '', i)

		return wrapper

	def __prep_elems(self):
		main = 'main'

		# mainframe #
		self.elem.update({main: Frame(self.root, background = 'white')})
		self.elem.get(main).pack(fill = 'both', expand = True)

		# create, and configure elements #
		for key, data in [(x, y) for x, y in self.elem.items() if x != main]:
			self.elem.update({key: data.get('type')(self.elem.get(main), **data.get('args'))})
			self.elem.get(key).grid(**data.get('grid'), padx=self.vars.get('pad'), pady=self.vars.get('pad'))
			if nopad := data.get('nopad'):
				if 'x' in nopad:
					self.elem.get(key).grid(padx=0)
				if 'y' in nopad:
					self.elem.get(key).grid(pady=0)
			if stick := data.get('sticky'):
				self.elem.get(key).grid(**stick)
			if border := data.get('borderfull'):
				self.elem.get(key).config(**border)
			if text := data.get('tooltip'):
				self.tooltips.append(ToolTip(self.elem.get(key), msg=text, delay=0.25))

		# grid settings #
		cols, rows = self.elem.get(main).grid_size()
		for i in range(rows):
			self.elem.get(main).rowconfigure(i, weight=1, minsize=40)
		for i in range(cols):
			self.elem.get(main).columnconfigure(i, weight=1, minsize=50)
		self.elem.get(main).columnconfigure(0, weight=1, minsize=130)

		# event binds, styles, and other settings #
		self.vars.get('style').configure('TCheckbutton', background='white')
		self.elem.get('tree-all').bind('<Return>', self.__add_by_btn)
		self.elem.get('tree-selected').bind('<Return>', self.__remove_by_btn)
		self.elem.get('tree-selected').heading('project', text='Projekt')
		self.elem.get('tree-selected').heading('point', text='Punkt')
		self.elem.get('tree-selected').column('project', width=40, minwidth=40)
		self.elem.get('tree-selected').column('point', width=40, minwidth=40)

		# set opening text #
		self.elem.get('text-opening').delete('1.0', 'end')
		self.elem.get('text-opening').insert('1.0', self.vars.get('opening-text'))

	def __set_data(self):
		# check if file exists #
		if not isfile(self.__get_file()):
			return

		# try to read data #
		try:
			with open(self.__get_file(), 'rt', encoding='utf-8') as f:
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

	def __get_date(self, date, dates):
		# check which timeframe is correct #
		date, *dates = list(map(self.__str2date, [date, *dates]))
		if chosen := [[dates[i], dates[i + 1]] for i in range(0, len(dates), 2) if dates[i] <= date <= dates[i + 1]]:
			return ' - '.join(list(map(self.__date2str, *chosen)))
		else:
			return None

	def __make_name(self):
		# get values #
		iid, name = self.elem.get('tree-selected').get_children()[0], []
		project, point = self.elem.get('tree-selected').item(iid, 'values')[0:2]

		# make new filename #
		name.append(sub('\s+', '', project.lower()))
		name.append(''.join(point.split('.')))
		name.append('_'.join(self.vars.get('var-date').get().split('.')[1:3]))

		# check for folders #
		folder = split('[/\\\\]+', self.vars.get('var-filename').get())[:-1]

		# set the filename #
		self.vars.get('var-filename').set('/'.join([*folder, '_'.join(name)]))

	@__sort
	@__check
	def __add_by_btn(self, event):
		iid = self.elem.get('tree-all').focus()

		# check wherever iid correct #
		if not iid:
			self.__throw_error(0)
			return

		# check wherever catalogue was selected #
		if self.elem.get('tree-all').tag_has('catalogue', iid):
			return

		# check wherever already not selected #
		if iid in self.elem.get('tree-selected').get_children():
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
			self.__throw_error(0)
			return

		# remove selected elements #
		self.elem.get('tree-selected').delete(iid)

	@__sort
	@__check
	def __add(self):
		iids = [(x, self.elem.get('tree-all').parent(x)) for x in self.elem.get('tree-all').selection()]

		# check wherever iids correct #
		if not iids:
			self.__throw_error(0)
			return

		# check wherever no parent was selected #
		if not all(list(map(bool, [x for _, x in iids]))):
			self.__throw_error(1)
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
		filename = '{}.docx'.format(self.vars.get('var-filename').get())

		# check wherever file exists #
		if isfile(self.__get_file(filename)):
			if not askokcancel('Plik już istnieje', message='Czy chcesz kontynuować?'):
				self.__throw_error(3)
				return

		# get projects, and prepare text #
		txt = ''
		if beg := self.elem.get('text-opening').get('1.0', 'end-1c'):
			txt = beg + '\n'
		items = self.elem.get('tree-selected').get_children()
		for project, parent in sorted(list(set([self.elem.get('tree-selected').item(iid, 'values')[0::3] for iid in items])), key=lambda x: x[0]):

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
			document.save(self.__get_file(filename))

			showinfo('Zapisywanie', message='Sukces')

		except Exception as e:
			self.__throw_error(e)


	# other functions #
	def __get_file(self, file=None):
		return '{}\\{}'.format(self.vars.get('workDir'), file if file else self.vars.get('workFile'))

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

	def __throw_error(self, error):
		match error:
			case 0:
				msg = 'Nic nie zostało wybrane'
			case 1:
				msg = 'Wybrany został projekt'
			case 2:
				msg = 'Punkt(y) już został wybrany'
			case 3:
				msg = 'Plik o tej nazwie już istnieje'
			case 4:
				msg = 'Data poza zasięgiem'
			case _:
				msg = error
		showerror(title='Błąd', message=msg)


# 'scroll-all': {
# 	'type': Scrollbar,
# 	'args': {'orient': 'vertical'},
# 	'grid': {'row': 0, 'column': 1, 'rowspan': 5},
# 	'sticky': {'sticky': 'NWES'},
# 	'nopad': 'xy'
# 	},

# self.elem.get(main).columnconfigure(1, weight=1, minsize=20)
# self.elem.get('tree-all').configure(xscrollcommand=self.elem.get('scroll-all').set)
# self.elem.get('scroll-all').configure(command=self.elem.get('tree-all').yview)

if __name__ == '__main__':
	TaxPrinter()
