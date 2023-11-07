from os import getcwd, remove
from os.path import isfile
from json import loads, dumps
from datetime import date
from time import time

from base64 import b64decode

from tkinter import Tk, StringVar as StrVar, Menu, Frame, Text
from tkinter.messagebox import showerror, showinfo, askokcancel
from tkinter.filedialog import askopenfilename
from tkinter.ttk import Entry, Button, Treeview, Label
from tkcalendar import DateEntry
from tktooltip import ToolTip

class DataEditor:
	def __init__(self):
		# root declaration #
		self.root = Tk()

		# declearing variables, and elements #
		self.vars = {
			'title': 'Data Editor',
			'minWidth': 400,
			'minHeight': 450,
			'maxSize': 650,
			'workDir': getcwd(),
			'workFile': 'data.json',
			'pad': 5,
			'unsaved': False,
			'var-name': StrVar(),
			'var-date-beg': StrVar(),
			'var-date-end': StrVar(),
			'tags': ('b', 'i', 'u', 's')
			}
		self.elem = {
			'tree-points': {
				'type': Treeview,
				'args': {'selectmode': 'browse', 'show': 'tree'},
				'grid': {'row': 0, 'column': 0, 'rowspan': 6},
				'sticky': 'NWES'
				},
			'tree-dates': {
				'type': Treeview,
				'args': {'selectmode': 'browse', 'show': 'tree'},
				'grid': {'row': 0, 'column': 1, 'columnspan': 5},
				'sticky': 'NWES'
				},
			'cal-dates-beg': {
				'type': DateEntry,
				'args': {'date_pattern': 'd.M.y', 'locale': 'pl_PL', 'textvariable': self.vars.get('var-date-beg')},
				'grid': {'row': 1, 'column': 1, 'columnspan': 2},
				'tooltip': 'Data początku okresu'
				},
			'cal-dates-end': {
				'type': DateEntry,
				'args': {'date_pattern': 'd.M.y', 'locale': 'pl_PL', 'textvariable': self.vars.get('var-date-end')},
				'grid': {'row': 2, 'column': 1, 'columnspan': 2},
				'tooltip': 'Data końca okresu'
				},
			'btn-dates-add': {
				'type': Button,
				'args': {'text': '\ud83d\udfa4', 'command': self.__add_date},
				'grid': {'row': 1, 'column': 3},
				'tooltip': 'Dodaj wskazane daty'
				},
			'btn-dates-del': {
				'type': Button,
				'args': {'text': '\u232b', 'command': self.__delete_date},
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
				'args': {'textvariable': self.vars.get('var-name')},
				'grid': {'row': 4, 'column': 1, 'columnspan': 5},
				'sticky': 'NWES'
				},
			'btn-points-add-project': {
				'type': Button,
				'args': {'text': '\ud83d\uddbf', 'command': self.__add_catalogue},
				'grid': {'row': 5, 'column': 1},
				'tooltip': 'Dodaj projekt'
				},
			'btn-points-add-entry': {
				'type': Button,
				'args': {'text': '\ud83d\udfa4', 'command': self.__add_item},
				'grid': {'row': 5, 'column': 2},
				'tooltip': 'Dodaj punkt kosztorysu'
				},
			'btn-points-change': {
				'type': Button,
				'args': {'text':'\u270e','command': self.__change_item},
				'grid': {'row': 5, 'column': 4},
				'tooltip': 'Zmień nazwę elementu'
				},
			'btn-points-del': {
				'type': Button,
				'args': {'text':'\u232b','command': self.__delete_item},
				'grid': {'row': 5, 'column': 5},
				'tooltip': 'Usuń element'
				},
			'text-field': {
				'type': Text,
				'args': {},
				'grid': {'row': 6, 'column': 0, 'columnspan': 6},
				'sticky': 'NWES',
				'borderfull': {'highlightthickness': 1, 'highlightbackground': 'gray'}
				},
			'btn-text': {
				'type': Button,
				'args': {'text': '\u166d', 'state': 'disabled', 'command': self.__save_text},
				'grid': {'row': 7, 'column': 1},
				'tooltip': 'Zapisz tekst'
				},
			'btn-save': {
				'type': Button,
				'args': {'text': 'Zapisz', 'command': self.__save_file, 'state': 'disabled'},
				'grid': {'row': 7, 'column': 4, 'columnspan': 2},
				'tooltip': 'Zapisz zmainy w pliku'
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

		# protocols #
		self.root.protocol('WM_DELETE_WINDOW', self.__close)

		# prepare elements #
		self.__prep_elems()
		self.__set_data()

		# set icon #
		file = self.__get_file('tmp.ico')
		while isfile(file):
			file = self.__get_file('{}.ico'.format(time()))
		with open(file, 'wb') as f:
			f.write(b64decode('AAABAAEAEBAQAAEABAAoAQAAFgAAACgAAAAQAAAAIAAAAAEABAAAAAAAgAAAAAAAAAAAAAAAEAAAAAAAAAAAAAAAlO//AP///wCUlP8ASv/qANvb2wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACIiIiIDAAAAIiIiIgUAAAAgAAACBAAAACIiIiIEAAAAIAAAAgQAAAAiIiIiBAAAACAAAAIEAAAAIiIiIgQAAAAgAAACBAAAACIiJVUBAAAAIAAFUAAAAAAiIiUAAAAAAAAAAAAAAAAAAAAAAAAAD//wAA//8AAOAXAADgFwAA4BcAAOAXAADgFwAA4BcAAOAXAADgFwAA4BcAAOAXAADgNwAA4H8AAP//AAD//wAA'))
		self.root.iconbitmap(file)
		remove(file)

		# start program #
		self.root.mainloop()

	# decorator for changes #
	def __safecheck(f):
		def wrapper(self):
			f(self)
			self.elem.get('btn-save').config(state='normal')
			self.vars.update({'unsaved': True})
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
				self.elem.get(key).grid(sticky=stick)
			if border := data.get('borderfull'):
				self.elem.get(key).config(**border)
			if (text := data.get('tooltip')):
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
		self.elem.get('menu-file').add_command(label='Wyjdź', command=self.__close)

		# grid settings #
		cols, rows = self.elem.get(main).grid_size()
		for i in range(rows):
			self.elem.get(main).rowconfigure(i, weight=1, minsize=40)
		for i in range(cols):
			self.elem.get(main).columnconfigure(i, weight=1, minsize=50)

		# event binds, styles, and other settings #
		self.elem.get(main).rowconfigure(0, weight=3, minsize=100)
		self.elem.get(main).rowconfigure(6, weight=3, minsize=100)
		self.elem.get(main).columnconfigure(0, weight=3, minsize=150)
		self.elem.get('tree-points').bind('<<TreeviewSelect>>', self.__points_select)
		self.elem.get('text-field').bind('<Key>', self.__text_change)

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
					self.elem.get('tree-points').insert('', 'end', i, text=project.get('name'), values=vals, open=False, tags=['catalogue'])
					for point in project.get('points'):
						self.elem.get('tree-points').insert(i, 'end', text=point.get('point'), values=[point.get('text')])

		except Exception as e:
			self.__throw_error(e)

	def __clear_data(self):
		# clear data #
		self.elem.get('tree-points').delete(*self.elem.get('tree-points').get_children())
		self.elem.get('tree-dates').delete(*self.elem.get('tree-dates').get_children())
		self.elem.get('text-field').delete('1.0', 'end')
		
	def __points_select(self, event):
		iid = self.elem.get('tree-points').focus()

		# check wherever iid correct #
		if not iid:
			return

		# set text #
		vals = self.elem.get('tree-points').item(iid, 'values')
		self.elem.get('text-field').delete('1.0', 'end')
		self.elem.get('text-field').insert('1.0', vals[0])

		# get catalogue #
		if not self.elem.get('tree-points').tag_has('catalogue', iid):
			iid = self.elem.get('tree-points').parent(iid)

		# set dates #	
		vals = self.elem.get('tree-points').item(iid, 'values')
		for item in self.elem.get('tree-dates').get_children():
			self.elem.get('tree-dates').delete(item)
		for i in range(1, len(vals), 2):
			self.elem.get('tree-dates').insert('', 'end', text='{} - {}'.format(vals[i], vals[i + 1]), values=[iid, vals[i], vals[i + 1]])

	def __text_change(self, event):
		# check wherever iid correct #
		if not self.elem.get('tree-points').focus():
			return

		# set button #
		self.elem.get('btn-text').config(text='\u2713', state='normal')

	@__safecheck
	def __save_text(self):
		iid = self.elem.get('tree-points').focus()

		# check wherever element is focused #
		if not iid:
			self.__throw_error(0)
			return

		# assign text #
		text = self.elem.get('text-field').get('1.0', 'end-1c')
		vals = list(self.elem.get('tree-points').item(iid, 'values'))
		vals[0] = text.strip()
		self.elem.get('tree-points').item(iid, values=vals)

		# set button #
		self.elem.get('btn-text').config(text='\u166d', state='disabled')

	@__safecheck
	def __delete_item(self):
		iid = self.elem.get('tree-points').focus()

		# check wherever element is focused #
		if not iid:
			self.__throw_error(0)
			return

		# delete item #
		self.elem.get('text-field').delete('1.0', 'end')
		self.elem.get('tree-points').delete(iid)

		# clear view #
		self.elem.get('tree-dates').delete(*self.elem.get('tree-dates').get_children())

	@__safecheck
	def __add_catalogue(self):
		name = self.vars.get('var-name').get()

		# check if name was given #
		if not name:
			self.__throw_error(1)
			return

		# check if name not taken #
		if name in [self.elem.get('tree-points').item(x, 'text') for x in self.elem.get('tree-points').get_children()]:
			self.__throw_error(2)
			return

		# add catalogue #
		focus = self.elem.get('tree-points').insert('', 'end', text=name, values=[''], tags=['catalogue'])
		self.elem.get('tree-points').selection_set(focus)
		self.elem.get('tree-points').focus(focus)

	@__safecheck
	def __add_item(self):
		iid, name, index = self.elem.get('tree-points').focus(), self.vars.get('var-name').get(), 'end'

		# check wherever element is focused #
		if not iid:
			self.__throw_error(0)
			return

		# check if name was given #
		if not name:
			self.__throw_error(1)
			return

		# get catalogue #
		if not self.elem.get('tree-points').tag_has('catalogue', iid):
			index, iid = self.elem.get('tree-points').index(iid), self.elem.get('tree-points').parent(iid)

		# check if name not taken #	
		if name in [self.elem.get('tree-points').item(x, 'text') for x in self.elem.get('tree-points').get_children(iid)]:
			self.__throw_error(2)
			return

		# insert element #
		focus = self.elem.get('tree-points').insert(iid, index, text=name, values=[''])
		self.elem.get('tree-points').selection_set(focus)
		self.elem.get('tree-points').focus(focus)

	@__safecheck
	def __change_item(self):
		iid, name = self.elem.get('tree-points').focus(), self.vars.get('var-name').get()

		# check wherever element is focused #
		if not iid:
			self.__throw_error(0)
			return

		# check if name was given #
		if not name:
			self.__throw_error(1)
			return

		# get catalogue #
		parent = '' if self.elem.get('tree-points').tag_has('catalogue', iid) else self.elem.get('tree-points').parent(iid)

		# check if name not taken #
		if name in [self.elem.get('tree-points').item(x, 'text') for x in self.elem.get('tree-points').get_children(parent)]:
			self.__throw_error(2)
			return

		# set new name #
		self.elem.get('tree-points').item(iid, text=name)

	@__safecheck
	def __delete_date(self):
		iid = self.elem.get('tree-dates').focus()

		# check wherever element is focused #
		if not iid:
			self.__throw_error(0)
			return

		# update values #
		parentiid = self.elem.get('tree-dates').item(iid, 'values')[0]
		vals = [self.elem.get('tree-points').item(parentiid, 'values')[0]]
		for date in [self.elem.get('tree-dates').item(x, 'values')[1:3] for x in self.elem.get('tree-dates').get_children() if x != iid]: 
			vals.extend(date)
		self.elem.get('tree-points').item(parentiid, values=vals)

		# delete item #
		self.elem.get('tree-dates').delete(iid)

	@__safecheck
	def __add_date(self):
		# get parent #
		parentiid = self.elem.get('tree-points').focus()
		if not self.elem.get('tree-points').tag_has('catalogue', parentiid):
			parentiid = self.elem.get('tree-points').parent(parentiid)

		# check wherever element is focused #
		if not parentiid:
			self.__throw_error(0)
			return

		# check if range is correct #
		dates = tuple(map(self.__str2date, [self.vars.get('var-date-beg').get(), self.vars.get('var-date-end').get()]))
		if dates[1] <= dates[0]:
			self.__throw_error(4)
			return

		# check for other dates #
		index = 'end'
		if self.elem.get('tree-dates').get_children():
			# check if date not taken #
			currentDates = [(iid, tuple(map(self.__str2date, (x, y)))) for iid, (x, y) in [(iid, self.elem.get('tree-dates').item(iid, 'values')[1:3]) for iid in self.elem.get('tree-dates').get_children()]]
			if not all([True if dates[1] < start or end < dates[0] else False for _, (start, end) in currentDates]):
				self.__throw_error(5)
				return

			# get right index #
			ommit = [0, len(currentDates) - 1]
			if dates[1] < currentDates[ommit[0]][1][0]:
				index = 0
			elif currentDates[ommit[1]][1][0] < dates[0]:
				index = ommit[1] + 1
			else:
				for i, (iid, _) in enumerate(currentDates):
					if i in ommit:
						continue
					if currentDates[i - 1][1][1]  < dates[0] and dates[1] < currentDates[i][1][0]:
						index = i

		# insert in right postion #
		dates = tuple(map(self.__date2str, dates))
		self.elem.get('tree-dates').insert('', index, text='{} - {}'.format(*dates), values=[parentiid, dates[0], dates[1]])

		# update values #
		vals = [self.elem.get('tree-points').item(parentiid, 'values')[0]]
		for date in [self.elem.get('tree-dates').item(x, 'values')[1:3] for x in self.elem.get('tree-dates').get_children()]: 
			vals.extend(date)
		self.elem.get('tree-points').item(parentiid, values=vals)

	def __save_file(self):
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
			file.append({
				'name': self.elem.get('tree-points').item(project, 'text'),
				'description': vals[0],
				'dates': dates,
				'points': sorted(points, key=lambda x: list(map(self.__exint, x.get('point').split('.'))))
				})

		# try to save file #
		try:
			with open(self.__get_file(), 'wt', encoding='utf-8') as f:
				f.write(dumps(file))
			showinfo(title='Zapisywanie', message='Sukces')
			self.elem.get('btn-save').config(state='disabled')
			self.vars.update({'unsaved': False})

		except Exception as e:
			self.__throw_error(e)

	def __close(self):
		if self.vars.get('unsaved'):
			if self.__ask_changes():
				self.root.destroy()	
		else:
			self.root.destroy()

	def __show_format(self):
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

	def __show_help(self):
		msg = \
			'Program do edytowania danych opisów.\n' \
			'\n' \
			'\u2022 Większość elemntów wyświetla opisy po najechaniu.\n' \
			'\u2022 Po kolumnach można poruszać się za pomocą strzałek.\n' \
			'\u2022 Dane podstawowo zapisane są w pliku "data.json",\n' \
			'można je przeładować, bądź wybrać inny plik danych.\n' \
			'\u2022 Aplikacja powinna znajdować się w tym samym, bądź\n' \
			'wyższym folderze, co plik danych.\n' \
			'\n' \
			'Program napisany w Python, z pomocą TKinter.'
		showinfo(title='Pomoc', message=msg)

	def __select_file(self):
		# get new path #
		path = askopenfilename(title='Wybierz plik', initialdir=self.vars.get('workDir'), filetypes=(('Plik JSON', '.json'), ), multiple=False).replace('/', '\\')
		if not path:
			return

		# check if path correct #
		if self.vars.get('workDir') not in path:
			self.__throw_error(6)
			return

		# check if extension correct #
		if '.json' not in path[-5:]:
			self.__throw_error(7)
			return

		# set new file #
		self.vars.update({'workFile': path.removeprefix(self.vars.get('workDir'))})

	def __reload(self):
		# check for changes #
		if self.vars.get('unsaved'):
			if not self.__ask_changes():
				return

		# reload data #
		self.__clear_data()
		self.__set_data()
		self.elem.get('btn-save').config(state='disabled')
		self.vars.update({'unsaved': False})

	def __ask_changes(self):
		return askokcancel(title='Niezapisane zmiany', message='Czy chcesz kontynuować?')
	
	def __throw_error(self, error):
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
				msg = 'Zła ścieżka pliku'
			case 7:
				msg = 'Złe rozszerzenie pliku'
			case _:
				msg = error
		showerror(title='Błąd', message=msg)


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


if __name__ == '__main__':
	DataEditor()
