from abc import ABC, abstractmethod

from sys import version_info
from os import getcwd
from datetime import date

from base64 import b64decode

from tkinter import Tk, PhotoImage, Menu
from tkinter.ttk import Style, Frame
from tkinter.messagebox import showerror
from tktooltip import ToolTip

class WindowApp(ABC):
	def __init__(self, title, icon, minSize, maxSize):
		# root declaration #
		self.root = Tk()

		# declearing variables, and elements #
		self.vars = {
			'title': title,
			'icon': icon,
			'size': {
				'min': minSize,
				'max': maxSize
				},
			'file': '{}\\{}'.format(getcwd(), 'data.json'),
			'py-ver': '{}.{}.{}'.format(version_info.major, version_info.minor, version_info.micro),
			'tk-ver': self.root.tk.call('info', 'patchlevel'),
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
		self.root.iconphoto(False, PhotoImage(data=b64decode(self.vars['icon'])))

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
		self.elem.update({main: Frame(self.root)})
		self.elem[main].pack(fill='both', expand=True)

		# create, and configure elements #
		for key, data in ((x, y) for x, y in self.elem.items() if x != main):
			self.elem.update({key: data['type'](self.elem[main], **data.get('args', {}))})
			options = {'padx': self.vars['pad'], 'pady': self.vars['pad']}
			if nopad := data.get('nopad'):
				if 'x' in nopad:
					options.update({'padx': 0})
				if 'y' in nopad:
					options.update({'pady': 0})
			if stick := data.get('sticky'):
				options.update({'sticky': stick})
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
			self.elem.update({key: Menu(self.root, tearoff=0)})
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
		for elem, (act, cmnd) in self.binds:
			self.elem.get(elem).bind(act, cmnd)

	@abstractmethod
	def pre(self):
		pass

	@abstractmethod
	def post(self):
		pass

	def show_help(self):
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
			'Python {}\n' \
			'TKinter {}' \
			.format(self.vars['py-ver'], self.vars['tk-ver'])
		showinfo(title='Pomoc', message=msg)

	@abstractmethod
	def close(self):
		pass

	def throw_error(self, error):
		match error:
			case 0:
				msg = 'Nieznany błąd'
			case _:
				msg = self.vars['errors'].get(error, error)
		showerror(title='Błąd', message=msg)

	# other functions #
	@staticmethod
	def str2date(d):
		return date(*reversed(tuple(int(x) for x in d.split('.'))))

	@staticmethod
	def date2str(d):
		return '{}.{}.{}'.format(d.day, d.month, d.year)

	@staticmethod
	def roman2int(n):
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