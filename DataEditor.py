from utilities import PrinterApp

from os import getcwd
from os.path import isfile
from json import loads, dumps
from functools import wraps

from tkinter import StringVar as StrVar, Text
from tkinter.ttk import Entry, Button, Treeview, Scrollbar, Label
from tkinter.messagebox import showinfo, askokcancel
from tkinter.filedialog import askopenfilename
from tkcalendar import DateEntry

# from accessify import private

class DataEditor(PrinterApp):
    # decorator for changes #
    def safecheck(f):
        @wraps(f)
        def wrapper(self, *args, **kwargs):
            f(self)
            self.elem['btn-save'].config(state='normal')
            self.vars['unsaved'] = True
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
                    self.elem['tree-points'].insert('', 'end', i, text=project['name'], values=vals, open=False, tags=['catalogue'])
                    for point in project['points']:
                        self.elem['tree-points'].insert(i, 'end', text=point['point'], values=[point['text']])

        except Exception as e:
            self.throw_error(e)

    def clear_data(self):
        # clear data #
        self.elem['tree-points'].delete(*self.elem['tree-points'].get_children())
        self.elem['tree-dates'].delete(*self.elem['tree-dates'].get_children())
        self.elem['txt-field'].delete('1.0', 'end')

    def points_select(self, event):
        iid = self.elem['tree-points'].focus()

        # check wherever iid correct #
        if not iid:
            return

        # set texts #
        self.vars['var']['name'].set(self.elem['tree-points'].item(iid, 'text'))
        self.elem['txt-field'].delete('1.0', 'end')
        self.elem['txt-field'].insert('1.0', self.elem['tree-points'].item(iid, 'values')[0])

        # get catalogue #
        if not self.elem['tree-points'].tag_has('catalogue', iid):
            iid = self.elem['tree-points'].parent(iid)

        # set dates #   
        vals = self.elem['tree-points'].item(iid, 'values')
        for item in self.elem['tree-dates'].get_children():
            self.elem['tree-dates'].delete(item)
        for i in range(1, len(vals), 2):
            self.elem['tree-dates'].insert('', 'end', text='{} - {}'.format(vals[i], vals[i + 1]), values=[iid, vals[i], vals[i + 1]])

    def text_change(self, event):
        # check wherever iid correct #
        if not self.elem['tree-points'].focus():
            return

        # set button #
        self.elem['btn-text'].config(text='\u2713', state='normal')

    @safecheck
    def save_text(self):
        iid = self.elem['tree-points'].focus()

        # check wherever element is focused #
        if not iid:
            self.throw_error(201)
            return

        # assign text #
        text = self.elem['txt-field'].get('1.0', 'end-1c')
        vals = list(self.elem['tree-points'].item(iid, 'values'))
        vals[0] = text.strip()
        self.elem['tree-points'].item(iid, values=vals)

        # set button #
        self.elem['btn-text'].config(text='\u166d', state='disabled')

    @safecheck
    def delete_item(self):
        iid = self.elem['tree-points'].focus()

        # check wherever element is focused #
        if not iid:
            self.throw_error(201)
            return

        # delete item #
        self.elem['txt-field'].delete('1.0', 'end')
        self.elem['tree-points'].delete(iid)

        # clear view #
        self.elem['tree-dates'].delete(*self.elem['tree-dates'].get_children())

    @safecheck
    def add_catalogue(self):
        name = self.vars['var']['name'].get()

        # check if name was given #
        if not name:
            self.throw_error(501)
            return

        # check if name not taken #
        if name in tuple(self.elem['tree-points'].item(x, 'text') for x in self.elem['tree-points'].get_children()):
            self.throw_error(502)
            return

        # add catalogue #
        focus = self.elem['tree-points'].insert('', 'end', text=name, values=[''], tags=['catalogue'])
        self.elem['tree-points'].selection_set(focus)
        self.elem['tree-points'].focus(focus)

    @safecheck
    def add_item(self):
        iid, name, index = self.elem['tree-points'].focus(), self.vars['var']['name'].get(), 'end'

        # check wherever element is focused #
        if not iid:
            self.throw_error(201)
            return

        # check if name was given #
        if not name:
            self.throw_error(501)
            return

        # get catalogue #
        if not self.elem['tree-points'].tag_has('catalogue', iid):
            index, iid = self.elem['tree-points'].index(iid), self.elem['tree-points'].parent(iid)

        # check if name not taken # 
        if name in tuple(self.elem['tree-points'].item(x, 'text') for x in self.elem['tree-points'].get_children(iid)):
            self.throw_error(502)
            return

        # insert element #
        focus = self.elem['tree-points'].insert(iid, index, text=name, values=[''])
        self.elem['tree-points'].selection_set(focus)
        self.elem['tree-points'].focus(focus)

    @safecheck
    def change_item(self):
        iid, name = self.elem['tree-points'].focus(), self.vars['var']['name'].get()

        # check wherever element is focused #
        if not iid:
            self.throw_error(201)
            return

        # check if name was given #
        if not name:
            self.throw_error(501)
            return

        # get catalogue #
        parent = '' if self.elem['tree-points'].tag_has('catalogue', iid) else self.elem['tree-points'].parent(iid)

        # check if name not taken #
        if name in tuple(self.elem['tree-points'].item(x, 'text') for x in self.elem['tree-points'].get_children(parent)):
            self.throw_error(502)
            return

        # set new name #
        self.elem['tree-points'].item(iid, text=name)

    @safecheck
    def delete_date(self):
        iid = self.elem['tree-dates'].focus()

        # check wherever element is focused #
        if not iid:
            self.throw_error(201)
            return

        # update values #
        parent_iid = self.elem['tree-dates'].item(iid, 'values')[0]
        vals = [self.elem['tree-points'].item(parent_iid, 'values')[0]]
        for date in [self.elem['tree-dates'].item(x, 'values')[1:3] for x in self.elem['tree-dates'].get_children() if x != iid]: 
            vals.extend(date)
        self.elem['tree-points'].item(parent_iid, values=vals)

        # delete item #
        self.elem['tree-dates'].delete(iid)

    @safecheck
    def add_date(self):
        # get parent #
        parent_iid = self.elem['tree-points'].focus()
        if not self.elem['tree-points'].tag_has('catalogue', parent_iid):
            parent_iid = self.elem['tree-points'].parent(parent_iid)

        # check wherever element is focused #
        if not parent_iid:
            self.throw_error(201)
            return

        # check if range is correct #
        dates = tuple(self.str2date(x) for x in (self.vars['var']['date-beg'].get(), self.vars['var']['date-end'].get()))
        if dates[1] <= dates[0]:
            self.throw_error(503)
            return

        # check for other dates #
        index = 'end'
        if self.elem['tree-dates'].get_children():
            # check if date not taken #
            current_dates = tuple((iid, tuple(self.str2date(z) for z in (x, y))) for iid, (x, y) in tuple((iid, self.elem['tree-dates'].item(iid, 'values')[1:3]) for iid in self.elem['tree-dates'].get_children()))
            if not all(dates[1] < start or end < dates[0] for _, (start, end) in current_dates):
                self.throw_error(504)
                return

            # get right index #
            ommit = len(current_dates) - 1
            if dates[1] < current_dates[0][1][0]:
                index = 0
            elif current_dates[ommit][1][0] < dates[0]:
                index = ommit + 1
            else:
                for i, (iid, _) in enumerate(current_dates):
                    if i in (0, ommit):
                        continue
                    if current_dates[i - 1][1][1]  < dates[0] and dates[1] < current_dates[i][1][0]:
                        index = i

        # insert in right postion #
        dates = tuple(self.date2str(x) for x in dates)
        self.elem['tree-dates'].insert('', index, text=' - '.join(dates), values=[parent_iid, *dates])

        # update values #
        vals = [self.elem['tree-points'].item(parent_iid, 'values')[0]]
        for date in tuple(self.elem['tree-dates'].item(x, 'values')[1:3] for x in self.elem['tree-dates'].get_children()): 
            vals.extend(date)
        self.elem['tree-points'].item(parent_iid, values=vals)

    def save_file(self):
        file = []

        # get iterable values, and make obj #
        for project in self.elem['tree-points'].get_children():
            vals, dates, points = self.elem['tree-points'].item(project, 'values'), [], []
            for i in range(1, len(vals), 2):
                dates.append({'from': vals[i], 'to': vals[i + 1]})
            for item in self.elem['tree-points'].get_children(project):
                points.append({
                    'point': self.elem['tree-points'].item(item, 'text'),
                    'text': self.elem['tree-points'].item(item, 'values')[0]
                    })
            points.sort(key=lambda x: tuple(self.roman2int(y) for y in x['point'].split('.')))
            file.append({
                'name': self.elem['tree-points'].item(project, 'text'),
                'description': vals[0],
                'dates': dates,
                'points': points
                })
        file.sort(key=lambda x: x['name'])

        # try to save file #
        try:
            with open(self.vars['file'], 'wt', encoding='utf-8') as f:
                f.write(dumps(file))
            showinfo(title='Zapisywanie', message='Sukces')
            self.elem['btn-save'].config(state='disabled')
            self.vars['unsaved'] = False

        except Exception as e:
            self.throw_error(e)

    def show_format(self):
        msg = (
            'Formatowanie opisu:\n'
            '\n'
            '<b> ... </b> - pogrubienie\n'
            '<i> ... </i> - kursywa\n'
            '<u> ... </u> - podkreślenie\n'
            '<s> ... </s> - przekreślenie\n'
            '\n'
            '<d> - okres (tylko w opisie projektu)\n'
            '\n'
            '<br> - nowa linia'
            )
        showinfo(title='Formatowanie', message=msg)

    def reload(self):
        # check for changes #
        if self.vars['unsaved']:
            if not self.ask_changes():
                return

        # reload data #
        self.clear_data()
        self.set_data()
        self.elem['btn-save'].config(state='disabled')
        self.vars['unsaved'] = False

    def ask_changes(self):
        return askokcancel(title='Niezapisane zmiany', message='Czy chcesz kontynuować?')

    # overridden #
    def pre(self):
        # important #
        self.vars.update({
            'title': 'Data Editor',
            'size': {
                'min': (400, 450),
                'max': (650, 650)
                }
            })
        
        # add, and set variables #
        self.vars.update({
            'file': '{}\\{}'.format(getcwd(), 'data.json'),
            'unsaved': False
            })
        self.vars['var'].update({
            'name': StrVar(),
            'date-beg': StrVar(),
            'date-end': StrVar()
            })
        self.elem.update({
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
                'args': {'date_pattern': 'd.M.y', 'locale': 'pl_PL', 'textvariable': self.vars['var']['date-beg']},
                'grid': {'row': 1, 'column': 1, 'columnspan': 2},
                'tooltip': 'Data początku okresu'
                },
            'cal-dates-end': {
                'type': DateEntry,
                'args': {'date_pattern': 'd.M.y', 'locale': 'pl_PL', 'textvariable': self.vars['var']['date-end']},
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
                'args': {'textvariable': self.vars['var']['name']},
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
                'args': {'wrap':'word'},
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
                'tooltip': 'Zapisz zmiany w pliku'
                }
            })
        self.grid['row'].update({
                0: {'weight': 3, 'minsize': 100},
                6: {'weight': 3, 'minsize': 100}
                })
        self.grid['col'].update({
                0: {'weight': 3, 'minsize': 150}
                })
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
            ('tree-points', ('<<TreeviewSelect>>', self.points_select)),
            ('txt-field', ('<Key>', self.text_change))
            ])

    def post(self):
        # styles, other settings, and actions #
        self.vars['style'].configure('TFrame', background='white')
        self.elem['tree-points'].configure(yscrollcommand=self.elem['scroll-points'].set)
        self.elem['tree-dates'].configure(yscrollcommand=self.elem['scroll-dates'].set)
        self.elem['scroll-points'].configure(command=self.elem['tree-points'].yview)
        self.elem['scroll-dates'].configure(command=self.elem['tree-dates'].yview)

        self.set_data()

    def close(self):
        if self.vars['unsaved']:
            if self.ask_changes():
                self.root.destroy() 
        else:
            self.root.destroy()


if __name__ == '__main__':
    DataEditor()
