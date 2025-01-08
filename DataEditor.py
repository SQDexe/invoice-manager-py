from typing import Any, Optional
from collections.abc import Iterator
from datetime import date
from tkinter import Event

from utils import PrinterApp, Function
from utils.consts import MIN_DATE
from utils.funcs import flatten, pair_cross, pair_up, point2tuple, sort2return

from os import getcwd
from os.path import isfile, join
from json import loads, dumps
from re import compile as recompile
from functools import wraps

from tkinter import StringVar as StrVar, Text
from tkinter.ttk import Entry, Button, Treeview, Scrollbar, Label
from tkinter.messagebox import showinfo, askokcancel
from tkcalendar import DateEntry

class DataEditor(PrinterApp):
    # attributes redeclaration #
    __slots__: tuple[()] = ()
    
    # decorator for changes #
    def safecheck(func: Function) -> Function:
        @wraps(func)
        def wrapper(self, *args: Any, **kwargs: Any) -> None:
            func(self, *args, **kwargs)
            if not self.vars.unsaved:
                self.elem.btn_save.config(state='normal')
                self.vars.unsaved = True
        return wrapper

    def set_data(self) -> None:
        # check if file exists #
        if not isfile(self.vars.file):
            self.throw_error(1)
            return

        # try to read data #
        data: list[dict[str, Any]] = []
        try:
            with open(self.vars.file, 'rt', encoding='utf-8') as file:
                data.extend(loads(file.read()))

        except Exception as e:
            self.throw_error(0, str(e))

        else:
            for i, project in enumerate(data):
                self.elem.tree_points.insert('', 'end', i, text=project['name'], values=(
                  project['description'],
                  *flatten((dates['from'], dates['to']) for dates in project['dates'])
                  ), open=False, tags=('catalogue', ))
                for point in project['points']:
                    self.elem.tree_points.insert(i, 'end', text=point['point'], values=(point['text'], ))

    def clear_data(self) -> None:
        # clear data #
        self.elem.tree_points.delete(*self.elem.tree_points.get_children())
        self.elem.tree_dates.delete(*self.elem.tree_dates.get_children())
        self.elem.txt_field.delete('1.0', 'end')

    def points_select(self, event: Optional[Event]=None, /) -> None:
        iid: str = self.elem.tree_points.focus()

        # check whether iid correct #
        if not iid:
            return

        # set texts #
        self.vars.var.name.set(self.elem.tree_points.item(iid, 'text'))
        self.elem.txt_field.delete('1.0', 'end')
        self.elem.txt_field.insert('1.0', self.elem.tree_points.item(iid, 'values')[0])

        # get catalogue #
        if not self.elem.tree_points.tag_has('catalogue', iid):
            iid = self.elem.tree_points.parent(iid)

        # set dates #   
        dates: tuple[str, ...] = self.elem.tree_points.item(iid, 'values')[1:]
        self.elem.tree_dates.delete(*self.elem.tree_dates.get_children())
        for beg, end in pair_up(dates):
            self.elem.tree_dates.insert('', 'end', values=(beg, end))

    def text_change(self, event: Optional[Event]=None, /) -> None:
        # check whether iid correct #
        if not self.elem.tree_points.focus():
            return

        # set button #
        self.elem.btn_text.config(text='\u2713', state='normal')

    @safecheck
    def text_save(self, event: Optional[Event]=None, /) -> None:
        iid: str = self.elem.tree_points.focus()

        # check whether element is focused #
        if not iid:
            self.throw_error(201)
            return

        # assign text #
        text: str = self.elem.txt_field.get('1.0', 'end-1c').strip()
        dates: tuple[str, ...] = self.elem.tree_points.item(iid, 'values')[1:]
        self.elem.tree_points.item(iid, values=(text, *dates))

        # set button #
        self.elem.btn_text.config(text='\u166d', state='disabled')

    @safecheck
    def delete_item(self) -> None:
        iid: str = self.elem.tree_points.focus()

        # check whether element is focused #
        if not iid:
            self.throw_error(201)
            return

        # delete item #
        self.elem.txt_field.delete('1.0', 'end')
        self.elem.tree_points.delete(iid)

        # clear view #
        self.elem.tree_dates.delete(*self.elem.tree_dates.get_children())

    @safecheck
    def add_catalogue(self) -> None:
        name: str = self.vars.var.name.get().strip()

        # check if name was given #
        if not name:
            self.throw_error(501)
            return

        # check if name not taken #
        if name in (self.elem.tree_points.item(child, 'text') for child in self.elem.tree_points.get_children()):
            self.throw_error(502)
            return

        # add catalogue #
        new_iid: str = self.elem.tree_points.insert('', 'end', text=name, values=('', ), tags=('catalogue', ))
        self.elem.tree_points.selection_set(new_iid)
        self.elem.tree_points.focus(new_iid)

    @safecheck
    def add_item(self) -> None:
        iid: str = self.elem.tree_points.focus()
        name: str = self.vars.var.name.get().strip()
        index: str = 'end'

        # check whether element is focused #
        if not iid:
            self.throw_error(201)
            return

        # check if name was given #
        if not name:
            self.throw_error(501)
            return

        # get catalogue #
        if not self.elem.tree_points.tag_has('catalogue', iid):
            index = self.elem.tree_points.index(iid)
            iid = self.elem.tree_points.parent(iid)

        # check if name not taken # 
        if name in (self.elem.tree_points.item(child, 'text') for child in self.elem.tree_points.get_children(iid)):
            self.throw_error(502)
            return

        # check if name is correct for points #
        if not self.vars.patterns.point_name.search(name):
            self.throw_error(506)
            return

        # insert element #
        new_iid: str = self.elem.tree_points.insert(iid, index, text=name, values=('', ))
        self.elem.tree_points.selection_set(new_iid)
        self.elem.tree_points.focus(new_iid)

    @safecheck
    def change_item(self) -> None:
        iid: str = self.elem.tree_points.focus()
        name: str = self.vars.var.name.get().strip()

        # check whether element is focused #
        if not iid:
            self.throw_error(201)
            return

        # check if name was given #
        if not name:
            self.throw_error(501)
            return

        is_catalogue: bool = self.elem.tree_points.tag_has('catalogue', iid)

        # get catalogue #
        parent_iid: str = '' if is_catalogue else self.elem.tree_points.parent(iid)

        # check if name not taken #
        if name in (self.elem.tree_points.item(child, 'text') for child in self.elem.tree_points.get_children(parent_iid)):
            self.throw_error(502)
            return

        # check if name is correct for points #
        if not is_catalogue and not self.vars.patterns.point_name.search(name):
            self.throw_error(506)
            return

        # set new name #
        self.elem.tree_points.item(iid, text=name)

    @safecheck
    def delete_date(self) -> None:
        iid: str = self.elem.tree_dates.focus()

        # check whether element is focused #
        if not iid:
            self.throw_error(201)
            return

        # get parent #
        parent_iid: str = self.elem.tree_points.focus()

        # check whether element is focused #
        if not parent_iid:
            self.throw_error(201)
            return

        # check if selected parent #
        if not self.elem.tree_points.tag_has('catalogue', parent_iid):
            parent_iid = self.elem.tree_points.parent(parent_iid)

        # update values #
        desc: str = self.elem.tree_points.item(parent_iid, 'values')[0]
        dates: tuple[str, ...] = flatten(
          self.elem.tree_dates.item(child, 'values')
          for child in self.elem.tree_dates.get_children()
          if child != iid
          )
        self.elem.tree_points.item(parent_iid, values=(desc, *dates))

        # delete item #
        self.elem.tree_dates.delete(iid)

    @safecheck
    def add_date(self) -> None:
        # get parent #
        parent_iid: str = self.elem.tree_points.focus()

        # check whether element is focused #
        if not parent_iid:
            self.throw_error(201)
            return

        # check if selected parent #
        if not self.elem.tree_points.tag_has('catalogue', parent_iid):
            parent_iid = self.elem.tree_points.parent(parent_iid)

        # check if range is correct #
        beg, end = tuple(self.str2date(x) for x in (self.vars.var.date_beg.get(), self.vars.var.date_end.get()))
        if end <= beg:
            self.throw_error(503)
            return

        # check for other dates #
        index: str
        if childs := self.elem.tree_dates.get_children():
            # check if date not taken #
            current_dates: tuple[tuple[date, date], ...] = tuple(
              tuple(str2date(d) for d in self.elem.tree_dates.item(iid, 'values'))
              for iid in childs
              )
            if not all(end < start or ending < beg for start, ending in current_dates):
                self.throw_error(504)
                return

            # get right index #
            if end < current_dates[0][0]:
                index = '0'
            elif current_dates[-1][1] < beg:
                index = 'end'
            else:
                for i, ((_, ending), (start, _)) in enumerate(pair_cross(current_dates), 1):
                    if ending < beg and end < start:
                        index = str(i)

        # insert in right postion #
        self.elem.tree_dates.insert('', index, values=(
            self.vars.var.date_beg.get(),
            self.vars.var.date_end.get()
            ))

        # update values #
        desc: str = self.elem.tree_points.item(parent_iid, 'values')[0]
        new_dates: Iterator[str] = flatten(
          self.elem.tree_dates.item(child, 'values')
          for child in self.elem.tree_dates.get_children()
          )
        self.elem.tree_points.item(parent_iid, values=(desc, *new_dates))

    def save_file(self) -> None:
        data: list[dict[str, Any]] = []

        # get iterable values, and make obj #
        for project in self.elem.tree_points.get_children():
            name: str = self.elem.tree_points.item(project, 'text')
            desc, *dates = self.elem.tree_points.item(project, 'values')
            dict_dates: list[dict[str, str]] = [{'from': beg, 'to': end} for beg, end in pair_up(dates)]
            points: list[dict[str, Any]] = [
              {'point': self.elem.tree_points.item(item, 'text'), 'text': self.elem.tree_points.item(item, 'values')[0]}
              for item in self.elem.tree_points.get_children(project)
              ]
            data.append({
              'name': name,
              'description': desc,
              'dates': dict_dates,
              'points': sort2return(points, key=lambda x: point2tuple(x['point']))
              })

        # try to save file #
        try:
            with open(self.vars.file, 'wt', encoding='utf-8') as file:
                file.write(dumps(sort2return(data, key=lambda x: x['name'])))

        except Exception as e:
            self.throw_error(0, str(e))

        else:
            showinfo(title='Zapisywanie', message='Sukces')
            self.elem.btn_save.config(state='disabled')
            self.vars.unsaved = False

    def reload(self) -> None:
        # check for changes #
        if self.ask_about_changes():
            return

        # reload data #
        self.clear_data()
        self.set_data()
        self.elem.btn_save.config(state='disabled')
        self.vars.unsaved = False

    def ask_about_changes(self) -> bool:
        return self.vars.unsaved and not askokcancel(title='Niezapisane zmiany', message='Czy chcesz kontynuować?')

    # overridden #
    def pre(self) -> None:
        # important #
        self.vars.title = 'Data Editor'
        self.vars.size.min = (400, 450)
        self.vars.size.max = (650, 650)

        # add, and set variables #
        self.vars.update(
          file = join(getcwd(), 'data.json'),
          unsaved = False
          )
        self.vars.var.update(
          name = StrVar(),
          date_beg = StrVar(),
          date_end = StrVar()
          )
        self.vars.patterns.update(
          point_name = recompile(r'^[\dIVXLCDM]+(\.[\dIVXLCDM]+)*$')
          )
        self.elem.update(
          tree_points = {
            'type': Treeview,
            'args': {'selectmode': 'browse', 'show': 'tree'},
            'grid': {'row': 0, 'column': 0, 'rowspan': 6},
            'sticky': 'NWES'
            },
          scroll_points = {
            'type': Scrollbar,
            'args': {'orient': 'vertical'},
            'grid': {'row': 0, 'column': 0, 'rowspan': 6},
            'sticky': 'NES'
            },
          tree_dates = {
            'type': Treeview,
            'args': {'selectmode': 'browse', 'show': 'tree', 'show': 'headings', 'columns': ('start', 'end')},
            'grid': {'row': 0, 'column': 1, 'columnspan': 5},
            'sticky': 'NWES'
            },
          scroll_dates = {
            'type': Scrollbar,
            'args': {'orient': 'vertical'},
            'grid': {'row': 0, 'column': 5},
            'sticky': 'NES'
            },
          cal_dates_beg = {
            'type': DateEntry,
            'args': {'date_pattern': 'd.M.y', 'locale': 'pl_PL', 'mindate': MIN_DATE, 'textvariable': self.vars.var.date_beg},
            'grid': {'row': 1, 'column': 1, 'columnspan': 2},
            'tooltip': 'Data początku okresu'
            },
          cal_dates_end = {
            'type': DateEntry,
            'args': {'date_pattern': 'd.M.y', 'locale': 'pl_PL', 'mindate': MIN_DATE, 'textvariable': self.vars.var.date_end},
            'grid': {'row': 2, 'column': 1, 'columnspan': 2},
            'tooltip': 'Data końca okresu'
            },
          btn_dates_add = {
            'type': Button,
            'args': {'text': '\ud83d\udfa4', 'command': self.add_date},
            'grid': {'row': 1, 'column': 3},
            'tooltip': 'Dodaj wskazane daty'
            },
          btn_dates_del = {
            'type': Button,
            'args': {'text': '\u232b', 'command': self.delete_date},
            'grid': {'row': 1, 'column': 5},
            'tooltip': 'Usuń daty'
            },
          blank = {
            'type': Label,
            'args': {'text': '', 'background': 'white'},
            'grid': {'row': 3, 'column': 4}
            },
          entry_points = {
            'type': Entry,
            'args': {'textvariable': self.vars.var.name},
            'grid': {'row': 4, 'column': 1, 'columnspan': 5},
            'sticky': 'NWES'
            },
          btn_points_add_project = {
            'type': Button,
            'args': {'text': '\ud83d\uddbf', 'command': self.add_catalogue},
            'grid': {'row': 5, 'column': 1},
            'tooltip': 'Dodaj projekt'
            },
          btn_points_add_entry = {
            'type': Button,
            'args': {'text': '\ud83d\udfa4', 'command': self.add_item},
            'grid': {'row': 5, 'column': 2},
            'tooltip': 'Dodaj punkt kosztorysu'
            },
          btn_points_change = {
            'type': Button,
            'args': {'text':'\u270e','command': self.change_item},
            'grid': {'row': 5, 'column': 4},
            'tooltip': 'Zmień nazwę elementu'
            },
          btn_points_del = {
            'type': Button,
            'args': {'text':'\u232b','command': self.delete_item},
            'grid': {'row': 5, 'column': 5},
            'tooltip': 'Usuń element'
            },
          txt_field = {
            'type': Text,
            'args': {'wrap':'word'},
            'grid': {'row': 6, 'column': 0, 'columnspan': 6},
            'sticky': 'NWES',
            'borderfull': {'highlightthickness': 1, 'highlightbackground': 'gray'}
            },
          btn_text = {
            'type': Button,
            'args': {'text': '\u166d', 'state': 'disabled', 'command': self.text_save},
            'grid': {'row': 7, 'column': 1},
            'tooltip': 'Zapisz tekst'
            },
          btn_save = {
            'type': Button,
            'args': {'text': 'Zapisz', 'command': self.save_file, 'state': 'disabled'},
            'grid': {'row': 7, 'column': 4, 'columnspan': 2},
            'tooltip': 'Zapisz zmiany w pliku'
            }
          )
        self.grid.row.update({
          0: {'weight': 3, 'minsize': 100},
          6: {'weight': 3, 'minsize': 100}
          })
        self.grid.col.update({
          0: {'weight': 3, 'minsize': 150}
          })
        self.menus.update(
          menu_main = {
            'main': True,
            'elements': [
              ('menu', {'label': 'Plik', 'menu': 'menu_file'}),
              ('command', {'label': 'Formatowanie', 'command': self.show_format}),
              ('command', {'label': 'Pomoc', 'command': self.show_help})
              ]
            },
          menu_file = {
            'elements': [
              ('command', {'label': 'Wybierz\u2026', 'command': self.select_file}),
              ('command', {'label': 'Przeładuj', 'command': self.reload}),
              ('separator', None),
              ('command', {'label': 'Wyjdź', 'command': self.close})
              ]
            }
          )
        self.binds.extend([
          ('tree_points', ('<<TreeviewSelect>>', self.points_select)),
          ('txt_field', ('<Key>', self.text_change)),
          ('txt_field', ('<Control-s>', self.text_save))
          ])

    def post(self) -> None:
        # styles, other settings, and actions #
        self.vars.style.configure('TFrame', background='white')
        self.elem.tree_dates.heading('start', text='Start')
        self.elem.tree_dates.heading('end', text='Koniec')
        self.elem.tree_dates.column('start', width=40, minwidth=40)
        self.elem.tree_dates.column('end', width=40, minwidth=40)
        self.elem.tree_points.configure(yscrollcommand=self.elem.scroll_points.set)
        self.elem.tree_dates.configure(yscrollcommand=self.elem.scroll_dates.set)
        self.elem.scroll_points.configure(command=self.elem.tree_points.yview)
        self.elem.scroll_dates.configure(command=self.elem.tree_dates.yview)

        self.set_data()

    def close(self) -> None:
        if not self.ask_about_changes():
            self.root.destroy()


if __name__ == '__main__':
    DataEditor()
