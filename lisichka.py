import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from tkinter import simpledialog
from tkcalendar import DateEntry
from tkinter import filedialog
import os
from datetime import datetime, timedelta
from tkinter import StringVar
import openpyxl

# Пути к CSV файлам
zapros1 = r'Zapros1.csv'
zapros2 = r'Zapros2.csv'


class CSVViewer:
    def __init__(self, root):
        self.root = root
        self.current_df = None
        self.original_df = None
        self.sort_column = None
        self.sort_order = True
        self.cache = {}
        self.current_file = None
        self.row_count_label = None

        self.social_count = self.get_row_count(zapros1)
        self.thematic_count = self.get_row_count(zapros2)

        self.setup_ui()
        self.modified = False # флаг изменения
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing) #обработчик закрытия
        self.dropdown_columns = {
            'Адресант': ['Физическое лицо', 'Юридическое лицо'],
            'Оплата': ['Платный', 'Бесплатный'],
            'Результат': ['Положительный', 'Отрицательный'],
            'Используемые материалы': ['По документам', 'По учетным данным']
        }

    def get_row_count(self, filepath):
        try:
            if filepath not in self.cache:
                self.cache[filepath] = pd.read_csv(
                    filepath,
                    sep=';',
                    encoding='utf-8-sig',
                    dtype=str,
                    na_filter=False
                )
            return len(self.cache[filepath])
        except Exception as e:
            print(f"Ошибка чтения CSV: {e}")
            return 0

    def setup_ui(self):
        # Настройка основных фреймов
        self.root.title("ЗМЕЙКА")
        self.root.state("zoomed")

        # Фрейм для кнопок выбора файлов
        self.frame_buttons = tk.Frame(self.root)
        self.frame_buttons.pack(fill=tk.X, padx=10, pady=10)
        # Фрейм для кнопки сохранить
        self.frame_buttons1 = tk.Frame(self.root)
        self.frame_buttons1.pack(fill=tk.X, padx=10, pady=10)

        # Фрейм для фильтров
        self.frame_filter = tk.Frame(self.root)
        self.frame_filter.pack(fill=tk.X, padx=10, pady=5)

        # Фрейм для кнопок управления
        self.frame_controls = tk.Frame(self.root)
        self.frame_controls.pack(fill=tk.X, padx=10, pady=5)

        # Фрейм для таблицы
        self.frame_table = tk.Frame(self.root, bd=1, relief=tk.SOLID)
        self.frame_table.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        # Фрейм для кнопок управления2
        self.frame_controls2 = tk.Frame(self.root)
        self.frame_controls2.pack(fill=tk.X, padx=10, pady=5)
        # Кнопки выбора данных
        self.btn1 = tk.Button(
            self.frame_buttons,
            text=f"Социально-правовые ({self.social_count})",
            command=lambda: self.load_csv(zapros1))
        self.btn1.pack(side=tk.LEFT, padx=5)

        self.btn2 = tk.Button(
            self.frame_buttons,
            text=f"Тематические({self.thematic_count})",
            command=lambda: self.load_csv(zapros2))
        self.btn2.pack(side=tk.LEFT, padx=5)
        # Кнопки управления
        self.btn_apply = tk.Button(self.frame_controls,
                                   text="Применить фильтры",
                                   command=self.apply_filters)
        self.btn_apply.pack(side=tk.LEFT, padx=5)

        self.btn_reset = tk.Button(self.frame_controls,
                                   text="Сбросить все фильтры и сортировку",
                                   command=self.reset_all)
        self.btn_reset.pack(side=tk.LEFT, padx=5)
        # сохранить в ексель
        self.btn_save_excel = tk.Button(
            self.frame_controls,
            text="Сохранить в Excel",
            command=self.save_to_excel
        )
        self.btn_save_excel.pack(side=tk.LEFT, padx=5)
        self.row_count_label = tk.Label(
            self.frame_controls,
            text="ИТОГО: 0",
            font=('Arial', 10, 'bold'),
            fg='#666666'
        )
        self.row_count_label.pack(side=tk.RIGHT, padx=10)

        # Настройка стилей таблицы
        self.configure_styles()
        # Кнопка добавления записи
        self.btn_add = tk.Button(self.frame_controls2,
                                 text="Новый запрос",
                                 command=self.add_new_row)
        self.btn_add.pack(side=tk.LEFT, padx=5)
        #удаление
        self.btn_delete = tk.Button(self.frame_controls2,
                                    text="Удалить запрос",
                                    command=self.delete_selected_row)
        self.btn_delete.pack(side=tk.LEFT, padx=5)

        # Кнопка формирования отчета
        self.btn_report = tk.Button(self.frame_controls2,
                                    text="Сформировать отчет",
                                    command=self.generate_report)
        self.btn_report.pack(side=tk.LEFT, padx=5)

        # Панель фильтров
        self.filter_frames = {}
        self.filter_types = {}
        self.filter_values = {}

        for i in range(15):
            frame = tk.Frame(self.frame_filter)
            frame.grid(row=0, column=i, padx=2, sticky='nsew')
            self.filter_frames[i] = frame

            # Выпадающий список типа фильтра
            type_var = StringVar()
            cb = ttk.Combobox(frame,
                              values=['содержит','не содержит', 'точное равенство', 'заполнено', 'не заполнено'],
                              textvariable=type_var,
                              state='readonly',
                              width=12)
            cb.set('содержит')
            cb.grid(row=0, column=0, sticky='ew')
            self.filter_types[i] = type_var

            # Поле значения фильтра
            entry = tk.Entry(frame, width=15)
            entry.grid(row=1, column=0, sticky='ew')
            self.filter_values[i] = entry

            # Привязка событий
            cb.bind('<<ComboboxSelected>>', lambda e, i=i: self.update_filter_entry_state(i))
            entry.bind('<Return>', self.apply_filters)



        # Кнопка сохранения
        self.btn_save = tk.Button(self.frame_buttons1,
                                 text="Сохранить изменения",
                                 command=self.save_changes,
                                 # bg="#4CAF50",
                                 fg="black",
                                 font=('Arial', 11, 'bold'),
                                 padx=14
                                  )
        self.btn_save.pack(side=tk.LEFT, padx=5)


        # Настройка стилей таблицы
        self.configure_styles()

    def configure_styles(self):
        style = ttk.Style()
        style.theme_use("clam")

        style.configure("Treeview",
                        background="white",
                        foreground="black",
                        rowheight=25,
                        fieldbackground="white",
                        bordercolor="gray",
                        borderwidth=1,
                        font=('Arial', 10))


        style.configure("Treeview.Heading",
                        background="#4B8BBE",
                        foreground="white",
                        padding=5,
                        font=('Arial', 10, 'bold'))

        style.map('Treeview',
                  background=[('selected', '#B3D9FF')],
                  foreground=[('selected', 'black')])

    def load_csv(self, filepath):
        try:
            # Очистка предыдущих данных
            for widget in self.frame_table.winfo_children():
                widget.destroy()

            # Загрузка данных с кэшированием
            if filepath not in self.cache:
                self.cache[filepath] = pd.read_csv(
                    filepath,
                    sep=';',
                    encoding='utf-8-sig',
                    dtype=str,
                    na_filter=False,

                )

            self.current_file = filepath
            self.original_df = self.cache[filepath]  # Ссылка на кэш
            self.current_df = self.original_df.copy()
            self.sort_column = None

            self.update_filter_labels()
            self.create_treeview()
            self.update_treeview()
            self.reset_filters()
            self.update_row_count()


        except Exception as e:
            tk.Label(self.frame_table, text=f"Ошибка: {e}", fg="red").pack()

    def update_filter_labels(self):
        columns = list(self.current_df.columns)
        for i in range(15):
            frame = self.filter_frames[i]
            # Очищаем предыдущие метки
            for widget in frame.winfo_children():
                if isinstance(widget, tk.Label) and widget.grid_info()["row"] == 2:
                    widget.destroy()

            if i < len(columns):
                # Добавляем новую метку с названием столбца
                label_text = columns[i][:15]
                tk.Label(frame, text=label_text).grid(row=2, column=0, sticky='ew')
                frame.grid()
            else:
                frame.grid_remove()

    def create_treeview(self):
        self.tree = ttk.Treeview(self.frame_table, show='headings')
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        vsb = ttk.Scrollbar(self.frame_table, orient="vertical", command=self.tree.yview)
        vsb.pack(side='right', fill='y')
        self.tree.configure(yscrollcommand=vsb.set)

        # Берем только первые 15 столбцов
        columns_to_display = list(self.current_df.columns[:15])
        self.tree["columns"] = columns_to_display

        for col in columns_to_display:
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_by_column(c))
            self.tree.column(col, width=120, anchor='w')

        # Привязка двойного клика для редактирования
        self.tree.bind('<Double-1>', self.edit_cell)

    def sort_by_column(self, col):
        if self.sort_column == col:
            self.sort_order = not self.sort_order
        else:
            self.sort_column = col
            self.sort_order = True

        self.current_df = self.current_df.sort_values(
            by=col,
            ascending=self.sort_order,
            key=lambda x: x.str.lower() if x.dtype == 'object' else x
        )
        self.update_treeview()

    def update_filter_entry_state(self, col_idx):
        """Обновление состояния поля ввода в зависимости от типа фильтра"""
        filter_type = self.filter_types[col_idx].get()
        entry = self.filter_values[col_idx]

        if filter_type in ['заполнено', 'не заполнено']:
            entry.config(state='disabled')
            entry.delete(0, tk.END)
        else:
            entry.config(state='normal')

    def apply_filters(self, event=None):
        filtered_df = self.original_df.copy()

        for col_idx in range(15):
            if col_idx >= len(self.original_df.columns):
                continue

            col_name = self.original_df.columns[col_idx]
            filter_type = self.filter_types[col_idx].get()
            filter_value = self.filter_values[col_idx].get().lower()

            if filter_type == 'содержит' and filter_value:
                filtered_df = filtered_df[filtered_df[col_name].str.lower().str.contains(filter_value)]
            elif filter_type == 'не содержит' and filter_value:
                filtered_df = filtered_df[~filtered_df[col_name].str.lower().str.contains(filter_value.lower())]
            elif filter_type == 'точное равенство' and filter_value:
                filtered_df = filtered_df[filtered_df[col_name].str.lower() == filter_value]
            elif filter_type == 'заполнено':
                filtered_df = filtered_df[filtered_df[col_name].astype(bool)]
            elif filter_type == 'не заполнено':
                filtered_df = filtered_df[~filtered_df[col_name].astype(bool)]

        self.current_df = filtered_df
        self.sort_column = None
        self.update_treeview()
        self.update_row_count()

    def update_filter_labels(self):
        columns = list(self.current_df.columns)
        for i in range(15):
            frame = self.filter_frames[i]
            if i < len(columns):
                # Обновляем текст над фильтром
                label_text = columns[i][:15]
                tk.Label(frame, text=label_text).grid(row=2, column=0, sticky='ew')
                frame.grid()
            else:
                frame.grid_remove()

    def update_treeview(self):
        self.tree.delete(*self.tree.get_children())
        for idx, row in self.current_df.iterrows():
            tag = 'evenrow' if idx % 2 == 0 else 'oddrow'
            self.tree.insert('', 'end', iid=str(idx), values=list(row), tags=(tag,))

    def reset_all(self):
        self.current_df = self.original_df.copy()
        self.sort_column = None
        self.sort_order = True
        self.reset_filters()
        self.update_treeview()
        self.update_row_count()

    def reset_filters(self):
        # Сброс значений фильтров и их типов
        for i in range(15):
            self.filter_types[i].set('содержит')
            self.filter_values[i].delete(0, tk.END)
            self.filter_values[i].config(state='normal')

    def edit_cell(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != 'cell':
            return

        column = self.tree.identify_column(event.x)
        item = self.tree.identify_row(event.y)
        col_index = int(column[1:]) - 1

        if 0 <= col_index < len(self.tree["columns"]):
            col_name = self.tree["columns"][col_index]
            current_values = self.tree.item(item, 'values')
            current_value = current_values[col_index]
            x, y, width, height = self.tree.bbox(item, column)

            if col_name in self.dropdown_columns:
                self.create_dropdown(col_name, x, y, width, height, item, col_index, current_value)
            else:
                self.create_entry(x, y, width, height, item, col_index, current_value)

    def create_dropdown(self, col_name, x, y, width, height, item, col_index, current_value):
        values = self.dropdown_columns[col_name]
        default = current_value if current_value in values else values[0]

        # Убираем state='readonly' чтобы разрешить ручной ввод
        cb = ttk.Combobox(self.frame_table, values=values)
        cb.set(default)
        cb.place(x=x, y=y, width=width, height=height)
        cb.focus_set()

        def save_edit(event=None):
            new_value = cb.get()
            self.update_cell_value(item, col_index, new_value)
            cb.destroy()

        cb.bind('<<ComboboxSelected>>', save_edit)
        cb.bind('<FocusOut>', save_edit)

    def create_entry(self, x, y, width, height, item, col_index, current_value):
        entry = tk.Entry(self.frame_table)
        entry.place(x=x, y=y, width=width, height=height)
        entry.insert(0, current_value)
        entry.focus_set()

        def save_edit(event=None):
            new_value = entry.get()
            self.update_cell_value(item, col_index, new_value)
            entry.destroy()

        entry.bind('<Return>', save_edit)
        entry.bind('<FocusOut>', save_edit)

    def update_cell_value(self, item, col_index, new_value):
        col_name = self.tree["columns"][col_index]
        current_values = list(self.tree.item(item, 'values'))
        current_values[col_index] = new_value
        self.tree.item(item, values=current_values)

        idx = int(item)
        self.original_df.at[idx, col_name] = new_value
        if idx in self.current_df.index:
            self.current_df.at[idx, col_name] = new_value
        self.modified = True

    def add_new_row(self):
        if self.original_df is None:
            messagebox.showwarning("Ошибка", "Сначала загрузите файл")
            return

        # Создаем диалоговое окно
        dialog = tk.Toplevel(self.root)
        dialog.title("Новый запрос")
        dialog.geometry("800x800")  # Увеличенный размер окна
        dialog.grab_set()

        # Определяем тип запроса
        current_type = "Социально-правовой" if self.current_file == zapros1 else "Тематический"
        type_label = tk.Label(dialog,
                              text=f"Тип запроса: {current_type}",
                              font=('Arial', 14, 'bold'))
        type_label.grid(row=0, columnspan=2, pady=10, sticky='n')

        # Контейнер для полей с прокруткой
        container = tk.Frame(dialog)
        container.grid(row=1, columnspan=2, sticky='nsew')

        canvas = tk.Canvas(container)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Получаем текущую дату минус 2 дня
        today = datetime.today()
        default_post_date = today - timedelta(days=2)

        # Поля для ввода (только первые 15 столбцов)
        fields = [
            ('Регистрационный номер:', '04-', 'entry'),
            ('Адресат (Физ./Юрид.) лицо:', ['', 'Физическое лицо', 'Юридическое лицо'], 'combobox'),
            ('Фамилия:', '', 'entry'),
            ('Имя:', '', 'entry'),
            ('Отчество:', '', 'entry'),
            ('Адрес заявителя:', '', 'text'),
            ('Дата поступления запроса:', default_post_date, 'date'),  # Устанавливаем дату по умолчанию
            ('Вид запроса:', ['', 'Зарплата', 'Стаж', 'Обучение', 'Переименование', 'Декретный отпуск', 'Награждение', 'Репрессии', 'Эвакуация', 'Решение', 'Договор', 'Наличие документов', 'Захоронение', 'Актовая запись', 'Копии документов', 'Занимаемая должность', 'Количество отработанных дней/часов'], 'combobox'),
            ('Содержание запроса:', '', 'text'),
            ('Результат запроса:', ['', 'Положительный', 'Отрицательный'], 'combobox'),
            ('Исполнитель:', '', 'entry'),
            ('Дата исполнения запроса:', '', 'date'),  # Пустое поле
            ('Используемые материалы:', ['', 'По документам', 'По учетным данным'], 'combobox'),
            ('Оплата за исполнение:', ['', 'Платный', 'Бесплатный'], 'combobox'),
            ('Поступление:', ['', 'По эл. почте', 'VipNET', 'Обращение лично', 'Почтовая связь', 'Другое'], 'combobox'),
        ]

        entries = {}

        for i, (label, default, field_type) in enumerate(fields):
            tk.Label(scrollable_frame,
                     text=label,
                     font=('Arial', 12)
                     ).grid(row=i, column=0, padx=10, pady=5, sticky='e')

            if field_type == 'combobox':
                var = tk.StringVar(value=default[0] if default else '')
                cb = ttk.Combobox(scrollable_frame,
                                  textvariable=var,
                                  values=default,
                                  width=40,
                                  font=('Arial', 12))
                cb.grid(row=i, column=1, padx=10, pady=5, sticky='ew')
                entries[label] = var

            elif field_type == 'date':
                entry = DateEntry(scrollable_frame,
                                  date_pattern='dd.mm.yyyy',
                                  width=12,
                                  font=('Arial', 12),
                                  allow_none=True)  # Разрешаем пустое значение
                if label == 'Дата исполнения запроса:':
                    entry.delete(0, 'end')  # Очищаем поле при создании
                elif isinstance(default, datetime):
                    entry.set_date(default)
                entry.grid(row=i, column=1, padx=10, pady=5, sticky='w')
                entries[label] = entry
            elif field_type == 'text':
                text = tk.Text(scrollable_frame,
                               height=4,
                               width=50,
                               font=('Arial', 12),
                               wrap=tk.WORD)
                text.grid(row=i, column=1, padx=10, pady=5, sticky='ew')
                entries[label] = text
            else:
                var = tk.StringVar(value=default)
                entry = tk.Entry(scrollable_frame,
                                 textvariable=var,
                                 width=50,
                                 font=('Arial', 12))
                entry.grid(row=i, column=1, padx=10, pady=5, sticky='ew')
                entries[label] = var

        # Кнопки сохранения/отмены
        def save():
            new_data = {
                'Регистрационный №': entries['Регистрационный номер:'].get(),
                'Адресат': entries['Адресат (Физ./Юрид.) лицо:'].get(),
                'Фамилия': entries['Фамилия:'].get(),
                'Имя': entries['Имя:'].get(),
                'Отчество': entries['Отчество:'].get(),
                'Адрес заявителя': entries['Адрес заявителя:'].get("1.0", "end-1c"),
                # Исправленный формат даты
                'Дата запроса': entries['Дата поступления запроса:'].get_date().strftime('%Y-%m-%d'),
                'Характер запроса': entries['Вид запроса:'].get(),
                'Содержание запроса': entries['Содержание запроса:'].get("1.0", "end-1c"),
                'Результат': entries['Результат запроса:'].get(),
                'Исполнитель': entries['Исполнитель:'].get(),
                # Исправленный формат даты
                'Дата исполнения': entries['Дата исполнения запроса:'].get_date().strftime('%Y-%m-%d') if entries['Дата исполнения запроса:'].get_date() else '',
                'Используемые материалы': entries['Используемые материалы:'].get(),
                'Оплата': entries['Оплата за исполнение:'].get(),
                'Поступление': entries['Поступление:'].get()  # 15-й столбец
            }

            self.original_df = pd.concat([self.original_df, pd.DataFrame([new_data])], ignore_index=True)
            self.cache[self.current_file] = self.original_df
            self.reset_all()
            self.modified = True
            dialog.destroy()

        btn_frame = tk.Frame(dialog)
        btn_frame.grid(row=2, columnspan=2, pady=10)

        tk.Button(btn_frame,
                  text="Сохранить",
                  command=save,
                  font=('Arial', 12),
                  width=15).pack(side=tk.LEFT, padx=20)

        tk.Button(btn_frame,
                  text="Отмена",
                  command=dialog.destroy,
                  font=('Arial', 12),
                  width=15).pack(side=tk.LEFT, padx=20)

        # Настройка весов строк и столбцов для адаптивности
        dialog.rowconfigure(1, weight=1)
        dialog.columnconfigure(0, weight=1)
        dialog.columnconfigure(1, weight=1)

    def edit_new_row(self, row_index):
        item_id = str(row_index)

        # Проверяем наличие элемента в дереве
        if not self.tree.exists(item_id):
            return

        # Даем дополнительное время на рендеринг
        self.tree.see(item_id)
        self.root.after(50, lambda: self.focus_and_edit(item_id))

    def focus_and_edit(self, item_id):
        self.tree.focus(item_id)
        self.tree.selection_set(item_id)

        # Получаем координаты первой колонки
        column = '#1'
        bbox = self.tree.bbox(item_id, column)

        if bbox:  # Проверяем что координаты получены
            self.edit_cell_programmatically(item_id, column)

    def edit_cell_programmatically(self, item_id, column):
        # Имитируем клик для редактирования
        x, y, width, height = self.tree.bbox(item_id, column)

        class Event:
            def __init__(self, x, y):
                self.x = x
                self.y = y

        event = Event(x + width // 2, y + height // 2)
        self.edit_cell(event)

    def delete_selected_row(self):
        if not self.original_df.empty:
            selected_items = self.tree.selection()
            if not selected_items:
                messagebox.showwarning("Ошибка", "Выберите строку для удаления")
                return

            # Удаление из DataFrame
            for item in selected_items:
                # Получаем индекс строки в оригинальном DataFrame
                idx = int(item)

                # Удаляем из оригинального DataFrame
                self.original_df = self.original_df.drop(index=idx)

                # Обновляем кэш
                self.cache[self.current_file] = self.original_df

                # Если строка есть в текущем DataFrame - удаляем и оттуда
                if idx in self.current_df.index:
                    self.current_df = self.current_df.drop(index=idx)

            # Сбрасываем индексы после удаления
            self.original_df.reset_index(drop=True, inplace=True)
            self.current_df.reset_index(drop=True, inplace=True)

            # Обновляем отображение
            self.update_treeview()
            self.modified = True

    def generate_report(self):
        if self.original_df is None:
            messagebox.showwarning("Ошибка", "Сначала загрузите файл с данными")
            return

        date_dialog = tk.Toplevel(self.root)
        date_dialog.title("Выберите период")

        # Устанавливаем даты по умолчанию: первый день текущего месяца и текущая дата
        today = datetime.today()
        first_day = today.replace(day=1)

        tk.Label(date_dialog, text="Дата начала:").grid(row=0, column=0, padx=5, pady=5)
        start_entry = DateEntry(date_dialog, date_pattern='dd.MM.yyyy', year=first_day.year,
                                month=first_day.month, day=first_day.day)
        start_entry.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(date_dialog, text="Дата окончания:").grid(row=1, column=0, padx=5, pady=5)
        end_entry = DateEntry(date_dialog, date_pattern='dd.MM.yyyy', year=today.year,
                              month=today.month, day=today.day)
        end_entry.grid(row=1, column=1, padx=5, pady=5)

        def create_report():
            try:
                startdate = pd.to_datetime(start_entry.get_date())
                enddate = pd.to_datetime(end_entry.get_date())

                if startdate > enddate:
                    messagebox.showerror("Ошибка", "Дата начала не может быть позже даты окончания")
                    return

                df = self.original_df.copy()
                df['Дата запроса'] = pd.to_datetime(df['Дата запроса'], errors='coerce')
                df['Дата исполнения'] = pd.to_datetime(df['Дата исполнения'], errors='coerce')

                startdate = pd.to_datetime(startdate)
                enddate = pd.to_datetime(enddate)

                df_p = df[(df['Дата запроса'] >= startdate) & (df['Дата запроса'] <= enddate)]
                num_p = len(df_p)

                df_e = df[(df['Дата исполнения'] >= startdate) & (df['Дата исполнения'] <= enddate)]
                num_e = len(df_e)

                df_doc = df_e[(df_e['Используемые материалы'] == 'По документам')]
                df_uch = df_e[(df_e['Используемые материалы'] == 'По учетным данным')]
                num_doc = len(df_doc)
                num_uch = len(df_uch)

                report = pd.DataFrame({
                    'Период': [f'{startdate.date()} — {enddate.date()}'],
                    'Поступило': [num_p],
                    'Исполнено': [num_e],
                    'Исполнено по документам': [num_doc],
                    'Исполнено по учетным данным': [num_uch]
                })

                # Формируем имя файла с датой и периодом
                today_str = datetime.today().strftime("%Y-%m-%d")
                period_str = f"{startdate.strftime('%d_%m_%Y')}-{enddate.strftime('%d_%m_%Y')}"
                base_name = f"Отчет_{today_str}_{period_str}"
                extension = ".xlsx"
                counter = 0

                # Проверяем существование файлов и добавляем счетчик при необходимости
                while True:
                    report_file = f"{base_name}{f' ({counter})' if counter else ''}{extension}"
                    if not os.path.exists(report_file):
                        break
                    counter += 1

                report.to_excel(report_file, index=False)

                # Пытаемся открыть отчет
                try:
                    os.startfile(report_file)
                    messagebox.showinfo("Успех", f"Отчет успешно создан:\n{report_file}")
                except Exception as e:
                    messagebox.showinfo("Успех",
                                        f"Отчет создан, но не удалось открыть автоматически:\n{report_file}\n\nОшибка: {str(e)}")

                date_dialog.destroy()

            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка при создании отчета: {str(e)}")

        tk.Button(date_dialog, text="Создать отчет", command=create_report).grid(row=2, columnspan=2, pady=10)

    def update_row_count(self):
        """Обновление метки с количеством строк"""
        if self.current_df is not None:
            count = len(self.current_df)
            self.row_count_label.config(
                text=f"ИТОГО: {count:,}".replace(',', ' ')
            )
        else:
            self.row_count_label.config(text="ИТОГО: 0")

    def save_to_excel(self):
        """Сохранение текущей таблицы в Excel"""
        if self.current_df is None or self.current_df.empty:
            messagebox.showwarning("Ошибка", "Нет данных для сохранения")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel файлы", "*.xlsx")],
            title="Сохранить как Excel файл"
        )

        if not file_path:
            return

        try:
            # Сохраняем текущий DataFrame
            self.current_df.to_excel(file_path, index=False, engine='openpyxl')

            # Открываем файл в проводнике
            os.startfile(file_path)
            messagebox.showinfo("Успех", f"Файл сохранен:\n{file_path}")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл:\n{str(e)}")

    def save_changes(self):
        if self.current_file is None:
            messagebox.showwarning("Ошибка", "Нет открытого файла для сохранения.")
            return
        try:
            self.cache[self.current_file].to_csv(
                self.current_file,
                sep=';',
                index=False,
                encoding='utf-8-sig'
            )
            messagebox.showinfo("Успех", "Изменения сохранены успешно!")
            self.modified = False  # Сбрасываем флаг после сохранения
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл: {e}")

    def on_closing(self):
        if self.modified:
            response = messagebox.askyesnocancel(
                "Сохранение изменений",
                "Сохранить изменения?",
                icon='question',
                default=messagebox.YES
            )
            if response is None:  # Отмена
                return
            elif response:  # Да
                try:
                    self.save_changes()
                except Exception as e:
                    messagebox.showerror("Ошибка", f"Не удалось сохранить: {str(e)}")
                    return
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = CSVViewer(root)
    root.mainloop()


