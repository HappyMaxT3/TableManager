import tkinter as tki
import tkinter.ttk as ttk
from tkinter import filedialog as fld
from tkinter import messagebox
from tkinter import simpledialog
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import io
from PIL import Image, ImageTk
from tkinter import filedialog
import os
import _datetime as time


# Получение пути к папке, в которой находится скрипт
script_dir = os.path.dirname(os.path.abspath(__file__))
# Получение пути на папку выше
parent_dir = os.path.dirname(script_dir)

def update_dataframe(event):
    global GDS
    i = event.widget.grid_info()['row'] - 1
    j = event.widget.grid_info()['column']
    new_value = event.widget.get(
    )  # Assuming this is the new value from the widget

    # Get the column name for proper type conversion
    column_name = GDS.columns[j]

    # Determine the correct type and cast the new value accordingly
    if GDS[column_name].dtype == 'int64':
        new_value = int(new_value)
    elif GDS[column_name].dtype == 'float64':
        new_value = float(new_value)
    elif GDS[column_name].dtype == 'object':
        new_value = str(new_value)
    # Add more conditions if you have other data types

    # Update the DataFrame with the new value
    GDS.at[i, column_name] = new_value


def save_table(GDS, file_data):
    for file in file_data:
        new_df = GDS[file_data[file]]
        new_df.drop_duplicates(inplace=True)

        # Убедитесь, что путь к файлу корректен
        data_dir = os.path.abspath(
            os.path.join(os.path.dirname(__file__), '../data'))

        if not os.path.exists(data_dir):
            os.makedirs(data_dir)

        if file.endswith('.csv'):
            new_df.to_csv(os.path.join(data_dir, file), index=False)
        else:
            new_df.to_pickle(os.path.join(data_dir, file), index=False)


def save_table_as_csv(GDS, file_data):
    for file in file_data:
        new_df = GDS[file_data[file]]
        new_df.drop_duplicates(inplace=True)

        # Открытие диалога сохранения файла
        ftypes = [('Text файлы', '*.csv'), ('Все файлы', '*')]
        dlg = fld.asksaveasfilename(filetypes=ftypes, defaultextension=".csv")

        # Проверка, что пользователь выбрал файл для сохранения
        if dlg:
            try:
                new_df.to_csv(dlg, index=False)
                messagebox.showinfo("Success", "File saved successfully")
            except Exception as e:
                messagebox.showerror(
                    "Error", f"An error occurred while saving the file: {e}")


def read_data():
    """
    Чтение и демонстрация данных из CSV или XLSX файла
    """
    global GDS, height, width, top, vrs, pnt, canvas, table_frame, file_data, report

    # Создание диалогового окна для выбора файла
    ftypes = [('CSV файлы', '*.csv'), ('Pickle файлы', '*.pkl'),
              ('Все файлы', '*')]
    fl = fld.askopenfilename(filetypes=ftypes)
    if not fl:
        return

    # Определение типа файла и чтение данных
    if fl.endswith('.csv'):
        new_GDS = pd.read_csv(fl)
    elif fl.endswith('.pkl'):
        new_GDS = pd.read_pickle(fl)
    else:
        messagebox.showerror("ERROR", "Incorrect file format")
        return

    file_name = os.path.basename(fl)
    file_data[file_name] = new_GDS.columns.tolist()

    # Объединение с предыдущей таблицей
    if GDS.empty:
        GDS = new_GDS
    else:
        common_keys = set(GDS.keys()).intersection(set(new_GDS.keys()))
        if common_keys:
            key = list(common_keys)[0]
            GDS = pd.merge(GDS, new_GDS, on=key, how='outer')
        else:
            messagebox.showerror("ERROR", "No common columns to merge on.")

    # Определение размеров таблицы
    height = GDS.shape[0]
    width = GDS.shape[1]

    # Очистка предыдущего содержимого
    clear_for_new()

    # Массив указателей на текстовые буферы для передачи данных
    vrs = np.empty(shape=(height, width), dtype="O")

    # Инициализация указателей на буферы
    for i in range(height):
        for j in range(width):
            vrs[i, j] = tki.StringVar()

    # Массив указателей на виджеты Entry
    pnt = np.empty(shape=(height, width), dtype="O")

    # Минимальная ширина столбцов в пикселях
    min_width = 10

    # Инициализация указателей на виджеты Entry
    for i in range(height):
        for j in range(width):
            pnt[i, j] = tki.Entry(table_frame,
                                  textvariable=vrs[i, j],
                                  font=('Arial', 12),
                                  bg="#121212",
                                  fg="#FFFFFF",
                                  justify="center")
            pnt[i, j].bind("<FocusOut>", update_dataframe)
            pnt[i, j].grid(row=i + 1, column=j, padx=1, pady=1)

    # Вывод заголовков столбцов
    for j in range(width):
        header_label = tki.Label(table_frame,
                                 text=GDS.columns[j],
                                 font=('Arial', 12, 'bold'),
                                 bg="#121212",
                                 fg="#FFFFFF",
                                 justify="center")
        header_label.grid(row=0, column=j, padx=1, pady=1)

    # Заполнение таблицы значениями
    for i in range(height):
        for j in range(width):
            cnt = GDS.iloc[i, j]
            vrs[i, j].set(str(cnt))

    # Настройка минимальной ширины столбцов
    for j in range(width):
        table_frame.grid_columnconfigure(j, minsize=min_width)

    # Установка размеров canvas для скроллирования
    canvas.update_idletasks()
    canvas.config(scrollregion=canvas.bbox("all"))


def rgb_hack(rgb):
    """
    Tcl/Tk не понимает rgb формат
    Использование RGB формата для цветов
    https://docs.python.org/3/library/stdtypes.html#old-string-formatting
    rgb - кортеж (r, g, b)
    """
    return "#%02x%02x%02x" % rgb


def clear_for_new():
    """
    Очистка фрейма top
    """
    global GDS
    for widgets in table_frame.winfo_children():
        widgets.destroy()


def clear_top():
    """
    Очистка фрейма top
    """
    global GDS
    for widgets in table_frame.winfo_children():
        widgets.destroy()
    GDS = pd.DataFrame([])


def clear_right():
    """
    Очистка фрейма right
    """
    global GDS
    for widgets in chart_frame.winfo_children():
        widgets.destroy()


def make_report(data_base, displayed_columns, column_name, column_value):
    """
    Функция для создания отчета на основе переданных параметров
    """
    column_values = list(map(int, column_value))
    index = data_base[column_name].isin(column_values)
    report = data_base.loc[index, displayed_columns]
    return report


def generate_pivot_table():
    """
    Функция для генерации отчета на основе пользовательского ввода
    """

    def submit_report():
        pivot_index = key_for_index.get()
        pivot_columns = key_for_columns.get()
        pivot_values = value.get()
        action = action_for_pivot.get()

        if pivot_index not in GDS.columns:
            messagebox.showerror("ERROR", "Incorrect column name")
            return

        if pivot_columns not in GDS.columns:
            messagebox.showerror("ERROR", "Incorrect column name")
            return

        if pivot_values not in GDS.columns:
            messagebox.showerror("ERROR", "Incorrect column name")
            return

        GDS[pivot_values] = GDS[pivot_values].astype(int)

        if action == "Print values":
            report = GDS.pivot_table(index=pivot_index,
                                                columns=pivot_columns,
                                                values=pivot_values,
                                                fill_value=0)
        elif action == "Add up":
            report = GDS.pivot_table(index=pivot_index,
                                                columns=pivot_columns,
                                                values=pivot_values,
                                                fill_value=0,
                                                aggfunc="sum")
        elif action == "Count quantity":
            report = GDS.pivot_table(index=pivot_index,
                                                columns=pivot_columns,
                                                values=pivot_values,
                                                fill_value=0,
                                                aggfunc="count")
        elif action == "find average":
            report = GDS.pivot_table(index=pivot_index,
                                                columns=pivot_columns,
                                                values=pivot_values,
                                                fill_value=0,
                                                aggfunc="mean")
        elif action == "find the number of unic values":
            report = GDS.pivot_table(index=pivot_index,
                                                columns=pivot_columns,
                                                values=pivot_values,
                                                fill_value=0,
                                                aggfunc="nunique")

        # Очищаем chart_frame
        for widget in chart_frame.winfo_children():
            widget.destroy()

        # Вставляем отчет в chart_frame
        report_text = tki.Text(chart_frame,
                               wrap="none",
                               bg="#121212",
                               fg="#FFFFFF",
                               font=('Arial', 12))
        report_text.pack(fill="both", expand=True)

        report_text.insert("1.0", report.to_string(index=False))

    # Создание окна для ввода данных
    input_window = tki.Toplevel(root)
    input_window.title("Generate Pivot Table")

    key_for_index = tki.StringVar(input_window)
    key_for_index.set("Choose key for index:")
    combobox = ttk.Combobox(input_window,
                            textvariable=key_for_index,
                            font=('Arial', 12))
    combobox['values'] = tuple(GDS.columns)
    combobox.grid(row=0, column=1, padx=10, pady=10)

    key_for_columns = tki.StringVar(input_window)
    key_for_columns.set("Choose key for columns:")
    combobox = ttk.Combobox(input_window,
                            textvariable=key_for_columns,
                            font=('Arial', 12))
    combobox['values'] = tuple(GDS.columns)
    combobox.grid(row=1, column=1, padx=10, pady=10)

    value = tki.StringVar(input_window)
    value.set("Choose value:")
    combobox = ttk.Combobox(input_window,
                            textvariable=value,
                            font=('Arial', 12))
    combobox['values'] = tuple(GDS.columns)
    combobox.grid(row=2, column=1, padx=10, pady=10)

    action_for_pivot = tki.StringVar(input_window)
    action_for_pivot.set("Choose action:")
    combobox = ttk.Combobox(input_window,
                            textvariable=action_for_pivot,
                            font=('Arial', 12))
    combobox['values'] = ("Print values", "Add up", "Count quantity", "find average", "find the number of unic values")
    combobox.grid(row=3, column=1, padx=10, pady=10)

    tki.Button(input_window,
               text="Submit",
               command=submit_report,
               font=('Arial', 12)).grid(row=4, column=0, columnspan=2, pady=20)

    input_window.transient(root)
    input_window.grab_set()


def generate_report():
    """
    Функция для генерации отчета на основе пользовательского ввода
    """

    def submit_report():
        column_name = column_name_var.get()
        column_value = column_value_var.get().split(", ")
        displayed_columns = displayed_columns_var.get().split(", ")

        if column_name not in GDS.columns:
            messagebox.showerror("ERROR", "Incorrect column name")
            return

        new_displayed_columns = list(
            filter(lambda x: x in GDS.columns, displayed_columns))
        if displayed_columns != new_displayed_columns:
            messagebox.showerror("ERROR", "Incorrect column names for report")
            return

        report = make_report(GDS, displayed_columns, column_name, column_value)
        report.drop_duplicates(inplace=True)

        # Очищаем chart_frame
        for widget in chart_frame.winfo_children():
            widget.destroy()

        # Вставляем отчет в chart_frame
        report_text = tki.Text(chart_frame,
                               wrap="none",
                               bg="#121212",
                               fg="#FFFFFF",
                               font=('Arial', 12))
        report_text.pack(fill="both", expand=True)

        report_text.insert("1.0", report.to_string(index=False))

    # Создание окна для ввода данных
    input_window = tki.Toplevel(root)
    input_window.title("Generate Report")

    column_name_var = tki.StringVar()
    column_value_var = tki.StringVar()
    displayed_columns_var = tki.StringVar()

    tki.Label(input_window, text="Column name:",
              font=('Arial', 12)).grid(row=0, column=0, padx=10, pady=10)
    tki.Entry(input_window, textvariable=column_name_var,
              font=('Arial', 12)).grid(row=0, column=1, padx=10, pady=10)

    tki.Label(input_window,
              text="Column values (comma separated):",
              font=('Arial', 12)).grid(row=1, column=0, padx=10, pady=10)
    tki.Entry(input_window, textvariable=column_value_var,
              font=('Arial', 12)).grid(row=1, column=1, padx=10, pady=10)

    tki.Label(input_window,
              text="Displayed columns (comma separated):",
              font=('Arial', 12)).grid(row=2, column=0, padx=10, pady=10)
    tki.Entry(input_window,
              textvariable=displayed_columns_var,
              font=('Arial', 12)).grid(row=2, column=1, padx=10, pady=10)

    tki.Button(input_window,
               text="Submit",
               command=submit_report,
               font=('Arial', 12)).grid(row=3, column=0, columnspan=2, pady=20)

    input_window.transient(root)
    input_window.grab_set()


def save_text_report(report_text):
    """
    Функция для сохранения текстового отчета в файл
    Parameters
    ----------
    report_text : str - текст отчета.
    """
    ftypes = [('Text файлы', '*.txt'), ('Все файлы', '*')]
    dlg = fld.asksaveasfilename(filetypes=ftypes, defaultextension=".txt")
    if dlg:
        with open(dlg, 'w') as f:
            f.write(report_text)
        messagebox.showinfo("Success", "successful saving")


def save_chart_image():
    # Установка пути к Ghostscript
    os.environ["GS"] = "../opt/homebrew/bin/gs"

    # Проверка пути к Ghostscript
    if not os.path.isfile(os.environ["GS"]):
        print("error: way to Ghostscript is wrong!")
        return

    # Создаем изображение в формате PostScript из канваса
    chart_postscript = chart_canvas.postscript(colormode="color")

    # Создаем изображение из формата PostScript
    chart_image = Image.open(io.BytesIO(chart_postscript.encode("utf-8")))

    # Открываем диалоговое окно для сохранения файла
    file_path = filedialog.asksaveasfilename(defaultextension=".png",
                                             filetypes=[("PNG files", "*.png"),
                                                        ("All files", "*.*")])

    if file_path:
        # Сохраняем изображение графика в формате PNG
        chart_image.save(file_path, format="PNG")


def create_bar_chart():

    def plot_chart():
        x_column = x_entry.get()
        y_column = y_entry.get()

        if x_column not in GDS.columns or y_column not in GDS.columns:
            messagebox.showerror(
                "ERROR",
                f"One or both of the columns '{x_column}' and '{y_column}' do not exist."
            )
            return

        # Преобразуем столбцы в числовой формат, если это необходимо
        x_data = pd.to_numeric(GDS[x_column], errors='coerce')
        y_data = pd.to_numeric(GDS[y_column], errors='coerce')

        # Проверяем, что столбцы содержат числовые данные
        if x_data.isna().all() or y_data.isna().all():
            messagebox.showerror("ERROR",
                                 "No numeric data in the specified columns.")
            return

        # Строим график
        fig, ax = plt.subplots()
        GDS.plot(x=x_column, y=y_column, kind='bar', ax=ax)

        # Уменьшаем размер изображения до подходящего
        fig.set_size_inches(3.5, 3.5)

        # Очищаем фрейм с графиком и вставляем новый график
        for widget in chart_frame.winfo_children():
            widget.destroy()

        # Экспортируем изображение графика в канвас
        chart_canvas = FigureCanvasTkAgg(fig, master=chart_frame)
        chart_canvas.draw()
        chart_canvas.get_tk_widget().pack(fill='both', expand=True)

        chart_window.destroy()

    chart_window = tki.Toplevel(root)
    chart_window.title("Create Bar Chart")

    x_label = tki.Label(chart_window, text="X-column:")
    x_label.grid(row=0, column=0)
    x_entry = tki.Entry(chart_window)
    x_entry.grid(row=0, column=1)

    y_label = tki.Label(chart_window, text="Y-column:")
    y_label.grid(row=1, column=0)
    y_entry = tki.Entry(chart_window)
    y_entry.grid(row=1, column=1)

    submit_button = tki.Button(chart_window, text="Plot", command=plot_chart)
    submit_button.grid(row=2, column=0, columnspan=2)


def create_pie_chart():

    def plot_chart():
        # Получаем значения из полей ввода для названий столбцов
        x_column = x_entry.get()

        # Проверяем, существует ли введенный столбец
        if x_column not in GDS.columns:
            messagebox.showerror(
                "ERROR", f"Column '{x_column}' does not exist in the file.")
            return

        # Проверяем, что столбец содержит числовые данные
        if pd.api.types.is_numeric_dtype(GDS[x_column]):
            # Строим график
            fig, ax = plt.subplots()
            GDS[x_column].value_counts().plot(kind='pie', ax=ax)

            # Уменьшаем размер изображения до подходящего
            fig.set_size_inches(
                4, 4)  # Примерное значение для подходящего размера

            # Очищаем фрейм с графиком и вставляем новый график
            for widget in chart_frame.winfo_children():
                widget.destroy()

            # Экспортируем изображение графика в канвас
            chart_canvas = FigureCanvasTkAgg(fig, master=chart_frame)
            chart_canvas.draw()
            chart_canvas.get_tk_widget().pack(fill='both', expand=True)

            chart_window.destroy()
        else:
            messagebox.showerror(
                "ERROR", "Column data must be numeric to plot a pie chart.")

    chart_window = tki.Toplevel(root)
    chart_window.title("Create Pie Chart")

    x_label = tki.Label(chart_window, text="Column:")
    x_label.grid(row=0, column=0)
    x_entry = tki.Entry(chart_window)
    x_entry.grid(row=0, column=1)

    submit_button = tki.Button(chart_window, text="Plot", command=plot_chart)
    submit_button.grid(row=2, column=0, columnspan=2)


def create_line_chart():

    def plot_chart():
        x_column = x_entry.get()
        y_column = y_entry.get()

        if x_column not in GDS.columns or y_column not in GDS.columns:
            messagebox.showerror(
                "ERROR",
                f"One or both of the columns '{x_column}' and '{y_column}' does not exist at file."
            )
            return

        # Преобразуем столбцы в числовой формат, если это необходимо
        x_data = pd.to_numeric(GDS[x_column], errors='coerce')
        y_data = pd.to_numeric(GDS[y_column], errors='coerce')

        # Проверяем, что столбцы содержат числовые данные
        if x_data.isna().all() or y_data.isna().all():
            messagebox.showerror("ERROR",
                                 "No numeric data in the specified columns.")
            return

        # Строим график
        fig, ax = plt.subplots()
        GDS.plot(x=x_column, y=y_column, kind='line', ax=ax)

        # Уменьшаем размер изображения до подходящего
        fig.set_size_inches(3.5, 3.5)

        # Очищаем фрейм с графиком и вставляем новый график
        for widget in chart_frame.winfo_children():
            widget.destroy()

        # Экспортируем изображение графика в канвас
        chart_canvas = FigureCanvasTkAgg(fig, master=chart_frame)
        chart_canvas.draw()
        chart_canvas.get_tk_widget().pack(fill='both', expand=True)

        chart_window.destroy()

    chart_window = tki.Toplevel(root)
    chart_window.title("Create Line Chart")

    x_label = tki.Label(chart_window, text="X Column:")
    x_label.grid(row=0, column=0)
    x_entry = tki.Entry(chart_window)
    x_entry.grid(row=0, column=1)

    y_label = tki.Label(chart_window, text="Y Column:")
    y_label.grid(row=1, column=0)
    y_entry = tki.Entry(chart_window)
    y_entry.grid(row=1, column=1)

    submit_button = tki.Button(chart_window, text="Plot", command=plot_chart)
    submit_button.grid(row=2, column=0, columnspan=2)


def create_scatter_chart():

    def plot_chart():
        # Получаем значения из полей ввода для названий столбцов
        x_column = x_entry.get()
        y_column = y_entry.get()

        # Проверяем, существуют ли введенные столбцы
        if x_column not in GDS.columns or y_column not in GDS.columns:
            messagebox.showerror(
                "ERROR",
                f"One or both of the columns '{x_column}' and '{y_column}' does not exist at file."
            )
            return

        # Преобразуем столбцы в числовой формат, если это необходимо
        x_data = pd.to_numeric(GDS[x_column], errors='coerce')
        y_data = pd.to_numeric(GDS[y_column], errors='coerce')

        # Проверяем, что столбцы содержат числовые данные
        if x_data.isna().all() or y_data.isna().all():
            messagebox.showerror("ERROR",
                                 "No numeric data in the specified columns.")
            return

        # Строим график
        fig, ax = plt.subplots()
        GDS.plot(x=x_column, y=y_column, kind='scatter', ax=ax)

        # Уменьшаем размер изображения до подходящего
        fig.set_size_inches(3.5, 3.5)

        # Очищаем фрейм с графиком и вставляем новый график
        for widget in chart_frame.winfo_children():
            widget.destroy()

        # Экспортируем изображение графика в канвас
        chart_canvas = FigureCanvasTkAgg(fig, master=chart_frame)
        chart_canvas.draw()
        chart_canvas.get_tk_widget().pack(fill='both', expand=True)

        chart_window.destroy()

    chart_window = tki.Toplevel(root)
    chart_window.title("Create Scatter Chart")

    x_label = tki.Label(chart_window, text="X Column:")
    x_label.grid(row=0, column=0)
    x_entry = tki.Entry(chart_window)
    x_entry.grid(row=0, column=1)

    y_label = tki.Label(chart_window, text="Y Column:")
    y_label.grid(row=1, column=0)
    y_entry = tki.Entry(chart_window)
    y_entry.grid(row=1, column=1)

    submit_button = tki.Button(chart_window, text="Plot", command=plot_chart)
    submit_button.grid(row=2, column=0, columnspan=2)


def show_file_menu(event):
    file_menu.post(event.x_root, event.y_root)


def show_report_menu(event):
    report_menu.post(event.x_root, event.y_root)


def change_color_theme(theme):
    if theme == "dark":
        # Dark theme colors
        root.config(bg="#21213A")
        file_menu_frame.config(bg="#1f1f1f")
        report_menu_frame.config(bg="#1f1f1f")
        chart_manager_frame.config(bg="#1f1f1f")
        settings_frame.config(bg="#1f1f1f")
        top.config(bg="white", fg='black')
        chart.config(bg="white", fg='black')
        file_label.config(fg="white", background="#333333"
                          )  # Красный цвет текста для подписи "File Manager"
        report_label.config(
            fg="white", background="#333333"
        )  # Красный цвет текста для подписи "Report Manager"
        chart_label.config(fg="white", background="#333333"
                           )  # Красный цвет текста для подписи "Chart Manager"
    elif theme == "light":
        # Light theme colors
        root.config(bg="#FFFFCC")
        file_menu_frame.config(
            bg="white")  # Кремовый фон для фрейма меню "Файл"
        report_menu_frame.config(
            bg="white")  # Кремовый фон для фрейма меню "Отчет"
        chart_manager_frame.config(
            bg="white")  # Кремовый фон для фрейма меню "График"
        settings_frame.config(
            bg="white")  # Кремовый фон для фрейма меню "Настройки"
        top.config(bg="black", fg="white")  # Кремовый фон для таблицы
        chart.config(bg="black", fg="white")  # фон для менеджера графиков
        file_label.config(fg="black", background="white"
                          )  # Черный цвет текста для подписи "File Manager"
        report_label.config(
            fg="black", background="white"
        )  # Черный цвет текста для подписи "Report Manager"
        chart_label.config(fg="black", background="white"
                           )  # Черный цвет текста для подписи "Chart Manager"
    elif theme == "spooky":
        root.config(bg="black")
        file_menu_frame.config(
            bg="#2E2E2E")  # Темно-серый фон для фрейма меню "Файл"
        report_menu_frame.config(
            bg="#2E2E2E")  # Темно-серый фон для фрейма меню "Отчет"
        chart_manager_frame.config(
            bg="#2E2E2E")  # Темно-серый фон для фрейма меню "График"
        settings_frame.config(
            bg="#2E2E2E")  # Темно-серый фон для фрейма меню "Настройки"
        top.config(bg="#2E2E2E", fg="red")  # Темно-серый фон для таблицы
        chart.config(bg="#2E2E2E",
                     fg="red")  # Темно-серый фон для менеджера графиков
        file_label.config(fg="red", background='#2E2E2E'
                          )  # Красный цвет текста для подписи "File Manager"
        report_label.config(
            fg="red", background='#2E2E2E'
        )  # Красный цвет текста для подписи "Report Manager"
        chart_label.config(fg="red", background='#2E2E2E'
                           )  # Красный цвет текста для подписи "Chart Manager"

        spooky_window = tki.Toplevel(root)
        spooky_window.attributes('-fullscreen',
                                 True)  # Cover the entire screen
        spooky_window.configure(bg="black")

        # Load the spooky GIF
        spooky_gif_path = os.path.abspath(
            os.path.join(os.path.dirname(__file__), '..', 'images',
                         "spooky_shooting.gif"))
        spooky_gif = Image.open(spooky_gif_path)

        # Create a list to store each frame of the GIF
        gif_frames = []
        for frame_number in range(spooky_gif.n_frames):
            spooky_gif.seek(frame_number)
            frame = spooky_gif.convert('RGBA').resize(
                (root.winfo_screenwidth(), root.winfo_screenheight()),
                Image.Resampling.LANCZOS)
            gif_frames.append(ImageTk.PhotoImage(frame))

        # Create a label to display the GIF frames
        spooky_label = tki.Label(spooky_window, bg="black")
        spooky_label.pack(fill="both", expand=True)

        # Function to update the GIF frames with a slower playback speed
        def update_gif(frame_number=0):
            try:
                spooky_label.configure(image=gif_frames[frame_number])
                frame_number += 1
                frame_number %= len(gif_frames)  # Loop through frames
                spooky_window.after(
                    500, update_gif,
                    frame_number)  # Increase delay to slow down playback
            except Exception as e:
                print(f"Error updating GIF: {e}")

        # Start updating the GIF frames
        update_gif()

        # Destroy the Toplevel window after 1 second
        spooky_window.after(2200, spooky_window.destroy)


def show_color_theme_menu():
    # Create a new Toplevel window for the color theme selection
    theme_window = tki.Toplevel(root)
    theme_window.title("Select Color Theme")

    # Define a function to set the selected theme and close the window
    def set_theme_and_close(theme):
        change_color_theme(theme)
        theme_window.destroy()

    # Create buttons for each theme
    dark_theme_button = tki.Button(theme_window,
                                   text="Dark Theme",
                                   command=lambda: set_theme_and_close("dark"))
    dark_theme_button.pack(pady=5)

    light_theme_button = tki.Button(
        theme_window,
        text="Light Theme",
        command=lambda: set_theme_and_close("light"))
    light_theme_button.pack(pady=5)

    skeleton_theme_button = tki.Button(
        theme_window,
        text="Spooky Theme",
        command=lambda: set_theme_and_close("spooky"))
    skeleton_theme_button.pack(pady=5)

    # Main loop for Tkinter events
    theme_window.mainloop()


# Функция для изменения шрифта
def change_font(font):
    file_label.config(font=(font, 14, 'bold'))
    report_label.config(font=(font, 14, 'bold'))
    chart_label.config(font=(font, 14, 'bold'))
    file_button.config(font=(font, 15, 'italic'))
    but_clear_file.config(font=(font, 15, 'italic'))
    but_clear_report.config(font=(font, 15, 'italic'))
    but_save.config(font=(font, 15, 'italic'))
    report_button.config(font=(font, 15, 'italic'))
    chart_button.config(font=(font, 15, 'italic'))
    settings_button.config(font=(font, 15, 'italic'))
    top.config(font=(font, 14, 'bold'))
    chart.config(font=(font, 14, 'bold'))


# Функция для отображения меню выбора шрифта
def show_font_menu():
    font = simpledialog.askstring(
        "Change font",
        "Enter font (Arial, Times, Courier or different font that Tkinter contains):"
    )
    if font:
        change_font(font)


# Подготовка просмотра
rgb = (30, 30, 30)  # Определение цвета

file_data = {}
GDS = pd.DataFrame([])
height = GDS.shape[0]
width = GDS.shape[1]
pnt = np.empty([])
vrs = np.empty([])
top = []
tbl = []
report = pd.DataFrame([])

# Построение окна
root = tki.Tk()
root.title("TableManager")
# Установка минимального размера окна
root.minsize(800, 600)

# цвет фона
background_color = rgb_hack((33, 33, 58))
root.configure(bg=background_color)

# Создание фрейма для кнопки "Файл"
file_menu_frame = tki.Frame(root,
                            bg=rgb_hack(rgb),
                            relief="raised",
                            borderwidth=2)
file_menu_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nw")

# Подпись фрейма "Файл"
file_label = tki.Label(file_menu_frame,
                       text="File Manager",
                       font=('Arial', 14, 'bold'),
                       bg=rgb_hack(rgb),
                       fg="#FFFFFF")
file_label.grid(row=0, column=0, padx=(10, 5), pady=5, sticky="w")

# Создание кнопки "File"
file_button = tki.Button(file_menu_frame,
                         text='File...',
                         font=('Arial', 15, 'italic'),
                         bg='#d3d3d3',
                         fg='#000000',
                         padx=10,
                         pady=5,
                         width=8)
file_button.grid(row=0, column=1, padx=(5, 0), pady=5)

# Привязка меню к кнопке "File"
file_button.bind("<Button-1>", show_file_menu)

file_menu = tki.Menu(root, tearoff=0)
file_menu.add_command(label="Read Data", command=read_data)
file_menu.add_command(label="Save As...", command=lambda: save_table_as_csv(GDS, file_data))
file_menu.add_command(label="Save", command=lambda: save_table(GDS, file_data))
file_menu.add_separator()
file_menu.add_command(label="Exit", command=root.destroy)

# Кнопка "Clear"
but_clear_file = tki.Button(file_menu_frame,
                            text='Clear Table',
                            font=('Arial', 15, 'italic'),
                            bg='#d3d3d3',
                            fg='#000000',
                            command=clear_top,
                            padx=10,
                            pady=5,
                            width=8)
but_clear_file.grid(row=1, column=1, padx=(5, 10), pady=5)

# Создаем фрейм (контейнер) для менеджера отчетов
report_menu_frame = tki.Frame(root,
                              bg=rgb_hack(rgb),
                              relief="raised",
                              borderwidth=2)
report_menu_frame.grid(row=0, column=1, padx=10, pady=10, sticky="nw")

# Подпись фрейма "Менеджер отчетов"
report_label = tki.Label(report_menu_frame,
                         text="Report Manager",
                         font=('Arial', 14, 'bold'),
                         bg=rgb_hack(rgb),
                         fg="#FFFFFF")
report_label.grid(row=0, column=0, padx=(10, 5), pady=5, sticky="w")

# Кнопка "Clear"
but_clear_report = tki.Button(report_menu_frame,
                              text='Clear Manager',
                              font=('Arial', 15, 'italic'),
                              bg='#d3d3d3',
                              fg='#000000',
                              command=clear_right,
                              padx=10,
                              pady=5,
                              width=8)
but_clear_report.grid(row=1, column=1, padx=(5, 10), pady=5)

# Создание кнопки "Report"
report_button = tki.Menubutton(report_menu_frame,
                               text='Report...',
                               font=('Arial', 15, 'italic'),
                               bg='#d3d3d3',
                               fg='#000000',
                               padx=10,
                               pady=5,
                               width=8)
report_button.grid(row=0, column=1, padx=(5, 0), pady=5)

# Привязка меню к кнопке "Report"
report_button.bind("<Button-1>", show_report_menu)

# Назначаем стиль кнопке "Report"
report_button.config(relief="raised",
                     font=('Arial', 15, 'italic'),
                     bg='#d3d3d3',
                     fg='#000000',
                     padx=10,
                     pady=5,
                     width=8)

# Создание меню для кнопки "Report"
report_menu = tki.Menu(report_button, tearoff=0)
report_menu.add_command(label="Generate Report", command=generate_report)
report_menu.add_separator()
report_menu.add_command(label="Generate Pivot table", command=generate_pivot_table)
report_menu.add_separator()
report_menu.add_command(label="Save report", command=lambda: save_text_report(report.to_string(index=False)))

# Устанавливаем меню для кнопки
report_button.config(menu=report_menu)

# Создаем фрейм менеджера графиков
chart_manager_frame = tki.Frame(root,
                                bg=rgb_hack(rgb),
                                relief="raised",
                                borderwidth=2)
chart_manager_frame.grid(row=0, column=2, padx=10, pady=10, sticky="nw")

# Подпись фрейма "Менеджер графиков"
chart_label = tki.Label(chart_manager_frame,
                        text="Chart Manager",
                        font=('Arial', 14, 'bold'),
                        bg=rgb_hack(rgb),
                        fg="white")
chart_label.grid(row=0, column=0, padx=(10, 5), pady=5, sticky="w")

# Создание кнопки "Chart..."
chart_button = tki.Menubutton(chart_manager_frame,
                              text='Chart...',
                              font=('Arial', 15, 'italic'),
                              bg='#d3d3d3',
                              fg='#000000',
                              padx=10,
                              pady=5,
                              width=8)
chart_button.grid(row=0, column=1, padx=(5, 0), pady=5)

# сохранение графика
but_save = tki.Button(chart_manager_frame,
                      text='Save...',
                      font=('Arial', 15, 'italic'),
                      bg='#d3d3d3',
                      fg='#000000',
                      command=save_chart_image,
                      padx=10,
                      pady=5,
                      width=8)
but_save.grid(row=1, column=1, padx=(5, 10), pady=5)

# Привязка меню к кнопке "Chart..."
chart_menu = tki.Menu(chart_button, tearoff=0)
chart_menu.add_command(label="Bar", command=create_bar_chart)
chart_menu.add_command(label="Pie chart", command=create_pie_chart)
chart_menu.add_command(label="Line", command=create_line_chart)
chart_menu.add_command(label="Scatter", command=create_scatter_chart)

# Устанавливаем меню для кнопки "Chart..."
chart_button.config(menu=chart_menu)

# фрейм для меню настроек
settings_frame = tki.Frame(root,
                           bg=rgb_hack(rgb),
                           relief="raised",
                           borderwidth=2)
settings_frame.grid(row=0, column=3, padx=10, pady=10, sticky="nw")

# кнопка настроек
work_folder_path = os.path.join(parent_dir, 'images')
settings_icon_path = os.path.join(work_folder_path, "settings_icon.png")
settings_icon_size = 40

gear_image = Image.open(settings_icon_path)
gear_image = gear_image.resize((settings_icon_size, settings_icon_size),
                               Image.Resampling.LANCZOS)
gear_photo = ImageTk.PhotoImage(gear_image)

# Создание кнопки с изображением
settings_button = tki.Menubutton(settings_frame,
                                 image=gear_photo,
                                 width=settings_icon_size,
                                 height=settings_icon_size,
                                 font=('Arial', 15, 'italic'),
                                 padx=10,
                                 pady=5)
settings_button.image = gear_photo  # Сохранение ссылки на изображение
settings_button.grid(row=0, column=0, padx=10, pady=10)

# меню настроек
settings_menu = tki.Menu(settings_button, tearoff=0)
settings_menu.add_command(label="Change theme", command=show_color_theme_menu)
settings_menu.add_command(label="Change fonts", command=show_font_menu)

settings_button.config(menu=settings_menu)

# Создаем фрейм (контейнер) в котором будет размещена таблица
top = tki.LabelFrame(root,
                     text="Table",
                     bg=rgb_hack(rgb),
                     font=('Arial', 14, 'bold'),
                     fg="#FFFFFF",
                     relief="raised",
                     borderwidth=2)
top.grid(column=0, row=1, columnspan=2, sticky="nsew", padx=10, pady=10)

# Создаем canvas и скроллбары для таблицы
canvas = tki.Canvas(top, bg=rgb_hack(rgb), width=600, height=350)
canvas.grid(row=0, column=0, sticky="nsew")

v_scrollbar = tki.Scrollbar(top, orient="vertical", command=canvas.yview)
v_scrollbar.grid(row=0, column=1, sticky="ns")

h_scrollbar = tki.Scrollbar(top, orient="horizontal", command=canvas.xview)
h_scrollbar.grid(row=1, column=0, sticky="ew")

canvas.configure(yscrollcommand=v_scrollbar.set,
                 xscrollcommand=h_scrollbar.set)

# Устанавливаем привязку для скроллбаров к холсту таблицы
canvas.bind('<Configure>',
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

# Создаем фрейм внутри canvas для таблицы
table_frame = tki.Frame(canvas, bg=rgb_hack(rgb))
canvas.create_window((0, 0), window=table_frame, anchor="nw")

# Создаем фрейм (контейнер) для графика
chart = tki.LabelFrame(root,
                       text="Manager",
                       bg=rgb_hack(rgb),
                       font=('Arial', 14, 'bold'),
                       fg="#FFFFFF",
                       relief="raised",
                       borderwidth=2)
chart.grid(column=2, row=1, padx=10, pady=10, sticky="nsew")

# Создаем canvas для графика
chart_canvas = tki.Canvas(chart, bg=rgb_hack(rgb), width=350, height=350)
chart_canvas.grid(row=0, column=0, sticky="nsew")

# Создаем вертикальный скроллбар для canvas
v_scrollbar_chart = tki.Scrollbar(chart,
                                  orient="vertical",
                                  command=chart_canvas.yview)
v_scrollbar_chart.grid(row=0, column=1, sticky="ns")

# Создаем горизонтальный скроллбар для canvas
h_scrollbar_chart = tki.Scrollbar(chart,
                                  orient="horizontal",
                                  command=chart_canvas.xview)
h_scrollbar_chart.grid(row=1, column=0, sticky="ew")

# Создаем фрейм внутри canvas для графика
chart_frame = tki.Frame(chart_canvas, bg=rgb_hack(rgb))
chart_canvas.create_window((0, 0), window=chart_frame, anchor="nw")

# Конфигурируем canvas для использования скроллбаров
chart_canvas.config(yscrollcommand=v_scrollbar_chart.set,
                    xscrollcommand=h_scrollbar_chart.set)


# Привязываем функцию к событию изменения размера холста
def configure_scroll_region(event):
    chart_canvas.config(scrollregion=chart_canvas.bbox("all"))


# Устанавливаем привязку для скроллбаров к холсту графика
chart_canvas.bind('<Configure>', configure_scroll_region)

# Устанавливаем небольшой отступ между полем для графика и вертикальным скроллбаром таблицы
chart_canvas.grid_rowconfigure(0, weight=1)
chart_canvas.grid_columnconfigure(0, weight=1)
chart_canvas.grid_columnconfigure(
    2, minsize=40)  # добавляем пустую колонку для отступа

# Определяем функцию mainloop() для цикла событий Tkinter
root.mainloop()
