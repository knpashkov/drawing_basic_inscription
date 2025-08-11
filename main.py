import tkinter as tk
from tkinter import filedialog

import win32com.client
import pythoncom
from win32com.client import Dispatch, gencache


class MainApplication(tk.Tk):
    def __init__(self):
        super().__init__()

        # Подключение Компаса
        self.kompas = Kompas()

        # Настройка главного окна
        self.main_frame = None
        self.btn_start = None
        self.name_designer = None
        self.name_control = None
        self.name_technologist = None
        self.name_norm = None
        self.name_approve = None
        self.pdf_check = None
        self.btn_choose = None

        self.title("Пакетное редактирование основной надписи")  # Заголовок окна
        # self.geometry("600x400")  # Размер окна (ширина x высота)
        self.minsize(width=300, height=self.winfo_reqheight())
        self.resizable(False, False)  # Разрешение на изменение размера (по ширине, по высоте)



        # Вызов метода для создания виджетов
        self.create_widgets()

        # Переменная для хранения списка выбранных файлов
        self.selected_files = []

    def create_widgets(self):
        """Создание и размещение виджетов в окне"""

        self.main_frame = tk.Frame(self)
        self.main_frame.pack(padx=5, pady=5)

        # Кнопка выбора файлов
        self.btn_choose = tk.Button(self.main_frame, text="Выбрать файлы", command=self.select_files)
        self.btn_choose.pack(fill='x', ipady=5)
        
        self.name_designer = InputStringWidget(self.main_frame, text="Разработчик")
        self.name_designer.pack(fill='x')

        self.name_control = InputStringWidget(self.main_frame, text="Проверяющий")
        self.name_control.pack(fill='x')

        self.name_technologist = InputStringWidget(self.main_frame, text="Техконтроль")
        self.name_technologist.pack(fill='x')

        self.name_norm = InputStringWidget(self.main_frame, text="Нормоконтроль")
        self.name_norm.pack(fill='x')

        self.name_approve = InputStringWidget(self.main_frame, text="Утверждающий")
        self.name_approve.pack(fill='x')

        self.pdf_check = CheckPdfWidget(self.main_frame)
        self.pdf_check.pack(fill='x')

        self.btn_start = tk.Button(self.main_frame, text="Начать", command=self.on_btn_start_click)
        self.btn_start.pack(fill='x', ipady=5)


    def on_btn_start_click(self):
        print(self.name_designer.get_value())


    def select_files(self):
        """Открытие диалога выбора файлов"""
        filetypes = (
            ("Документы Компас", "*.cdw *.spw"),
            ("Чертежи Компас", "*.cdw"),
            ("Спецификации Компас", "*.spw")
        )

        files = filedialog.askopenfilenames(
            title="Выберите файлы",
            initialdir="/",
            filetypes=filetypes
        )

        if files:
            self.selected_files = list(files)

    def change_files(self):
        for path in self.selected_files:
            self.kompas.change_document(path)


class InputStringWidget(tk.Frame):
    def __init__(self, master, text):
        super().__init__(master)

        self.position_label = tk.Label(self, text=text, anchor='w', width=20)
        self.position_label.pack(side=tk.LEFT, ipady=5)
        self.surname_input = tk.Entry(self)
        self.surname_input.pack(side=tk.LEFT, fill='x', expand=True, ipady=5)
        self.date_input = tk.Entry(self)
        self.date_input.pack(side=tk.LEFT, fill='x', expand=True, ipady=5)

    def get_value(self):
        return self.surname_input.get()


class CheckPdfWidget(tk.Frame):
    def __init__(self, master):
        super().__init__(master)

        self.check_pdf_var = tk.IntVar()

        self.pdf_check = tk.Checkbutton(self, variable=self.check_pdf_var)
        self.pdf_check.pack(side=tk.LEFT)
        self.pdf_label = tk.Label(self, text='Создать PDF')
        self.pdf_label.pack(side=tk.LEFT, fill='x', expand=True, ipady=5)


    def get_check(self):
        return self.check_pdf_var


class Kompas:
    def __init__(self):
        # Подключаем API интерфейсов
        self.kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
        self.kompas_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)

        # Подключаем объекты верхнего уровня
        self.application = self.kompas_api7_module.IApplication(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(self.kompas_api7_module.IApplication.CLSID, pythoncom.IID_IDispatch))  # IApplication
        self.kompas_object = self.kompas_api5_module.KompasObject(Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(self.kompas_api5_module.KompasObject.CLSID, pythoncom.IID_IDispatch))
        self.kompas6_constants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants

    def change_document(self, path: str):
        self.application.Documents.Open(path, False, False)



if __name__ == "__main__":
    app = MainApplication()
    app.mainloop()