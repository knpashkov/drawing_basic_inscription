import tkinter as tk
from tkinter import filedialog


class MainApplication(tk.Tk):
    def __init__(self):
        super().__init__()

        # Настройка главного окна
        self.button = None
        self.title("Пакетное редактирование основной надписи")  # Заголовок окна
        self.geometry("800x600")  # Размер окна (ширина x высота)
        self.resizable(False, False)  # Разрешение на изменение размера (по ширине, по высоте)

        # Вызов метода для создания виджетов
        self.create_widgets()

        # Переменная для хранения списка выбранных файлов
        self.selected_files = []

    def create_widgets(self):
        """Создание и размещение виджетов в окне"""

        # Кнопка выбора файлов
        self.button = tk.Button(self, text="Выбрать файлы",
                                command=self.on_button_click)
        self.button.pack(pady=10)

    def on_button_click(self):
        """Обработчик нажатия кнопки"""
        print("Кнопка была нажата!")
        self.select_files()

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
            pass


if __name__ == "__main__":
    app = MainApplication()
    app.mainloop()