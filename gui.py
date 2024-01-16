import tkinter as tk
from tkinter import filedialog
from functools import partial
from file_splitter import split_work_schedules
from main import process_daily_entries
from work_time_utils import update_employee_absences
from reconciliation import check_plan_actual_time

class App:
    def __init__(self, master):
        self.master = master
        master.title("WorkTimeTracker")

        # Установим ширину и высоту окна
        master.geometry("800x400")

        # Фреймы для кнопок
        self.frame_buttons = [tk.Frame(master) for _ in range(4)]

        # Кнопки для выбора файлов
        self.file_buttons = [
            ("Выбрать файл для разделения графиков", filedialog.askopenfilename, split_work_schedules),
            ("Выбрать файл для проверки турникета", filedialog.askopenfilename, process_daily_entries),
            ("Выбрать файл для заполнения плановых табелей", filedialog.askopenfilename, update_employee_absences),
            ("Выбрать файл для проверки отклонений в табеле", filedialog.askopenfilename, check_plan_actual_time)
        ]

        # Создание и размещение кнопок в фреймах
        for i, (text, file_dialog_func, process_func) in enumerate(self.file_buttons):
            tk.Label(self.frame_buttons[i % 4], text=text).pack(pady=5)
            tk.Button(self.frame_buttons[i % 4], text="Выбрать файл", command=partial(self.choose_file, file_dialog_func)).pack(pady=5)
            tk.Button(self.frame_buttons[i % 4], text="Сохранить", command=partial(self.save_file, process_func)).pack(pady=5)
            tk.Button(self.frame_buttons[i % 4], text="Запустить", command=partial(self.run_function, file_dialog_func, process_func)).pack(pady=5)

        # Размещение фреймов
        for i, frame in enumerate(self.frame_buttons):
            frame.grid(row=0, column=i, padx=10)

        # Метка для вывода результата
        self.result_label = tk.Label(master, text="", fg="black")
        self.result_label.grid(row=1, column=0, columnspan=4, pady=10)  # Изменил здесь

    def choose_file(self, file_dialog_func):
        file_path = file_dialog_func(title="Выберите файл", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.result_label.config(text=f"Выбран файл: {file_path}", fg="black")

    def save_file(self, process_func):
        # Здесь вы можете добавить функциональность для сохранения файла, если это применимо
        self.result_label.config(text="Файл сохранен", fg="green")

    def run_function(self, file_dialog_func, process_func):
        file_path = file_dialog_func(title="Выберите файл", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            result = process_func(file_path)
            self.result_label.config(text=result, fg="green" if result else "red")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
