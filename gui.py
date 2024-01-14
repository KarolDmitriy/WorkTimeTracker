import pandas as pd
from openpyxl import load_workbook
from datetime import datetime, timedelta
from openpyxl.styles import PatternFill
import tkinter as tk
from tkinter import filedialog

def check_plan_actual_time_gui():
    root = tk.Tk()
    root.title("Проверка фактического и планового времени")

    def browse_comments_file():
        comments_file_path = filedialog.askopenfilename(title="Выберите файл comments_file", filetypes=[("Excel files", "*.xlsx")])
        comments_file_var.set(comments_file_path)

    def browse_work_schedule_file():
        work_schedule_file_path = filedialog.askopenfilename(title="Выберите файл work_schedule_file", filetypes=[("Excel files", "*.xlsx")])
        work_schedule_file_var.set(work_schedule_file_path)

    def run_check():
        comments_file_path = comments_file_var.get()
        work_schedule_file_path = work_schedule_file_var.get()

        if not comments_file_path or not work_schedule_file_path:
            result_label.config(text="Выберите оба файла", fg="red")
            return

        check_plan_actual_time(comments_file_path, work_schedule_file_path)
        result_label.config(text="Процесс успешно завершен!", fg="green")

    # Переменные для хранения путей к файлам
    comments_file_var = tk.StringVar()
    work_schedule_file_var = tk.StringVar()

    # Метка и кнопка для выбора файла comments_file
    tk.Label(root, text="Выберите файл comments_file:").grid(row=0, column=0)
    tk.Entry(root, textvariable=comments_file_var, width=50).grid(row=0, column=1)
    tk.Button(root, text="Обзор", command=browse_comments_file).grid(row=0, column=2)

    # Метка и кнопка для выбора файла work_schedule_file
    tk.Label(root, text="Выберите файл work_schedule_file:").grid(row=1, column=0)
    tk.Entry(root, textvariable=work_schedule_file_var, width=50).grid(row=1, column=1)
    tk.Button(root, text="Обзор", command=browse_work_schedule_file).grid(row=1, column=2)

    # Кнопка для запуска проверки
    tk.Button(root, text="Запустить проверку", command=run_check).grid(row=2, column=0, columnspan=3, pady=10)

    # Метка для вывода результата
    result_label = tk.Label(root, text="", fg="black")
    result_label.grid(row=3, column=0, columnspan=3)

    root.mainloop()

# Пример использования
# check_plan_actual_time_gui()
