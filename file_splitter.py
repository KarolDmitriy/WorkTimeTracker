# file_splitter.py

import openpyxl
from datetime import datetime, time

def format_date(date):
    if isinstance(date, datetime):
        return date.strftime("%d.%m.%Y")
    return None

def format_time(time_value):
    if isinstance(time_value, time):
        return time_value.strftime("%H:%M:%S")
    return None

def create_work_schedule_file(output_file, rows):
    try:
        workbook = openpyxl.Workbook()
        sheet = workbook.active

        # Добавляем строку заголовков
        header_row = ["Дата входа", "Время входа", "Дата выхода", "Время выхода"]
        sheet.append(header_row)

        # Записываем остальные строки в новый файл
        for row in rows:
            sheet.append(row)

        # Сохраняем новый файл
        workbook.save(output_file)

    except Exception as e:
        print(f"Ошибка при создании файла: {e}")

def split_work_schedules(input_file="data/work_schedules.xlsx", output_folder="data/graph"):
    try:
        # Открываем файл work_schedules
        schedules_workbook = openpyxl.load_workbook(input_file)
        schedules_sheet = schedules_workbook.active

        # Создаем папку для выходных файлов, если её нет
        import os
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        current_schedule = None
        current_schedule_rows = []

        for row in schedules_sheet.iter_rows(min_row=2, values_only=True):
            # Если текущий график изменился, создаем новый файл
            if row[0] != current_schedule:
                if current_schedule_rows:
                    # Создаем новый файл и копируем строки
                    output_file = f"{output_folder}/{current_schedule}.xlsx"
                    create_work_schedule_file(output_file, current_schedule_rows)

                # Обновляем текущий график и очищаем строки
                current_schedule = row[0]
                current_schedule_rows = []

            # Форматируем дату и время
            formatted_row = [
                format_date(row[1]),
                format_time(row[2]),
                format_date(row[3]),
                format_time(row[4]),
                # ... другие столбцы
            ]

            # Добавляем текущую строку к текущему графику
            current_schedule_rows.append(formatted_row)

        # Создаем файл для последнего графика
        if current_schedule_rows:
            output_file = f"{output_folder}/{current_schedule}.xlsx"
            create_work_schedule_file(output_file, current_schedule_rows)

    except Exception as e:
        print(f"Ошибка при выполнении операций: {e}")

# Вызов функции в main.py
from file_splitter import split_work_schedules

def main():
    split_work_schedules()
    # Другие операции...

if __name__ == "__main__":
    main()
