import openpyxl

def duplicate_and_append_rows(sheet, start_row, start_column, end_row, end_column, times):
    for _ in range(times):
        for row in range(start_row, end_row + 1):
            new_row = []
            for column in range(start_column, end_column + 1):
                new_row.append(sheet.cell(row=row, column=column).value)
            sheet.append(new_row)


# def duplicate_and_append_rows(sheet, start_row, start_column, end_row, end_column, times):
#     for _ in range(times):
#         for row in range(start_row, end_row + 1):
#             new_row = []
#             for column in range(start_column, end_column + 1):
#                 source_cell = sheet.cell(row=row, column=column)
#                 new_cell = sheet.cell(row=sheet.max_row + 1, column=column, value=source_cell.value)
#
#                 # Копирование форматирования
#                 new_cell.font = copy(source_cell.font)
#                 new_cell.border = copy(source_cell.border)
#                 new_cell.fill = copy(source_cell.fill)
#                 new_cell.number_format = copy(source_cell.number_format)
#                 new_cell.protection = copy(source_cell.protection)
#                 new_cell.alignment = copy(source_cell.alignment)

# Открываем файл
file_path = 'data/attendance_template.xlsx'
workbook = openpyxl.load_workbook(file_path)
sheet = workbook.active

# Указываем параметры диапазона и количество вставок
start_row, end_row = 13, 16
start_column, end_column = 1, 42  # AP соответствует 42
times_to_duplicate = 10

# Дублируем и добавляем строки 10 раз
duplicate_and_append_rows(sheet, start_row, start_column, end_row, end_column, times_to_duplicate)

# Сохраняем изменения
workbook.save(file_path)
