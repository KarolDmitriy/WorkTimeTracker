# указываем файл output_file, проставляем в ячейку F7 количество дней в текущем месяце, и он заполняется плановыми часами
from openpyxl import load_workbook

def update_employee_absences(input_file, absences_file="data/work_schedules_for_the_year.xlsx",
                             output_file="data/ПАЛ/Аналитическая лаборатория январь 2024.xlsx"):
    try:
        # Открываем указанный файл excel
        workbook = load_workbook(input_file)

        # Записываем значение ячейки F7 в переменную numberOfDays
        sheet = workbook.active
        numberOfDays = sheet.cell(row=7, column=6).value

        # Если ячейка F7 содержит целое число, используем его
        if not isinstance(numberOfDays, int):
            numberOfDays = 0

        # Открываем файл employee_absences
        absences_workbook = load_workbook(absences_file)
        absences_sheet = absences_workbook.active

        startRow = 13
        rowToName = 13

        for _ in range((sheet.max_row - startRow) // 4 + 1):
            # Записываем значения ячеек M1, P1, AQ13 в переменную monthYearGraph
            monthYearGraph = f"{sheet.cell(row=1, column=13).value} {sheet.cell(row=1, column=16).value} {sheet.cell(row=startRow, column=43).value}"

            # Находим ячейку с таким же значением, как в monthYearGraph
            absences_value_cell = None
            for row in absences_sheet.iter_rows(min_col=1, max_col=1, min_row=1, max_row=absences_sheet.max_row):
                for cell in row:
                    if cell.value == monthYearGraph:
                        absences_value_cell = cell
                        break
                if absences_value_cell:
                    break

            # Если ячейка найдена, копируем диапазон значений
            if absences_value_cell:
                source_range = absences_sheet.iter_cols(min_col=absences_value_cell.column + 1,
                                                        max_col=absences_value_cell.column + numberOfDays,
                                                        min_row=absences_value_cell.row, max_row=absences_value_cell.row)
                values_to_paste = [cell[0].value for cell in source_range]

                # Вставляем диапазон в указанный файл excel
                for i, value in enumerate(values_to_paste):
                    # Проверяем, есть ли буквы в конце значения и является ли значение строкой
                    if value and isinstance(value, str) and value[-1].isalpha():
                        # Убираем букву в конце, заменяем запятую на точку, преобразуем в float и вставляем
                        value_without_letter = value[:-1].replace(',', '.')
                        sheet.cell(row=rowToName, column=i + 6, value=float(value_without_letter))
                        # В следующую ячейку вставляем цифру 8
                        sheet.cell(row=rowToName + 1, column=i + 6, value=8)
                    else:
                        # Если нет букв в конце или не является строкой, просто вставляем значение
                        sheet.cell(row=rowToName, column=i + 6, value=value)

            # Итерация по работникам
            startRow += 4
            rowToName += 4

        # Сохраняем изменения
        workbook.save(output_file)
        absences_workbook.save(absences_file)

    except Exception as e:
        print(f"Ошибка: {e}")

# Пример использования
update_employee_absences("data/ПАЛ/Аналитическая лаборатория январь 2024.xlsx")
