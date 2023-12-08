import pandas as pd
import openpyxl
import os
from tqdm import tqdm

def process_daily_entries(file_path="data/daily_entries.xlsx", output_dir="data/graph", limit=1000):
    print("Запущена функция process_daily_entries")
    try:
        # Загрузка данных из файла daily_entries (нужно пересохранить через блокнот)
        daily_entries = pd.read_excel(file_path, sheet_name='Лист1')

        # Создание объекта pd.ExcelWriter перед началом цикла
        output_dir = "data/graph"
        output_file = "comments.xlsx"
        comment_file_path = os.path.join(output_dir, output_file)

        # Проверка наличия директории, и создание, если она отсутствует
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # Создание нового excel файла с комментариями
        if not os.path.exists(comment_file_path):
            df = pd.DataFrame(
                columns=["Комментарий", "Табельный", "ФИО", "Дата входа", "Время входа", "Дата выхода", "Время выхода",
                         "График"])
            df.to_excel(comment_file_path, index=False, engine='openpyxl', sheet_name='Sheet1')

        # Итерация по строкам из файла daily_entries
        for index, row in daily_entries.iterrows():
            # Преобразование строки в список
            row_list = row.tolist()

            # Преобразование даты в нужный формат
            formatted_row = [
                row_list[0],  # Табельный
                row_list[1],  # ФИО
                row_list[2].strftime("%d.%m.%Y"),  # Дата входа
                row_list[3].strftime("%H:%M:%S"),  # Время входа
                row_list[4].strftime("%d.%m.%Y"),  # Дата выхода
                row_list[5].strftime("%H:%M:%S"),  # Время выхода
                row_list[6],  # График
            ]

            # Путь к файлу с графиком
            graph_file_path = os.path.join(output_dir, row_list[-1] + ".xlsx")

            print(f"Обрабатываемая строка: {formatted_row}")

            # Проверка наличия файла с графиком
            if os.path.exists(graph_file_path):
                # Загрузка данных из файла с графиком
                graph_data = pd.read_excel(graph_file_path, engine='openpyxl')

                # Преобразование даты в нужный формат для сравнения
                formatted_date = row_list[2].strftime("%d.%m.%Y")

                # Поиск строки по значению второго индекса из строки daily_entries
                matching_row = graph_data[graph_data.iloc[:, 0] == formatted_date]

                # индекс строки из graph_data
                if not matching_row.empty:

                    if row_list[6] in ['График 98 бригада 1', 'График 98 бригада 2']:
                        matching_time = matching_row.iloc[0, 1]

                        if pd.notna(matching_time) and formatted_row[3] <= matching_time:
                            comment = f"Вход вовремя"
                            row_list[3] = matching_row.iloc[0, 1]
                            write_comment_to_excel(comment, formatted_row, comment_file_path, row_list)
                        elif pd.isna(matching_time):
                            comment = "Выход в выходной день"
                            write_comment_to_excel(comment, formatted_row, comment_file_path, row_list)
                        else:
                            comment = f"Проверить вход"
                            write_comment_to_excel(comment, formatted_row, comment_file_path, row_list)

                    if row_list[6] in ['График 1 бригада 1', 'График 1 бригада 2', 'График 1 бригада 3','График 1 бригада 4',
                                       'График 2 бригада 1', 'График 2 бригада 2', 'График 2 бригада 3','График 2 бригада 4',
                                       'График 3 бригада 1', 'График 3 бригада 2', 'График 3 бригада 3','График 3 бригада 4',
                                       'График 5 бригада 1', 'График 5 бригада 2', 'График 5 бригада 3', 'График 97 бригада 1']:
                        matching_time = matching_row.iloc[0, 1]

                        if pd.notna(matching_time) and formatted_row[3] <= matching_time:
                            comment = f"Вход вовремя"
                            row_list[3] = matching_row.iloc[0, 1]
                            write_comment_to_excel(comment, formatted_row, comment_file_path, row_list)
                        elif pd.isna(matching_time):
                            comment = "Выход в выходной день"
                            write_comment_to_excel(comment, formatted_row, comment_file_path, row_list)
                        else:
                            comment = f"Проверить вход"
                            write_comment_to_excel(comment, formatted_row, comment_file_path, row_list)
                    # Реализовать сравнение времени
                    # if not formatted_row[3] == matching_row.iloc[0, 1]:
                    #     print("Различие времени.")
                    #     # Написать комментарий в новом excel файле
                    #     comment = f"Различие времени: {formatted_row[3]} вместо {matching_row.iloc[0, 1]}"
                    #     write_comment_to_excel(comment, formatted_row, comment_file_path)
                else:
                    print("Не найдена соответствующая строка в графике.")
                    # Написать комментарий в новом excel файле
                    comment = f"График не содержит информации для {row_list[2]}"
                    write_comment_to_excel(comment, formatted_row, comment_file_path, row_list)
            else:
                print("Файл графика не найден.")
                # Написать комментарий в новом excel файле
                comment = f"Файл графика не найден для {row_list[-1]}"
                write_comment_to_excel(comment, formatted_row, comment_file_path, row_list)

            # Условие для остановки после обработки 10 строк
            if index + 1 >= limit:
                print(f"Обработано {limit} строк. Завершаем выполнение.")
                break

    except Exception as e:
        print(f"Произошла ошибка: {e}")

def write_comment_to_excel(comment, row_list, comment_file_path, formatted_row):
    try:
        # Загрузка существующего файла
        xls = openpyxl.load_workbook(comment_file_path)

        # Выбор нужного листа
        sheet = xls['Sheet1']

        # Определение следующей свободной строки
        next_row = sheet.max_row + 1

        # Обновление значения в row_list[3] для текущей строки
        row_list[3] = sheet.cell(row=next_row, column=4).value

        # Запись данных в новую строку
        formatted_row[2] = formatted_row[2].strftime("%d.%m.%Y")  # Преобразование формата даты
        formatted_row[4] = formatted_row[4].strftime("%d.%m.%Y")
        for col_num, value in enumerate([comment] + formatted_row, 1):
        # for col_num, value in enumerate([comment] + row_list[1:], 1):
            sheet.cell(row=next_row, column=col_num, value=value)

        # Сохранение изменений
        xls.save(comment_file_path)
        xls.close()  # Добавьте эту строку для корректного закрытия файла после сохранения



    except Exception as e:
        print(f"Ошибка при записи комментария в файл: {e}")

# Вызов функции в main.py
# from work_time_utils import process_daily_entries

def main():
    process_daily_entries(limit=1000)

if __name__ == "__main__":
    main()
