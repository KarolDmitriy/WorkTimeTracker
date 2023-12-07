import pandas as pd
import openpyxl
import os
from tqdm import tqdm

def process_daily_entries(file_path="data/daily_entries.xlsx", output_dir="data/graph", limit=10):
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
                    print("Найдена соответствующая строка в графике.")

                    # Сравнение времени
                    if not matching_row.iloc[0, 0] == formatted_row[3]:
                        print("Различие времени.")
                        # Написать комментарий в новом excel файле
                        comment = f"Различие времени: {matching_row.iloc[0, 0]} вместо {formatted_row[3]}"
                        write_comment_to_excel(comment, formatted_row, comment_file_path)
                else:
                    print("Не найдена соответствующая строка в графике.")
                    # Написать комментарий в новом excel файле
                    comment = f"График не содержит информации для {row_list[2]}"
                    write_comment_to_excel(comment, formatted_row, comment_file_path)
            else:
                print("Файл графика не найден.")
                # Написать комментарий в новом excel файле
                comment = f"Файл графика не найден для {row_list[-1]}"
                write_comment_to_excel(comment, formatted_row, comment_file_path)

            # Условие для остановки после обработки 10 строк
            if index + 1 >= limit:
                print(f"Обработано {limit} строк. Завершаем выполнение.")
                break

    except Exception as e:
        print(f"Произошла ошибка: {e}")

def write_comment_to_excel(comment, row_list, comment_file_path):
    try:
        # Загрузка существующего файла
        xls = openpyxl.load_workbook(comment_file_path)

        # Выбор нужного листа
        sheet = xls['Sheet1']

        # Определение следующей свободной строки
        next_row = sheet.max_row + 1

        # Запись данных в новую строку
        for col_num, value in enumerate([comment] + row_list, 1):
            sheet.cell(row=next_row, column=col_num, value=value)

        # Сохранение изменений
        xls.save(comment_file_path)

    except Exception as e:
        print(f"Ошибка при записи комментария в файл: {e}")

# Вызов функции в main.py
# from work_time_utils import process_daily_entries

def main():
    process_daily_entries(limit=10)

if __name__ == "__main__":
    main()
