import pandas as pd
import openpyxl
import os
from tqdm import tqdm


def process_daily_entries(file_path="data/daily_entries.xlsx", output_dir="data/graph"):
    try:
        # Загрузка данных из файла daily_entries
        daily_entries = pd.read_excel(file_path, engine='openpyxl')

        # Создание директории output_files, если её нет
        # os.makedirs(output_dir, exist_ok=True)

        # Итерация по строкам из файла daily_entries
        for index, row in tqdm(daily_entries.iterrows(), total=len(daily_entries), desc="Processing Entries"):
            # Преобразование строки в список
            row_list = row.tolist()

            # Путь к файлу с графиком
            graph_file_path = os.path.join(output_dir, row_list[-1] + ".xlsx")

            print(f"Обрабатываемая строка: {row_list}")
            print(f"Путь к файлу графика: {graph_file_path}")

            # # Проверка наличия файла с графиком
            # if os.path.exists(graph_file_path):
            #     print("Файл графика существует.")
            #
            #     # Загрузка данных из файла с графиком
            #     graph_data = pd.read_excel(graph_file_path, engine='openpyxl')
            #
            #     # Поиск строки по значению второго индекса из строки daily_entries
            #     matching_row = graph_data[graph_data.iloc[:, 0] == row_list[2]]
            #
            #     if not matching_row.empty:
            #         print("Найдена соответствующая строка в графике.")
            #
            #         # Сравнение времени
            #         if matching_row.iloc[0, 0] != row_list[3]:
            #             print("Различие времени.")
            #             # Написать комментарий в новом excel файле
            #             comment = f"Различие времени: {matching_row.iloc[0, 0]} вместо {row_list[3]}"
            #             write_comment_to_excel(comment, row_list)
            #     else:
            #         print("Не найдена соответствующая строка в графике.")
            #         # Написать комментарий в новом excel файле
            #         comment = f"График не содержит информацию для {row_list[2]}"
            #         write_comment_to_excel(comment, row_list)
            # else:
            #     print("Файл графика не найден.")
            #     # Написать комментарий в новом excel файле
            #     comment = f"Файл графика не найден для {row_list[-1]}"
            #     write_comment_to_excel(comment, row_list)
    except Exception as e:
        print(f"Ошибка: {e}")


# def write_comment_to_excel(comment, row_list):
#     try:
#         # Создание нового excel файла с комментариями
#         comment_file_path = "output_files/comments.xlsx"
#         if not os.path.exists(comment_file_path):
#             df = pd.DataFrame(
#                 columns=["Комментарий", "Табельный", "ФИО", "Дата входа", "Время входа", "Дата выхода", "Время выхода",
#                          "График"])
#             df.to_excel(comment_file_path, index=False, engine='openpyxl')
#
#         # Добавление строки с комментарием
#         comment_df = pd.DataFrame([[comment] + row_list],
#                                   columns=["Комментарий", "Табельный", "ФИО", "Дата входа", "Время входа",
#                                            "Дата выхода", "Время выхода", "График"])
#         comment_df.to_excel(comment_file_path, index=False, header=False, mode='a', engine='openpyxl')
#     except Exception as e:
#         print(f"Ошибка при записи комментария в файл: {e}")


# Вызов функции в main.py
from work_time_utils import process_daily_entries


def main():
    process_daily_entries()


if __name__ == "__main__":
    main()
