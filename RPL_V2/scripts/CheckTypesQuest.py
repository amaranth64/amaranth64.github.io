import pandas as pd
import os

# Получаем текущую директорию
current_dir = os.getcwd()

for filename in os.listdir(current_dir):

    if filename.endswith('.xlsx'):

        # Загрузить файл Excel
        sheets = pd.read_excel(filename, sheet_name=None)

        # Определить функции для проверки условий
        def determine_value(row):
            c_val = row[2]
            ans = row[6]
            other_vals = [row[col] for col in [5, 7, 9, 11, 13, 15]]

            if ans == 'Да' or ans == 'Нет' or row[8] == row[10]:
                return 6
            elif c_val != 'no_pic' and all(val == 'no_pic' for val in other_vals):
                return 1
            elif all(str(val).startswith('num_') for val in other_vals):
                return 4
            elif c_val != 'no_pic' and all(val != 'no_pic' for val in other_vals):
                return 2
            elif c_val == 'no_pic' and all(val == 'no_pic' for val in other_vals):
                return 3
            elif c_val == 'no_pic' and all(val != 'no_pic' for val in other_vals):
                return 5
            else:
                return None

        # Обновить столбец A или проверить существующие значения
        def update_or_check_value(row):
            expected_value = determine_value(row)
            current_value = row[0]

            if pd.isna(current_value):  # Если значение в A пустое, заполняем
                return expected_value
            else:  # Если значение в A не пустое, проверяем правильность
                if current_value == expected_value:
                    return current_value  # Если значение правильное, оставляем его
                else:
                    print(f"Файл {filename}. Неверное значение в строке {row.name + 2}: ожидалось {expected_value}, но найдено {current_value}")
                    return expected_value  # Возвращаем текущее значение, но можно добавить логику для исправления


        for sheet_name, df in sheets.items():
            df['type'] = df.apply(update_or_check_value, axis=1)
            sheets[sheet_name] = df

        # Сохранить обновленный файл Excel с изменениями на всех листах
        with pd.ExcelWriter(filename) as writer:
            for sheet_name, df in sheets.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)

