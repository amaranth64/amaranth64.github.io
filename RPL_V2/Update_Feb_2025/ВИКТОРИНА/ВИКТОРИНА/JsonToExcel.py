import os
import json
import openpyxl

# Получаем текущую директорию
current_dir = os.getcwd()

# Проходим по всем файлам в директории
for filename in os.listdir(current_dir):
    # Проверяем, является ли файл JSON-файлом
    if filename.endswith('.json'):
        # Открываем JSON-файл для чтения
        with open(filename, 'r', encoding='utf-8') as file:
            data = json.load(file)

        workbook = openpyxl.Workbook()
        worksheet = workbook.active

        # Получаем названия параметров из первого элемента
        headers = list(data[0].keys())

        # Записываем названия параметров в первую строку
        worksheet.append(headers)

        # Записываем значения параметров в последующие строки
        for item in data:
            row = [item.get(header, '') for header in headers]
            worksheet.append(row)

        # Создаем имя для нового JSON-файла
        new_filename = os.path.splitext(filename)[0]

        # Открываем новый JSON-файл для записи
        workbook.save(f'{new_filename}.xlsx')

        print(f'Файл {new_filename} создан.')
