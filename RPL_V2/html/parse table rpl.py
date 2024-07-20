import requests
from bs4 import BeautifulSoup


def parse(file, url, tag, clas=False):
    # URL сайта
    url = f"https://soccer365.ru/competitions/{url}"

    # Путь к локальному HTML-файлу
    local_html_path = file

    # Отправка GET-запроса на сайт
    response = requests.get(url)

    # Проверка статуса запроса
    if response.status_code == 200:
        # Парсинг HTML-кода страницы
        soup = BeautifulSoup(response.text, 'html.parser')

        # Поиск тега div с id "competition_table"

        if clas:
            competition_table = soup.find('div', class_=tag)
        else:
            competition_table = soup.find('div', id=tag)
        # Проверка наличия искомого тега
        if competition_table:
            # Содержимое тега competition_table
            new_content = competition_table.prettify()

            # Чтение локального HTML-файла
            with open(local_html_path, 'r', encoding='utf-8') as file:
                local_soup = BeautifulSoup(file, 'html.parser')

            # Поиск тега div с id "competition_table" в локальном файле
            if clas:
               local_competition_table = local_soup.find('div', class_=tag)
            else:
                local_competition_table = local_soup.find('div', id=tag)

            if local_competition_table:
                # Замена содержимого тега
                local_competition_table.replace_with(BeautifulSoup(new_content, 'html.parser'))

                # Запись обновленного HTML-кода обратно в файл
                with open(local_html_path, 'w', encoding='utf-8') as file:
                    file.write(str(local_soup))

                print(f"Содержимое тега '{tag}' успешно обновлено в локальном файле.")
            else:
                print(f"Тег с id '{tag}' не найден в локальном файле.")
        else:
            print(f"Тег с id '{tag}' не найден на сайте.")
    else:
        print(f"Ошибка при запросе страницы: {response.status_code}")


def delete_a(local_html_path):

    # Чтение локального HTML-файла
    with open(local_html_path, 'r', encoding='utf-8') as file:
        local_soup = BeautifulSoup(file, 'html.parser')

    # Поиск всех тегов <a> и удаление атрибута href
    for a_tag in local_soup.find_all('a'):
        a_tag.attrs.pop('href', None)

    # Запись обновленного HTML-кода обратно в файл
    with open(local_html_path, 'w', encoding='utf-8') as file:
        file.write(str(local_soup))

def remove_div_by_class(html_file_path, class_name):
    """
    Удаляет теги <div> с указанным классом из HTML-файла.

    :param html_file_path: Путь к исходному HTML-файлу.
    :param class_name: Класс, по которому будут удалены теги <div>.
    """
    # Чтение локального HTML-файла
    with open(html_file_path, 'r', encoding='utf-8') as file:
        soup = BeautifulSoup(file, 'html.parser')

    # Поиск и удаление всех тегов <div> с указанным классом
    for div in soup.find_all('div', class_=class_name):
        div.decompose()

    # Запись обновленного HTML-кода обратно в файл
    with open(html_file_path, 'w', encoding='utf-8') as file:
        file.write(str(soup))

    print(f"Теги <div> с классом '{class_name}' успешно удалены.")


parse('tableRPL.html', '13/', 'competition_table')
parse('tableFNL.html', '687/', 'competition_table')
parse('cupRussia.html', '786/', 'page_main_content', clas=True)
#parse('Статистика ФНЛ.html', '687/players/', 'mrgt5', clas=True)

with open('calendarRPL.html', 'w') as outfile:
    parse('itogRPL.html', '13/results/', 'result_data')
    with open('itogRPL.html') as infile:
        outfile.write(infile.read())

    parse('itogRPL.html', '13/shedule/', 'result_data')
    with open('itogRPL.html') as infile:
        outfile.write(infile.read())

delete_a('tableRPL.html')
delete_a('tableFNL.html')
delete_a('cupRussia.html')
#delete_a('Статистика ФНЛ.html')

remove_div_by_class('calendarRPL.html', 'icons')
remove_div_by_class('cupRussia.html', 'icons')
#remove_div_by_class('Статистика ФНЛ.html', 'pager')

with open('calendarRPL.html', 'r', encoding='utf-8') as outfile:
    d = outfile.read()
    d = d.replace('Календарь РПЛ', 'Результаты', 1)


with open('calendarRPL.html', 'w', encoding='utf-8') as outfile:
    outfile.write(d)
