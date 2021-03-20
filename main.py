import urllib
from urllib.request import urlopen
from urllib.parse import urljoin
from lxml.html import fromstring
import xlsxwriter
from itertools import groupby

ITEM_PATH = 'tr td'
ITEM_PATH1 = '.bigtext'
ITEM_PATH2 = '#abstract2'


def conect(url):
    req = urllib.request.Request(
        url,
        data=None,
        headers={
            'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_9_3) AppleWebKit/537.36 (KHTML, like Gecko) '
                          'Chrome/35.0.1916.47 Safari/537.36',
            'Cookie': "SCookieID=Cookies"
        }
    )

    f = urllib.request.urlopen(req)
    list_html = f.read().decode('utf-8')
    list_doc = fromstring(list_html)

    return list_doc


def parser_vacancies():
    url = 'https://elibrary.ru/itembox_items.asp?id=1102414'
    list_doc = conect(url)

    urls = []
    for elem in list_doc.cssselect(ITEM_PATH):
        try:
            a = elem.cssselect('a')[0]
        except IndexError:
            pass

        if 'https://elibrary.ru/item.asp?id=' == str(urljoin(url, a.get('href'))[:32]):
            urls.append(urljoin(url, a.get('href')))

    list_urls = [el for el, _ in groupby(urls)]

    sources = []
    for ind in range(0, 2):
        # Имя
        name = ''
        list_doc = conect(list_urls[ind])
        for elem in list_doc.cssselect(ITEM_PATH1):
            try:
                name = elem.cssselect('p')[0].text
            except ValueError:
                pass

        # Журнал
        magazine_list = []

        for elem in list_doc.cssselect(ITEM_PATH):
            try:
                for i in range(len(elem) + 1):
                    font = elem.cssselect('a')[i].text
                    magazine_list.append(font)
            except ValueError:
                pass

        magazine = magazine_list[74]

        # key
        keys = []
        for elem in list_doc.cssselect(ITEM_PATH):
            try:
                for i in range(len(elem) + 1):
                    a = elem.cssselect('a')[i]
                    if 'https://elibrary.ru/keyword_items.asp?id=' == str(urljoin(url, a.get('href'))[:41]):
                        keys.append(a.text)
            except ValueError:
                pass

        keys_list = [el for el, _ in groupby(keys)]

        # АННОТАЦИЯ
        annotation = ''
        for elem in list_doc.cssselect(ITEM_PATH2):
            try:
                annotation = elem.cssselect('p')[0].text
            except ValueError:
                pass

        # Кол-во страниц
        number_of_pages = []
        for elem in list_doc.cssselect(ITEM_PATH):
            try:
                for i in range(len(elem) + 1):
                    font = elem.cssselect('font')[i].text
                    number_of_pages.append(font)
            except ValueError:
                pass

        year = number_of_pages[99]

        if ind == 0:
            number_of_pages = number_of_pages[100].split('-')
        else:
            number_of_pages = number_of_pages[102].split('-')

        number_of_pages = int(number_of_pages[1]) + 1 - int(number_of_pages[0])

        source = {'name': name,
                  'magazine': magazine,
                  'keys_list': str(keys_list),
                  'annotation': annotation,
                  'number_of_pages': number_of_pages,
                  'year': year}
        sources.append(source)
    return sources


def export_excel(filename, vacancies):
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': True})
    field_names = ('Тема', 'Год', 'Ключивые слова (Key words)', 'Аннотация', 'Страницы', 'Журнал')
    for i, field in enumerate(field_names):
        worksheet.write(0, i, field, bold)
    fields = ('name', 'year', 'keys_list', 'annotation', 'number_of_pages', 'magazine')
    for row, vacancy in enumerate(vacancies, start=1):
        for col, field in enumerate(fields):
            worksheet.write(row, col, vacancy[field])

    workbook.close()


def main():
    export_excel(r'F:\sources.xlsx', parser_vacancies())


if __name__ == '__main__':
    main()
