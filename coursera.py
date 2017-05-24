import requests
import lxml
from bs4 import BeautifulSoup
from openpyxl import Workbook
from lxml import html


def get_courses_list(link):
    page = requests.get(link)
    tree = html.fromstring(page.content)
    courses_list = tree.xpath('//loc/text()')
    return courses_list


def get_course_info(course_slug, quantity_courses):
    list_info_courses = []
    for link in course_slug[:quantity_courses]:
        page = requests.get(link).text
        soup = BeautifulSoup(page, 'html.parser')
        name_course = soup.find('h1', class_='title display-3-text').text
        start_date_course = soup.find(
                'div', class_='startdate rc-StartDateString caption-text').text
        lang_course = soup.find('div', class_='rc-Language').text
        count_week_course = soup.find_all('div', class_='week')
        rating_course = soup.find(
                'div', class_='ratings-text bt3-visible-xs')
        if not rating_course == None:
            rating_course = rating_course.text
        else:
            rating_course = 'Рейтинг не указан'
        list_info_courses.append(
                            [name_course,
                                lang_course,
                                start_date_course,
                                len(count_week_course),
                                rating_course])
    return list_info_courses


def output_courses_info_to_xlsx(filepath, list_info_courses):
    wb = Workbook()
    ws = wb.active
    for r,  line in enumerate(list_info_courses):
        for c, value in enumerate(line):
            ws.cell(row=r+1, column=c+1).value = value
    wb.save(filepath+'.xlsx')


if __name__ == '__main__':
    quantity_courses = int(
        input('Укажите какое количество курсов вам необходимо: '))
    filepath = input('Укажите имя файла: ')
    courses_list = get_courses_list(
        'https://www.coursera.org/sitemap~www~courses.xml')
    list_info_courses = get_course_info(courses_list, quantity_courses)
    output_courses_info_to_xlsx(filepath, list_info_courses)
