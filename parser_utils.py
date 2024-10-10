from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

GROUP_KEY_WORD = "группа"
DAY_OF_WEEK_KEY_WORD = "день недели"
TIME_KEY_WORD = "время"

DAYS_OF_WEEK_LIST_KEY_WORDS = [
    "понедельник", "вторник", "среда", "четверг", "пятница", "суббота"]


TEXT_FONT_KEY_WORD = "Times New Roman"

DEFAULT_CELL_COLOR = "#e1ffe1"

UPPER_HEAD_CELL_COLOR = "#ff9999"
UPPER_CELL_COLOR = "#ffcccc"

LOWER_HEAD_CELL_COLOR = "#99ccff"
LOWER_CELL_COLOR = "#ccecff"


def set_cell_background_color(cell, color):
    """ Функция для изменения цвета фона ячейки """
    # Получаем доступ к XML ячейки
    cell_properties = cell._element.get_or_add_tcPr()
    # Создаем элемент для заливки цвета
    shading = OxmlElement('w:shd')
    shading.set(qn('w:fill'), color)
    # Добавляем элемент к свойствам ячейки
    cell_properties.append(shading)
