""" Функции для оформления таблиц """

import docx

import docx.enum
import docx.enum.table
import docx.enum.text
import docx.shared
import docx.table
from docx.shared import Inches, Cm

from docx import Document
from parser_utils import *

TEXT_FONT_KEY_WORD = "Times New Roman"

DEFAULT_CELL_COLOR = "#e1ffe1"

UPPER_HEAD_CELL_COLOR = "#ff9999"
UPPER_CELL_COLOR = "#ffcccc"

LOWER_HEAD_CELL_COLOR = "#99ccff"
LOWER_CELL_COLOR = "#ccecff"


def change_cell_style(cur_cell: docx.table._Cell, font_name: str,
                      cell_size: int):
    """ Изменение общих характеристик ячейки """
    cur_cell.vertical_alignment = docx.enum.table.WD_ALIGN_VERTICAL.CENTER
    cur_cell.paragraphs[0].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER

    cur_cell_font = cur_cell.paragraphs[0].runs[0].font
    cur_cell_font.name = font_name
    cur_cell_font.size = docx.shared.Pt(cell_size)


def set_cell_background_color(cell, color):
    """ Функция для изменения цвета фона ячейки """
    # Получаем доступ к XML ячейки
    cell_properties = cell._element.get_or_add_tcPr()
    # Создаем элемент для заливки цвета
    shading = OxmlElement('w:shd')
    shading.set(qn('w:fill'), color)
    # Добавляем элемент к свойствам ячейки
    cell_properties.append(shading)


def create_table_head(doc: docx.Document):
    """ Создание шапки таблицы с распершеннным расписанием """

    # создаем экземпляр таблицы
    schedule_table = doc.add_table(rows=2, cols=4)
    schedule_table.style = 'Table Grid'

    schedule_table.cell(0, 0).text = "Группа"
    schedule_table.cell(0, 0).merge(schedule_table.cell(0, 1))

    schedule_table.cell(1, 0).text = "день"
    schedule_table.cell(1, 1).text = "время"

    schedule_table.cell(0, 2).text = "Верхняя"
    schedule_table.cell(0, 2).merge(schedule_table.cell(1, 2))
    schedule_table.cell(0, 3).text = "Нижняя"
    schedule_table.cell(0, 3).merge(schedule_table.cell(1, 3))

    # оформляем ячейку "Группа"
    group_cell = schedule_table.cell(0, 0)
    group_cell.width = Cm(3.5)
    group_cell.paragraphs[0].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    change_cell_style(group_cell, TEXT_FONT_KEY_WORD, 12)
    set_cell_background_color(group_cell, DEFAULT_CELL_COLOR)

    # оформляем ячейку "День"
    day_of_week_cell = schedule_table.cell(1, 0)
    day_of_week_cell.width = Cm(2.5)
    day_of_week_cell.paragraphs[0].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    change_cell_style(day_of_week_cell, TEXT_FONT_KEY_WORD, 12)
    set_cell_background_color(day_of_week_cell, DEFAULT_CELL_COLOR)

    # оформляем ячейку "Время"
    time_cell = schedule_table.cell(1, 1)
    time_cell.width = Cm(1)
    time_cell.paragraphs[0].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    change_cell_style(time_cell, TEXT_FONT_KEY_WORD, 12)
    set_cell_background_color(time_cell, DEFAULT_CELL_COLOR)

    # оформляем ячейку "Верхняя"
    upper_cell = schedule_table.cell(0, 2)
    upper_cell.width = Cm(8.5)
    upper_cell.vertical_alignment = docx.enum.table.WD_ALIGN_VERTICAL.CENTER
    upper_cell.paragraphs[0].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER

    upper_cell_font = upper_cell.paragraphs[0].runs[0].font
    upper_cell_font.name = TEXT_FONT_KEY_WORD
    upper_cell_font.size = docx.shared.Pt(16)
    upper_cell_font.bold = True
    set_cell_background_color(upper_cell, UPPER_HEAD_CELL_COLOR)

    # оформляем ячейку "Нижняя"
    lower_cell = schedule_table.cell(0, 3)
    lower_cell.width = Cm(8.5)
    lower_cell.vertical_alignment = docx.enum.table.WD_ALIGN_VERTICAL.CENTER
    lower_cell.paragraphs[0].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER

    lower_cell_font = lower_cell.paragraphs[0].runs[0].font
    lower_cell_font.name = TEXT_FONT_KEY_WORD
    lower_cell_font.size = docx.shared.Pt(16)
    lower_cell_font.bold = True
    set_cell_background_color(lower_cell, LOWER_HEAD_CELL_COLOR)


def add_schedule_row(schedule_table: docx.table.Table, time: str, upper_schedule: str, lower_schedule: str):
    """ Добавление строки в таблицу """

    cells = schedule_table.add_row().cells

    # # оформляем ячейку "День"
    # change_cell_style(cells[0], TEXT_FONT_KEY_WORD, 10)
    set_cell_background_color(cells[0], DEFAULT_CELL_COLOR)

    # оформляем ячейку "Время"
    cells[1].text = time
    change_cell_style(cells[1], TEXT_FONT_KEY_WORD, 11)
    set_cell_background_color(cells[1], DEFAULT_CELL_COLOR)

    # оформляем ячейку "Верхней недели"
    cells[2].text = upper_schedule
    change_cell_style(cells[2], TEXT_FONT_KEY_WORD, 11)
    set_cell_background_color(
        cells[2], DEFAULT_CELL_COLOR if upper_schedule == "" else UPPER_CELL_COLOR)

    # оформляем ячейку "Нижней недели"
    cells[3].text = lower_schedule
    change_cell_style(cells[3], TEXT_FONT_KEY_WORD, 11)
    set_cell_background_color(
        cells[3], DEFAULT_CELL_COLOR if lower_schedule == "" else LOWER_CELL_COLOR)
