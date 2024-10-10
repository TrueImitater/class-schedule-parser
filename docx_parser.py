from docx.table import Table
from docx.table import _Row, _Rows

from parser_utils import *


def find_group_row_idx(table: Table) -> int:
    """ Поиск индекса первой колонки в которой упомянаются группы """

    for (idx, row) in enumerate(table.rows):
        if row.cells[0].text.lower() == GROUP_KEY_WORD:
            return idx


def get_groups_list(group_row: _Row) -> dict[str, int]:
    """ Получение словаря вида наименование группы/номер колонки """
    group_dict = {}

    for (idx, cell) in enumerate(group_row.cells):
        if cell.text.lower() != GROUP_KEY_WORD:
            group_dict[cell.text] = idx

    return group_dict


def get_time_col_idx(table: Table) -> int:
    """ Получение индекса колонки времени пары """

    for (idx, row) in enumerate(table.rows):
        if row.cells[0].text.lower() == DAY_OF_WEEK_KEY_WORD:
            for (cell_idx, cell) in enumerate(row.cells):
                if cell.text.lower() == TIME_KEY_WORD:
                    return cell_idx

    return -1


def get_start_table_row_idx(table: Table) -> int:
    """ Получение индекса строки начала полезной информации в таблице """

    start_row_idx = -1

    for (idx, row) in enumerate(table.rows):
        if row.cells[0].text.lower() in DAYS_OF_WEEK_LIST_KEY_WORDS:
            return idx
