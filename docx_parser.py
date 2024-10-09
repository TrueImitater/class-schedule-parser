from docx.table import Table
from docx.table import _Row

from typing import Optional

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


def get_some_setting_and_table_start_idx(table: Table) -> tuple[Optional[int], Optional[int], Optional[int]]:
    """ получение индекса колонки дня недели, времени пары, а также
     индекса строки с первой парой """

    day_col_idx = time_col_idx = -1

    for (idx, row) in enumerate(table.rows):
        if row.cells[0].text.lower() == DAY_OF_WEEK_KEY_WORD:
            day_col_idx = 0
            for (cell_idx, cell) in enumerate(row.cells):
                if cell.text.lower() == TIME_KEY_WORD:
                    time_col_idx = cell_idx

    if day_col_idx > -1:
        for (idx, row) in enumerate(table.rows[day_col_idx + 1:]):
            if row.cells[0].text.lower() in DAYS_OF_WEEK_LIST_KEY_WORDS:
                return 0, time_col_idx, idx

    return None, None, None
