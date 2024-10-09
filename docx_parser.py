from docx.table import Table
from docx.table import _Row

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
