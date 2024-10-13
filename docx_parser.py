from docx.table import Table
from docx.table import _Row, _Rows

from parser_utils import *


def parse_table(table_rows: _Rows, day_col_idx: int, time_col_idx: int, group_col_idx: int) -> dict:
    """ Парсинг таблицы - создание словаря с парами по дням недели """

    cur_day = table_rows[0].cells[day_col_idx].text

    schedule_dict = {}
    schedule_dict[cur_day] = []

    for row in table_rows:
        if cur_day != row.cells[day_col_idx].text.strip():
            schedule_dict[cur_day].append([END_OF_DAY, END_OF_DAY])
            cur_day = row.cells[day_col_idx].text.strip()
            schedule_dict[cur_day] = []

        schedule_dict[cur_day].append(
            [row.cells[time_col_idx].text.strip(), row.cells[group_col_idx].text.strip()])

    schedule_dict[cur_day].append([END_OF_DAY, END_OF_DAY])
    return schedule_dict
