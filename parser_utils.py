from docx import Document
from docx.table import Table
from docx.table import _Row, _Rows
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

GROUP_KEY_WORD = "группа"
DAY_OF_WEEK_KEY_WORD = "день недели"
TIME_KEY_WORD = "время"

DAYS_OF_WEEK_LIST_KEY_WORDS = [
    "понедельник", "вторник", "среда", "четверг", "пятница", "суббота"]

END_OF_DAY = "STOP"


class TableCoreParams():
    """ Ключевые параметры по загруженной таблице """

    def __init__(self, table: Table):
        # экземпляр загруженной таблицы
        self.table = table
        # индекс строки с которой начинаются пары
        self.start_row_idx = 0
        # индекс колонки, в которой записан день недели
        self.day_col_idx = 0  # по умолчанию 0
        # индекс колонки, в которой записано время начала пары
        self.time_col_idx = 1  # по умолчанию 1
        # словарь соотношения индекса колонки к наименованию группы
        self.group_name_indexes = dict()

        # получаем список присутствующих в docx-файле групп
        self.set_groups_list()

        # устанавливаем индекс строки, с которой начинаются пары
        self.set_start_table_row_idx()
        if (self.start_row_idx < 0):
            raise Exception(
                "[ERROR] Не удается найти строку, с которой начинаются пары")

        # устанавливаем индекс колонки, в которой записано время начала пары
        self.set_time_col_idx()
        if (self.start_row_idx < 0):
            raise Exception(
                "[ERROR] Не удается индекс колонки, в которой записано время начала пары")

    def find_group_row_idx(self) -> int:
        """ Поиск индекса первой колонки в которой упомянаются группы """

        for (idx, row) in enumerate(self.table.rows):
            if row.cells[0].text.lower() == GROUP_KEY_WORD:
                return idx

    def set_groups_list(self) -> None:
        """ Заполнение словаря вида наименование группы/номер колонки """

        # получаем индекс строки наименованиями групп
        group_row_idx = self.find_group_row_idx()
        if group_row_idx < 0:
            raise Exception(
                "[ERROR] Не удается найти строку с наименованиями групп!")

        group_row = self.table.rows[group_row_idx]

        self.group_name_indexes = {}

        for (idx, cell) in enumerate(group_row.cells):
            if cell.text.lower() != GROUP_KEY_WORD:
                self.group_name_indexes[cell.text] = idx

    def set_start_table_row_idx(self) -> None:
        """ Получение индекса строки начала полезной информации в таблице """

        self.start_row_idx = -1

        for (idx, row) in enumerate(self.table.rows):
            if row.cells[0].text.lower() in DAYS_OF_WEEK_LIST_KEY_WORDS:
                self.start_row_idx = idx
                return

    def set_time_col_idx(self) -> None:
        """ Получение индекса колонки времени пары """

        self.time_col_idx = -1

        for (idx, row) in enumerate(self.table.rows):
            if row.cells[0].text.lower() == DAY_OF_WEEK_KEY_WORD:
                for (cell_idx, cell) in enumerate(row.cells):
                    if cell.text.lower() == TIME_KEY_WORD:
                        self.time_col_idx = cell_idx
                        return

    def __repr__(self):
        return f"TableCoreParams(start_row_idx = {self.start_row_idx}, day_col_idx = {self.day_col_idx}, time_col_idx = {self.time_col_idx}, group_name_indexes = {self.group_name_indexes})"
