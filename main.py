import argparse
import docx

from docx_parser import parse_table
from parser_utils import TableCoreParams

if __name__ == "__main__":
    parser = argparse.ArgumentParser()

    parser.add_argument("-i", "--input-file", type=str,
                        help="Путь до файла с расписанием", required=True)
    parser.add_argument("-g", "--groups", nargs="+", type=str,
                        help="Наименование групп, расписание которых необходимо распарсить", required=True)
    parser.add_argument("-o", "--output-dir", default="./test.docx", type=str,
                        help="Директория, в которую будут сохранены документы с расписаниями")

    # парсим параметры из командной строки
    input_params = parser.parse_args()

    # пробуем прочитать файл
    docx_file = docx.Document(input_params.input_file)

    # создаем объект, содержащий основные параметры по анализируемой таблице
    table_core_params = TableCoreParams(docx_file.tables[0])
    print(table_core_params)

    # для каждой из указанной пользователем группы создаем таблицу
    for group_name in input_params.groups:
        # если группа присутствует в загруженном файле
        if table_core_params.group_name_indexes.get(group_name, None) is not None:
            parse_table(table_core_params.table.rows, table_core_params.day_col_idx,
                        table_core_params.time_col_idx, table_core_params.group_name_indexes[group_name])
