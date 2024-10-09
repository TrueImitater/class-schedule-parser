import argparse


if __name__ == "__main__":
    parser = argparse.ArgumentParser()

    parser.add_argument("-i", "--input-file", type=str,
                        help="Путь до файла с расписанием", required=True)
    parser.add_argument("-g", "--groups", nargs="+", type=str,
                        help="Наименование групп, расписание которых необходимо распарсить", required=True)
    parser.add_argument("-o", "--output-dir", default="./test.docx", type=str,
                        help="Директория, в которую будут сохранены документы с расписаниями")
