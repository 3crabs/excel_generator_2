from openpyxl import load_workbook


def test_1():
    wb = load_workbook("../resources/template.xlsx")
    for sheet_name in wb.sheetnames:
        print(sheet_name)


def test_2():
    read_book = load_workbook("/home/vladimir/Проекты/Генератор файлов/Задание/Копия реестр 2021.xlsx")
    for sheet_name in read_book.sheetnames:
        print(sheet_name)


if __name__ == '__main__':
    # test_1()
    test_2()
