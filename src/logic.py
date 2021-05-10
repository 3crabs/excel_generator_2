from shutil import copyfile

import openpyxl
from openpyxl import load_workbook


def read_file(file_path: str) -> list:
    wb = load_workbook(file_path)
    sheet = wb.get_sheet_by_name(wb.sheetnames[0])
    return sheet.rows


m = {
    '1': [{'s': 'стр.1_2', 'c': 'AK9'}, {'s': 'стр.1_2', 'c': 'CM9'}],
    '2': [],
    '3': [],
    '4': [],
    '5': [],
    '6': [],
    '7': [],
    '8': [],
    '9': [],
    '10': [],
    '11': [],
    '12': [{'s': 'стр.1_2', 'c': 'G9'}],
    '13': [{'s': 'стр.1_2', 'c': 'BI9'}],
    '14': [],
    '15': [],
    '16': [],
}


def make_invoice_file(invoice: list, template_file_path: str, save_dir_path: str):
    if invoice[0].value is None:
        return
    new_file_path = save_dir_path + "/" + invoice[0].value + ".xlsx"
    copyfile(template_file_path, new_file_path)
    wb = openpyxl.load_workbook(filename=new_file_path)

    i = 1
    for cell in invoice:
        for place in m[str(i)]:
            wb[place['s']][place['c']] = cell.value
        i += 1

    wb.save(new_file_path)
