# Wrapper for the Python-Excel library (xlwt).

import xlwt
from sys import exit

class ExcelWriter:
    workbook = False
    workbook_name = 'NewWorkBook.xls'
    sheets = []

    def __init__(self, workbook_name=False, first_sheet_name=False):
        if workbook_name:
            self.workbook_name = workbook_name
        if first_sheet_name:
            self.create_sheet(first_sheet_name)
        else:
            default_sheet_name = self.get_default_sheet_name()
            self.create_sheet(default_sheet_name)

    def create_workbook(self):
        self.workbook = xlwt.Workbook()

    def save_workbook(self):
        self.workbook.save(self.workbook_name)

    def create_sheet(self, sheet_name):
        sheet = self.workbook.add_sheet(sheet_name)
        self.sheets.append(sheet)

    def get_default_sheet_name(self):
        # Add one so we don't make a "New Sheet 0"
        number_of_sheets = len(self.sheets) + 1
        name = "New Sheet {0}".format(number_of_sheets)
        return name

    def write_cell(
            self,
            sheet_number,
            row_number,
            cell_number,
            content,
            style=''
    ):
        self.sheets[sheet_number].write(row_number, cell_number, content, style)

    def create_header_row(self, sheet_number, headers, row_number=0):
        style = xlwt.easyxf('bold on; underline: on; horiz center')
        number_of_headers = len(headers)
        for i, header in range(0, number_of_headers), headers:
            self.write_cell(sheet_number, row_number, i, header, style)
