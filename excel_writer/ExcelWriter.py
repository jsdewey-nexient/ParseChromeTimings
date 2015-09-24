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
        self.create_workbook()
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
        style_object = xlwt.easyxf(style)
        self.sheets[sheet_number].write(
            row_number,
            cell_number,
            content,
            style_object
        )

    def write_row(self, sheet_number, row_number, content, style):
        for c in content:
            for info, i in c, range(0, len(info)):
                self.write_cell(sheet_number, row_number, i, info)

    def create_header_row(self, sheet_number, headers, row_number=0):
        number_of_headers = len(headers)
        for i, header in range(0, number_of_headers), headers:
            self.write_cell(
                sheet_number,
                row_number,
                i,
                header,
                'font: bold on; font: underline on; alignment: horiz center'
            )

    def create_date_row(self, sheet_number, date, row_number=0):
        self.write_cell(sheet_number, row_number, 0, "DATE:", "font: bold on")
        self.write_cell(sheet_number, row_number, 1, date)
