# Wrapper for the Python-Excel library (xlwt).

import xlwt
from time import strptime

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
            style='',
            num_format=''
    ):
        style_object = xlwt.easyxf(style, num_format)
        self.sheets[sheet_number].write(
            row_number,
            cell_number,
            content,
            style_object
        )

    def write_entry_chrome(self, sheet_number, entry, row_number=0):
        timings = entry["timings"]
        num_style = "0.000"
        self.write_cell(sheet_number, row_number, 0, entry["url"], "font: bold on")

        for key, cell in list(zip(timings, range(1, len(timings) + 1))):
            self.write_cell(
                sheet_number,
                row_number,
                cell,
                float(timings[key]),
                '',
                num_style
            )

        self.write_cell(sheet_number, row_number, 8, entry["total"], '', num_style)

        return row_number + 1

    def create_header_row(self, sheet_number, headers, row_number=0):
        number_of_headers = len(headers)
        for i, header in list(zip(range(0, number_of_headers), headers)):
            self.write_cell(
                sheet_number,
                row_number,
                i,
                header,
                'font: bold on; font: underline on; alignment: horiz center'
            )
        return row_number + 1

    def create_date_row(self, sheet_number, date, row_number=0):
        split_date = str.split(date, ".")
        date_struct = strptime(split_date[0], "%Y-%m-%dT%H:%M:%S")
        new_date = "{0}/{1}/{2} at {3}:{4}:{5}".format(
            date_struct[1],
            date_struct[2],
            date_struct[0],
            date_struct[3],
            date_struct[4],
            date_struct[5]
        )
        self.write_cell(sheet_number, row_number, 0, "DATE:", "font: bold on; align: horiz right")
        self.write_cell(sheet_number, row_number, 1, new_date)
        return row_number + 1

    def create_timings_row(self, sheet_number, url, timing_dictionary, row_number=0):
        style = "font: bold on; align: horiz right"
        num_style = "0.000"

        self.write_cell(sheet_number, row_number, 0, "URL:", style)
        self.write_cell(sheet_number, row_number, 1, url)
        row_number += 1

        self.write_cell(sheet_number, row_number, 0, "Content Load (ms):", style)
        self.write_cell(
            sheet_number,
            row_number,
            1,
            timing_dictionary["onContentLoad"],
            '',
            num_style
        )
        row_number += 1

        self.write_cell(sheet_number, row_number, 0, "Load (ms):", style)
        self.write_cell(
            sheet_number,
            row_number,
            1,
            timing_dictionary["onLoad"],
            '',
            num_style
        )
        row_number += 2

        return row_number
