# Where the magic happens...

from chrome_parser.JsonReader import JsonReader
from excel_writer.ExcelWriter import ExcelWriter
import re

# Browser-Agnostic Settings
EXCEL_FILE_NAME = "Github.xlsx"

# Google Chrome Settings
HAR_FILE_PATH = "./github.com.har"
HAR_FILE_PATHS = [
    "./github1.har",
    "./github2.har",
    "./github3.har",
    "./github4.har",
    "./github5.har"
]
CHROME_COLUMN_NAMES = [
    "URL",
    "Blocked (ms)",
    "DNS (ms)",
    "Connect (ms)",
    "Send (ms)",
    "Wait (ms)",
    "Receive (ms)",
    "SSL (ms)",
    "Total (ms)"
]
# Internet Explorer Settings
XML_FILE_PATH = ""
IE_COLUMN_NAMES = [
    "URL",
    "Connect (ms)",
    "Send (ms)",
    "Wait (ms)",
    "Receive (ms)",
    "Total (ms)"
]

if __name__ == "__main__":
    excel = ExcelWriter(EXCEL_FILE_NAME)
    sheet = 0
    row = 0

    print("Starting...")

    for path in HAR_FILE_PATHS:
        chrome = JsonReader(path)

        #date = chrome.get_start_datetime()
        #excel.create_date_row(row, date)
        #row += 2  # Skip two rows before writing in more data.

        page_timings = chrome.get_entire_page_timings()
        for url in page_timings:
            row = excel.create_timings_row(sheet, url, page_timings[url], row)

        entries = chrome.get_all_entry_timings()
        row = excel.create_header_row(sheet, CHROME_COLUMN_NAMES, row)

        for entry in entries:
            row = excel.write_entry_chrome(sheet, entry, row)

        row += 3

    excel.save_workbook()
