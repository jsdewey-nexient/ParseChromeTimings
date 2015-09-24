# Where the magic happens...

from chrome_parser.JsonReader import JsonReader
from excel_writer.ExcelWriter import ExcelWriter
from pprint import pprint

# Settings
CHROME_COLUMN_NAMES = [  # For Chrome as Chrome includes more information.
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

IE_COLUMN_NAMES = [  # For IE.
    "URL",
    "Connect (ms)",
    "Send (ms)",
    "Wait (ms)",
    "Receive (ms)",
    "Total (ms)"
]

HAR_FILE_PATH = "./chrome2.har"

if __name__ == "__main__":
    chrome = JsonReader(HAR_FILE_PATH)
    pprint(chrome.get_start_time())
