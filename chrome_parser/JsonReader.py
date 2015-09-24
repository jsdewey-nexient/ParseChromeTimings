# Wrapper for Python's JSON library to handle the HTTP Archive (.har) files
# exported by Chrome.

import json
import re
from sys import exit

class JsonReader:
    json_data = False
    entries = False

    def __init__(self, file_path):
        self.open(file_path)

    def open(self, file_path):
        try:
            with open(file_path, 'r') as fp:
                self.load(fp)
        except (OSError, IOError) as e:
            error_string = "Couldn't open the HAR file; Exception: {0}"\
                .format(e)
            print(error_string)
            exit()

    def load(self, fp):
        self.json_data = json.load(fp)
        self.json_data = self.json_data["log"]
        self.entries = self.json_data["entries"]

    def get_entire_page_timings(self, url=False):
        result = {}
        for page in self.json_data["pages"]:
            if url and url in page:
                return page["pageTimings"]
            result[page["title"]] = page["pageTimings"]
        return result

    def get_entry(self, url):
        for e in self.entries:
            if url in e["request"]:
                return e

    def get_entry_timings(self, entry):
        entry_url = self.format_url(entry["request"]["url"])
        result = {
            "url": entry_url,
            "total": entry["time"],
            "timings": {}
        }
        timings = entry["timings"]

        for key in timings:
            time = timings[key]
            if time <= 0:
                result["timings"][key] = 0.0
            else:
                result["timings"][key] = "{0:.3f}".format(timings[key])

        return result

    def get_all_entry_timings(self):
        timing_list = []
        for entry in self.entries:
            result = self.get_entry_timings(entry)
            timing_list.append(result)
        return timing_list

    def format_url(self, url):
        extracted = re.findall("^[^?]+", url)
        return extracted[0]

    def get_start_datetime(self):
        return self.json_data["pages"][0]["startedDateTime"]
