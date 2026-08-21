import os
from title_work_automator import AutomatorApp

class MockApp:
    def __init__(self):
        self.viewer_pid_dir = "/Volumes/davidlls/assignments/PID 42-01011.000"
    def log(self, msg):
        print("[LOG]", msg)
    def _fetch_kofile_name_search(self, search_params, pid_dir):
        # We can just steal the method from the class
        AutomatorApp._fetch_kofile_name_search(self, search_params, pid_dir)

app = MockApp()
search_params = [
    {"name": "Phillips Ross", "acquisition_date": "", "disposal_date": "10/13/1974"},
    {"name": "Vanfossen Patricia", "acquisition_date": "03/12/2023", "disposal_date": ""}
]
app._fetch_kofile_name_search(search_params, "/Volumes/davidlls/assignments/PID 42-01011.000")
