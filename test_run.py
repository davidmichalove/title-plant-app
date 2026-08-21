from title_work_automator import AutomatorApp
import tkinter as tk

class DummyApp(AutomatorApp):
    def __init__(self):
        self.gdf = None
        self.req_headers = {
            'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        }
    def setup_ui(self):
        pass
    def log(self, msg):
        print("LOG:", msg)

app = DummyApp()
# Find a parcel in WAR
app.load_shapefile()
war_row = app.gdf[app.gdf['twp'].str.upper() == 'WAR'].iloc[0]
print("Testing WAR parcel:", war_row['parcel_no'])
app.run_process(war_row['parcel_no'])
