import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import ttk
from pathlib import Path
import gspread
from google.oauth2.service_account import Credentials
import re
import datetime
import sys
from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer
import time

# Run everytime a file is created
class FolderHandler(FileSystemEventHandler):
    def on_created(self, event):
        if not event.is_directory:
            time.sleep(1)
            update_sheet(event.src_path, "created")


def resource_path(relative):
    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / relative
    return Path(__file__).parent / relative

def parse_url(link):
    sheet_id_match = re.search(r'/d/([a-zA-Z0-9-_]+)', link)
    sheet_id = sheet_id_match.group(1) if sheet_id_match else None

    gid_match = re.search(r'gid=([0-9]+)', link)
    gid = gid_match.group(1) if gid_match else "0"

    return sheet_id, gid


# Gspread API initialization
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
creds = Credentials.from_service_account_file(resource_path("data/sheets-automation-485701-ebb7ea60f622.json"), scopes=SCOPES)
client = gspread.authorize(creds)

#default_path = "C:/Users/iqbal/Downloads/"
default_path = "O:/"

#window pop-up setup
user_link = ''
def send():
    global user_link
    user_link = entry.get()
    window.destroy()

window = tk.Tk()
window.title("Select Option")
window.geometry("350x175")

default_URL = tk.StringVar(value='https://docs.google.com/spreadsheets/d/1mARf98z1tTqTimLuweBU10VGqXJilK9cpAYh2CchmDo/edit?gid=0#gid=0')

tk.Label(window, text="Enter spreadsheet URL:").pack(pady=10)
entry = ttk.Entry(window, textvariable=default_URL ,width=25)
entry.pack()
tk.Button(window, text="Send", command=send).pack(pady=15)
window.mainloop()


#Initialize the watchdog
observer = Observer()
observer.schedule(
    FolderHandler(),
    path=default_path,
    recursive=False
)
observer.start()


# Google Sheets setup
def update_sheet(path, event_type):
    # put 3 most recent report.xml files into options[]
    if event_type == "created":
        folder = Path(default_path)
        files = [
            f for f in folder.iterdir()
            if f.is_file() and "report" in f.name.lower()
        ]

        top_1 = sorted(files, key=lambda f: f.stat().st_mtime, reverse=True)[:1]

        index = 0
        for f in top_1:
            option = f.name
            print(option)
            index += 1


        print(f"{event_type}: {path}")
        id,gid = parse_url(user_link)
        URL = f'https://docs.google.com/spreadsheets/d/{id}/edit?gid={gid}#gid={gid}'
        print(URL)
        print(id,gid)
        month = datetime.datetime.now().strftime("%B")

        report = f'{default_path}{option}'
        COA = f'{option[0:13]}  .assemble.xml'
        assemble = f'{default_path}{COA}'

        ss = client.open_by_url(URL)
        for sheet in ss:
            if month[0:3] in sheet.title:
                month_str = sheet.title
                print(month_str)
        sheet = ss.worksheet(month_str) # sheet is now the first sheet
        rows = sheet.get_all_values()
        header = rows[0] # first row


        #Find SN, COA, and Product Key
        SN_column = []
        SN_row = 1
        COA_index = 1
        key_index = 1
        SN = product_key = COA = ''

        i = 0
        for item in header:
            i += 1
            if item == 'Serial #':
                SN_column = sheet.col_values(i)
            if item == 'New COA':
                COA_index = i
            if item == 'Product Key':
                key_index = i


        #XML Parsing Setup
        treeReport = ET.parse(report)
        xmlReport = treeReport.getroot() # xml object
        treeAssemble = ET.parse(assemble)
        xmlAssemble = treeAssemble.getroot() # xml object

        for child in xmlReport:
            if child.tag == 'SerialNumber':
                #print(child.text)
                SN = child.text

        for child in xmlAssemble:
            if child.tag == 'ProductKey':
                product_key = child.text
                #print('child.text')
            if child.tag == 'ProductKeyID':
                COA = child.text

        j = 0
        for item in SN_column:
            j+=1
            if item == SN:
                #print('working?')
                SN_row = j
                sheet.update_cell(j, COA_index,COA)
                sheet.update_cell(j, key_index, product_key)

try:
    while True:
        time.sleep(1)
except KeyboardInterrupt:
    observer.stop()

observer.join()