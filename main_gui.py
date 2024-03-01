from tkinter import filedialog as fd
from datetime import datetime
from config import config
import sys
import tkinter
import json
import requests
import csv
import sqlite3
import os


db = sqlite3.connect("MDE-Offboarder.sqlite")  # creates a new database if it doesn't exist. Otherwise, it'll open
cur = db.cursor()
cur.execute("CREATE TABLE IF NOT EXISTS offboarded (_id INTEGER PRIMARY KEY, host TEXT, machine_id TEXT, "
            "ofb_time TEXT NOT NULL, ofb_by TEXT NOT NULL)")


class Header:

    ACCEPT = '"accept": "application/json"'

    def __init__(self, tenant_id, app_id, app_secret, url):
        self._tenant_id = tenant_id
        self._app_id = app_id
        self._app_secret = app_secret
        self._url = url

    @property
    def tenant_id(self):
        return self._tenant_id

    @property
    def app_id(self):
        return self._app_id

    @property
    def app_secret(self):
        return self._app_secret

    @property
    def url(self):
        return self._url

    @url.setter
    def url(self, new_url):
        if isinstance(new_url, str):
            self._url = new_url


def azure_token():
    tenant = config.APIkeys.tenant_id
    token_header = Header(config.APIkeys.tenant_id, config.APIkeys.app_id, config.APIkeys.app_secret,
                          url=f"https://login.microsoftonline.com/{tenant}/oauth2/token")
    token_header.url = f"https://login.microsoftonline.com/{token_header.tenant_id}/oauth2/token"
    resource_app_id_uri = 'https://api-us.securitycenter.microsoft.com'

    body = {
        'resource': resource_app_id_uri,
        'client_id': token_header.app_id,
        'client_secret': token_header.app_secret,
        'grant_type': 'client_credentials'
    }

    req = requests.post(token_header.url, body)
    response = req.text
    json_response = json.loads(response)
    return json_response['access_token']


def offboard():
    machines_to_offboard = []
    with open("devices.csv", "r") as hosts:
        reader = csv.reader(hosts, delimiter=',')
        next(reader, None)  # skips the headers
        for row in reader:
            machines_to_offboard.append(row[1])

    strfilter = list(filter(None, machines_to_offboard))
    print("Attempting to offboard the following machines:")
    for k in strfilter:
        print(f"\t- {k}")
    update_cursor = db.cursor()

    run_by = os.getlogin()
    bearer_token = azure_token()
    current_dt = datetime.now()

    headers = {
        'Content-Type': 'application/json',
        'Authorization': "Bearer " + bearer_token
    }

    body = {
        "Comment": f"Offboarded machine by automation at {current_dt}."
    }

    with open('devices.csv', 'r') as dfile:
        csv_reader = csv.reader(dfile)
        for index, csv_row in enumerate(csv_reader):
            hostname = csv_row[1]
            device_id = csv_row[0]
            if index != 0:
                url = f'https://api-us.securitycenter.windows.com/api/machines/{csv_row[0]}/offboard'
                req = requests.post(url, headers=headers, json=body)
                response = req.text
                json_response = json.loads(response)
                if 'error' in json_response:
                    print(f"[+] ERROR '{json_response['error']['message']}' detected on {hostname}")
                else:
                    print(f"Request posted for {hostname}\nRequest ID: {json_response['id']}\n"
                          f"Status: {json_response['status']}")
                    update_sql = "INSERT INTO offboarded(host, machine_id, ofb_time, ofb_by) VALUES(?, ?, ?, ?)"
                    update_cursor.execute(update_sql, (hostname, device_id, current_dt, run_by))
                    db.commit()
                    cur.close()


def openfile():
    opened_file = fd.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    return opened_file


def redirector(input_string):
    text_widget2.insert(tkinter.INSERT, input_string)


main_window = tkinter.Tk()

main_window.title("MDE Offboarder")   # title of the window
main_window.geometry('800x480-8-200')    # size of the window plus the offsets

input_text_label = tkinter.Label(main_window, text="CSV File Input")# label for the window, first takes the window then the text as its arguments
input_text_label.grid(row=0, column=1)
leftFrame = tkinter.Frame(main_window)
leftFrame.grid(row=1, column=1)

canvas = tkinter.Canvas(leftFrame, borderwidth=1)
canvas.grid(row=2, column=0)

text_widget1 = tkinter.Text(canvas, height=10)
text_widget1.grid(row=0, column=0, pady=(1, 20))
output_text_label = tkinter.Label(main_window, text="Offboarding Status")
output_text_label.grid(row=1, column=1, pady=(1, 5))
text_widget2 = tkinter.Text(canvas, height=10)
text_widget2.grid(row=2, column=0)
with open('devices.csv') as twt:
    file_data = twt.readlines()
for i in file_data:
    text_widget1.insert(tkinter.END, i)


rightFrame = tkinter.Frame(main_window)
rightFrame.grid(row=1, column=2, sticky='n')

# creating the buttons
file_selector_button = tkinter.Button(rightFrame, text="Open", width=10, command=openfile)
offboard_button = tkinter.Button(rightFrame, text="Offboard", width=10, command=offboard)

# placing the buttons on the grid
offboard_button.grid(row=2, column=0, pady=(1, 20))
file_selector_button.grid(row=0, column=0, pady=(1, 20))

# configure columns
main_window.columnconfigure(0, weight=1)
main_window.columnconfigure(1, weight=1)
main_window.columnconfigure(2, weight=1)

# initializing and starting the main window loop


if __name__ == "__main__":
    # TODO: Redirect text output from STDOUT to GUI
    # sys.stdout.write = redirector
    main_window.mainloop()
