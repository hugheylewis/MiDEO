from tkinter import filedialog as fd
from datetime import datetime
from config import config
import tkinter
import json
import requests
import csv
import sqlite3
import os
import sys


db = sqlite3.connect("MDE-Offboarder.sqlite")  # creates a new database if it doesn't exist. Otherwise, it'll open
cur = db.cursor()
cur.execute("CREATE TABLE IF NOT EXISTS offboarded (_id INTEGER PRIMARY KEY, host TEXT, machine_id TEXT, "
            "ofb_time TEXT NOT NULL, ofb_by TEXT NOT NULL)")


class StdoutRedirector:
    """Redirects output from STDOUT to tkinter widget"""
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, message):
        self.text_widget.insert(tkinter.END, message)
        self.text_widget.see(tkinter.END)


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
    with open(file_path.name, "r") as hosts:
        reader = csv.reader(hosts, delimiter=',')
        next(reader, None)  # skips the headers
        for row in reader:
            machines_to_offboard.append(row[1])

    strfilter = list(filter(None, machines_to_offboard))
    output_text_widget.delete(1.0, tkinter.END)
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

    with open(file_path.name, 'r') as dfile:
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
                    output_text_widget.insert(tkinter.END, f"[-] ERROR '{json_response['error']['message']}' detected "
                                                           f"on {hostname}\n")
                else:
                    print(f"Request posted for {hostname}\nRequest ID: {json_response['id']}\n"
                          f"Status: {json_response['status']}")
                    update_sql = "INSERT INTO offboarded(host, machine_id, ofb_time, ofb_by) VALUES(?, ?, ?, ?)"
                    update_cursor.execute(update_sql, (hostname, device_id, current_dt, run_by))
                    db.commit()
                    cur.close()


def open_file():
    # TODO: Replace global variable with a returned function value. This is just to prove that it works
    global file_path
    file_path = fd.askopenfile(filetypes=[("CSV files", "*.csv")])
    is_valid, message = validate_csv_format(file_path)
    if file_path and is_valid:
        input_text_widget.delete(1.0, tkinter.END)
        input_text_widget.insert(tkinter.END, "The following hostnames will be offboarded from MDE:\n")
        output_text_widget.delete(1.0, tkinter.END)
        output_text_widget.insert(tkinter.END, "Ready to offboard")
        with open(file_path.name) as twt:
            file_data = csv.reader(twt, delimiter=',')
            next(file_data, None)
            counter = 1
            for i in file_data:
                input_text_widget.insert(tkinter.END, f"\t[{counter}] {i[1]}\n")
                counter += 1
        opened_file_text.set(f"Opened file: {file_path.name}")
        num_devices_var.set(f"Number of devices: {counter - 1}")
    else:
        output_text_widget.delete(1.0, tkinter.END)
        input_text_widget.delete(1.0, tkinter.END)
        input_text_widget.insert(tkinter.END, "See error below")
        print("CSV file is not valid:", message)


def validate_csv_format(csv_file_path):
    # Define expected column names and data types
    expected_columns = ["ï»¿Device ID", "Device Name", "Domain", "First Seen", "Last device update", "OS Platform",
                        "OS Distribution", "OS Version", "OS Build", "Windows 10 Version", "Tags", "Group",
                        "Is AAD Joined", "Device IPs", "Risk Level", "Exposure Level", "Health Status",
                        "Onboarding Status", "Device Role", "Cloud Platforms", "Managed By", "Antivirus status",
                        "Is Internet Facing"]
    expected_data_types = [str, str, str]

    # Open the CSV file
    with open(csv_file_path.name, 'r') as file:
        csv_reader = csv.reader(file)
        headers = next(csv_reader)  # Read the header row
        if headers != expected_columns:
            return False, "Column names do not match the expected format."

        # Validate each row
        line_number = 2  # Start from the second line (after the header)
        for row in csv_reader:
            if len(row) != len(expected_columns):
                return False, f"Row {line_number}: Incorrect number of columns."

            for i, (value, expected_type) in enumerate(zip(row, expected_data_types)):
                try:
                    expected_type(value)  # Attempt to convert value to expected data type
                except ValueError:
                    return False, f"Row {line_number}, Column {expected_columns[i]}: Invalid data type."

            line_number += 1

    return True, "CSV file conforms to the specified format."


def message_window():
    win = tkinter.Toplevel()
    win.title('Help Menu')
    with open('help_menu_text.txt', 'r') as hmt:
        message = hmt.read()
        help_menu_label = tkinter.Label(win, text=message, justify='left')
        help_menu_label.grid(row=0, column=0, sticky='w', padx=(5, 0), pady=(5, 0))
        help_menu_label.grid_rowconfigure(0, weight=1)
        tkinter.Button(win, text='OK', width=10, command=win.destroy).grid(pady=(15, 15))


def offboard_single_device_window():
    def scroll_text_view(*args):
        single_device_offboard_status.yview(*args)

    def get_single_device():
        bearer_token = azure_token()
        device_hostname = entry.get()

        url = f"https://api-us.securitycenter.microsoft.com/api/machines?$filter=computerDnsName+eq+'{device_hostname}'"
        headers = {
            'Content-Type': 'application/json',
            'Authorization': "Bearer " + bearer_token
        }
        req = requests.get(url, headers=headers)
        response = req.text
        json_response = json.loads(response)

        for i in json_response['value']:
            device_id = i['id']
            last_ip = i['lastIpAddress']
            os_platform = i['osPlatform']
            all_known_ips = i['ipAddresses']
            entra_joined_status = i['isAadJoined']
            single_device_offboard_status.delete(1.0, tkinter.END)
            single_device_offboard_status.insert(tkinter.END, f"Hostname: {device_hostname}\nDevice ID: "
                                                              f"{device_id}\nEntra Joined?: {entra_joined_status}\n"
                                                              f"Last Known IP: {last_ip}\nOS: {os_platform}"
                                                              f"\nAll Known IPs: \n")
            for ip in all_known_ips:
                single_device_offboard_status.insert(tkinter.END, f"\tIP: {ip['ipAddress']}\n\tMAC: "
                                                                  f"{ip['macAddress']}\n\n")
            return device_hostname, device_id

    def offboard_single_device():
        device_hostname, device_id = get_single_device()
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

        url = f'https://api-us.securitycenter.windows.com/api/machines/{device_id}/offboard'
        req = requests.post(url, headers=headers, json=body)
        response = req.text
        json_response = json.loads(response)
        if 'error' in json_response:
            single_device_offboard_status.delete(1.0, tkinter.END)
            single_device_offboard_status.insert(tkinter.END, f"[-] ERROR '{json_response['error']['message']}"
                                                              f"' detected on {device_hostname}\n\n")
        else:
            single_device_offboard_status.delete(1.0, tkinter.END)
            single_device_offboard_status.insert(tkinter.END, f"Request posted for {device_hostname}\nRequest ID: "
                                                              f"{json_response['id']}\nStatus: "
                                                              f"{json_response['status']}\n\n")
            update_sql = "INSERT INTO offboarded(host, machine_id, ofb_time, ofb_by) VALUES(?, ?, ?, ?)"
            update_cursor.execute(update_sql, (device_hostname, device_id, current_dt, run_by))
            db.commit()
            cur.close()
        pass

    win = tkinter.Toplevel()
    win.geometry('535x300')
    win.title('Offboard Single Device')

    top_frame = tkinter.Frame(win)
    top_frame.grid(row=0, column=0, sticky='w')
    tkinter.Label(top_frame, text="Offboard a single device from MDE").grid(row=0, column=0, sticky='w', padx=(5, 0), pady=(5, 0))

    second_top_frame = tkinter.Frame(win)
    second_top_frame.grid(row=1, column=0, sticky='w')
    entry = tkinter.Entry(second_top_frame)
    entry.grid(row=1, column=0, padx=(8, 0), pady=(10, 0), sticky='w')
    tkinter.Button(second_top_frame, text='Get Info', width=10, command=get_single_device).grid(row=1, column=1, sticky='w', padx=(15, 0))
    tkinter.Button(second_top_frame, text='Offboard', width=10, command=offboard_single_device).grid(row=1, column=2, sticky='w', padx=(10, 0))

    middle_frame = tkinter.Frame(win)
    middle_frame.grid(row=2, column=0, sticky='w', padx=(5, 0), pady=(20, 0))
    single_device_offboard_status = tkinter.Text(middle_frame, height=8, width=60)
    single_device_offboard_status.grid(row=0, column=0, columnspan=2)
    scrollbar = tkinter.Scrollbar(middle_frame, orient="vertical", command=scroll_text_view)
    scrollbar.grid(row=0, column=3, sticky='ns')
    single_device_offboard_status.config(yscrollcommand=scrollbar.set)
    done_button_frame = tkinter.Frame(win)
    done_button_frame.grid(row=3, column=0)
    tkinter.Button(done_button_frame, text='Done', width=10, command=win.destroy).grid(row=0, column=1, pady=(15, 0))


main_window = tkinter.Tk()

main_window.title("MDE Offboarder")   # title of the window
main_window.geometry('800x480-400-200')    # size of the window plus the offsets

input_text_label = tkinter.Label(main_window, text="CSV File Input")
input_text_label.grid(row=0, column=1)
leftFrame = tkinter.Frame(main_window)
leftFrame.grid(row=1, column=1)

canvas = tkinter.Canvas(leftFrame, borderwidth=1)
canvas.grid(row=2, column=0)

# text boxes
input_text_widget = tkinter.Text(canvas, height=10)
input_text_widget.grid(row=0, column=0, pady=(1, 35))
output_text_label = tkinter.Label(main_window, text="Offboarding Status")
output_text_label.grid(row=1, column=1, pady=(1, 5))
output_text_widget = tkinter.Text(canvas, height=10)
output_text_widget.grid(row=2, column=0)

# statistics canvas
bottom_frame = tkinter.Frame(main_window)
bottom_frame.grid(row=2, column=1, sticky='w')
stats_canvas = tkinter.Canvas(bottom_frame, borderwidth=1)
stats_canvas.grid(row=2, column=0)
opened_file_text = tkinter.StringVar()
num_devices_var = tkinter.StringVar()
statistics_widget = tkinter.Label(stats_canvas, textvariable=opened_file_text)
statistics_widget.grid(row=0, column=0, sticky='w', padx=(10, 0))
num_devices_widget = tkinter.Label(stats_canvas, textvariable=num_devices_var)
num_devices_widget.grid(row=1, column=0, sticky='w', padx=(10, 0))


input_text_widget.insert(tkinter.END, 'Click "Open" and select a valid CSV file')


rightFrame = tkinter.Frame(main_window)
rightFrame.grid(row=1, column=2, sticky='n')

# creating the buttons
help_button = tkinter.Button(rightFrame, text='Help', width=10, command=message_window)
file_selector_button = tkinter.Button(rightFrame, text="Open", width=10, command=open_file)
offboard_button = tkinter.Button(rightFrame, text="Offboard", width=10, command=offboard)
single_device_button = tkinter.Button(rightFrame, text="Single Device", width=10, command=offboard_single_device_window)

# placing the buttons on the grid
help_button.grid(row=0, column=0, padx=(15, 5), pady=(0, 100))
offboard_button.grid(row=5, column=0, padx=(15, 5), pady=(1, 20))
file_selector_button.grid(row=4, column=0, padx=(15, 5), pady=(1, 20))
single_device_button.grid(row=9, column=0, padx=(15, 5), pady=(120, 0))

# configure columns
main_window.columnconfigure(0, weight=1)
main_window.columnconfigure(1, weight=1)
main_window.columnconfigure(2, weight=1)

# initializing and starting the main window loop
if __name__ == "__main__":
    sys.stdout = StdoutRedirector(output_text_widget)
    main_window.mainloop()
