import json
import requests
import csv
import sqlite3
from datetime import datetime
from config import config

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

    initials = input("Enter your initials to claim responsibility for this action: ")
    bearer_token = azure_token()
    current_dt = datetime.now()

    headers = {
        'Content-Type': 'application/json',
        'Authorization': "Bearer " + bearer_token
    }

    body = {
        "Comment": f"Offboarded machine by automation at {current_dt}. {initials}."
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
                    update_cursor.execute(update_sql, (hostname, device_id, current_dt, initials))
                    db.commit()
                    cur.close()


if __name__ == "__main__":
    offboard()
