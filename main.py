from datetime import datetime
import json
import requests

# TODO: Allow users to upload .csv file of hostnames OR IP addresses to offboard
machines_to_offboard = ['<HOSTNAME HERE>']


class Header:

    ACCEPT = '"accept": "application/json"'

    def __init__(self, url, tenant_id, app_id, app_secret):
        self._url = url
        self._tenant_id = tenant_id
        self._app_id = app_id
        self._app_secret = app_secret

    @property
    def tenant_id(self):
        return self.tenant_id

    @property
    def app_id(self):
        return self.app_id

    @property
    def app_secret(self):
        return self.app_secret

    @property
    def url(self):
        return self._url

    @url.setter
    def url(self, new_endpoint):
        if isinstance(new_endpoint, str) and "https://api-us.securitycenter.microsoft.com" in new_endpoint:
            self._url += new_endpoint

    def __str__(self):
        return f"{self.ACCEPT}, 'Authorization': 'Bearer {self._access_key};secretKey={self._secret_key}'"

    def asdict(self):
        return {'accept': "application/json", 'X-ApiKeys': f'accessKey={self._access_key};secretKey={self._secret_key}'}


def aad_token():
    tenant_id = ""
    app_id = ""
    app_secret = ""
    url = f""
    resourceAppIdUri = 'https://api-us.securitycenter.microsoft.com'
    body = {
        'resource': resourceAppIdUri,
        'client_id': app_id,
        'client_secret': app_secret,
        'grant_type': 'client_credentials'
    }

    req = requests.post(url, body)
    response = req.text
    json_response = json.loads(response)
    return json_response['access_token']


def get_machine_id(aad_token):
    ids_to_offboard = []

    url = 'https://api-us.securitycenter.microsoft.com/api/machines'
    headers = {
        'Content-Type': 'application/json',
        'Authorization': "Bearer " + aad_token
    }
    req = requests.get(url, headers=headers)
    response = req.text
    json_response = json.loads(response)

    for i in json_response['value']:
        for k in machines_to_offboard:
            if k in str(i['computerDnsName']):
                ids_to_offboard.append(i['id'])
    return ids_to_offboard


def offboard():
    bearer_token = aad_token()
    machine_id_list = get_machine_id(bearer_token)
    eastern_dt = datetime.now()
    headers = {
        'Content-Type': 'application/json',
        'Authorization': "Bearer " + bearer_token
    }
    body = {
        "Comment": f"Offboarded machine by automation at {eastern_dt}."
    }
    for j in machine_id_list:
        url = f'https://api-us.securitycenter.windows.com/api/machines/{j}/offboard'
        req = requests.post(url, headers=headers, json=body)
        response = req.text
        json_response = json.loads(response)
        print(json_response)


if __name__ == "__main__":
    new_token = aad_token()
    print(get_machine_id(new_token))
    offboard()
