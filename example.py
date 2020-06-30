import os
import pickle
from typing import List

import valispace

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

from datetime import datetime

# Setup Google Sheets API and put the path to credentials.json
# token.pickle will appear after the first run of the script in the same folder
credentials_path = "<path>/credentials.json"
token_path = "<path>/token.pickle"

# Put here your Valispace data
valispace_url = "https://<organisation_name>.valispace.com/"
valispace_project_name = "<project_name>"
valispace_username = "<valispace_username>"
valispace_password = "<valispace_password>"

# Configure ID of the online spreadsheet, for example, "1zWBZzAl4g-xyaNn9XSg0S1yYjET0ZnhPgXcKbuIFhyM"
# You can take it from the link to the spreadsheet
SPREADSHEET_ID = "<SPREADSHEET_ID>"

# If modifying these scopes, delete the file token.pickle
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']


def get_valispace_comps(url: str, project_name: str,
                        username: str, password: str) -> list:
    valispace_api = valispace.API(url=url,
                                  username=username, password=password)
    project_comps = valispace_api.get_component_list(project_name=project_name)
    comps_list = []
    for comp in project_comps:
        comps_list.append([comp["id"], comp["name"], comp["parent"]])

    return comps_list


def init_google_sheets_api():
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time

    if os.path.exists(token_path):
        with open(token_path, 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(token_path, 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)
    return service


def write_to_google_sheets(info: List[list], range: str,
                           spreadsheet_id: str, service: build):
    body = {'values': info}
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id, range=range,
        valueInputOption='RAW', body=body).execute()


def write_last_update_time(spreadsheet_id: str, service: build):
    values = [
        [
            "Last update ", str(datetime.today().isoformat())
        ],
    ]
    body = {"values": values}

    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id, range="H1:I1",
        valueInputOption="RAW", body=body).execute()


def main():
    service = init_google_sheets_api()
    get_valispace_comps(valispace_url, valispace_project_name,
                        valispace_username, valispace_password)

    # Write title to the spreadsheet
    titles = [["ID", "Name", "Parent ID"]]
    titles_range = "A2:C2"
    write_to_google_sheets(info=titles, range=titles_range,
                           spreadsheet_id=SPREADSHEET_ID, service=service)

    # Write time to the spreadsheet
    write_last_update_time(spreadsheet_id=SPREADSHEET_ID, service=service)

    # Write components data to the spreadsheet
    comp_range = "A3:C"
    comp_data = get_valispace_comps(url=valispace_url,
                                    project_name=valispace_project_name,
                                    username=valispace_username,
                                    password=valispace_password)
    write_to_google_sheets(info=comp_data, range=comp_range,
                           spreadsheet_id=SPREADSHEET_ID, service=service)


main()
