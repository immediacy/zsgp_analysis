"""
Shows basic usage of the Apps Script API.
Call the Apps Script API to create a new script project, upload a file to the
project, and log the script's URL to the user.
"""
import os.path, traceback

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient import errors
import gspread
import pandas as pd

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


def connect_ss():
    """Calls the Apps Script API."""
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    try:
        if os.path.exists("token.json"):
            creds = Credentials.from_authorized_user_file("token.json", SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    "credentials.json", SCOPES
                )
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open("token.json", "w") as token:
                token.write(creds.to_json())
        client = gspread.authorize(creds)
        return client
    except:
        traceback.print_exc()
        return 0


def read_values(connection, spreadssheet_id, sheetname, df_flag=True):

    try:
        # sheet = service.spreadsheets()
        sheet = connection.open_by_key(spreadssheet_id).worksheet(sheetname)

        if df_flag:
            values = pd.DataFrame(sheet.get_all_records())
        else:
            values = sheet.get_all_values()

        return values

    except errors.HttpError as error:
        # The API encountered a problem.
        print(error.content)
