import logging
logger = logging.getLogger(__name__)

import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import google.auth

class Sheets:
    
    def __init__(self, crendentials, bot_credentials, token, scopes):
        self.crendentials = crendentials
        self.bot_credentials = bot_credentials
        self.token = token
        self.scopes = scopes
        self.authenticate()
        
    def authenticate(self):
        creds = None
        if os.path.exists("src/keys/token.json"):
            creds = Credentials.from_authorized_user_file(self.token, self.scopes)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(self.crendentials, self.scopes)
                creds = flow.run_local_server(port=0)
        with open("src/keys/token.json","w") as token:
            token.write(creds.to_json())
        logger.info("Autenticação com Sheet bem-sucedida!")
        
    def get_values(self, spreadsheet_id, range_name):
        creds, _ = google.auth.load_credentials_from_file(self.bot_credentials, self.scopes)
        
        try:
            service = build("sheets", "v4", credentials=creds)
            result = ( service.spreadsheets()
                    .values()
                    .get(spreadsheetId=spreadsheet_id, range=range_name)
                    .execute() )
            rows = result.get("values", [])
            logger.info(print(f"{len(rows)} rows retrieved"))
            return result
        except HttpError as error:
            print(f"An error occurred: {error}")
            return error

    def batch_get_values(self, spreadsheet_id, _range_names):
        creds, _ = google.auth.load_credentials_from_file(self.bot_credentials, self.scopes)
        
        try:
            service = build("sheets", "v4", credentials=creds)
            range_names = [
                # Range names ...
            ]
            # [START_EXCLUDE silent]
            range_names = _range_names
            # [END_EXCLUDE]
            result = (
                service.spreadsheets()
                .values()
                .batchGet(spreadsheetId=spreadsheet_id, ranges=range_names)
                .execute()
            )
            ranges = result.get("valueRanges", [])
            print(f"{len(ranges)} ranges retrieved")
            return result
        except HttpError as error:
            print(f"An error occurred: {error}")
            return error