# calendar_utils.py

import os
import pickle
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import streamlit as st
from config import SCOPES, TOKEN_PATH, CREDENTIALS_FILE

def authenticate_google():
    creds = None
    if os.path.exists(TOKEN_PATH):
        with open(TOKEN_PATH, 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            try:
                from google_auth_oauthlib.flow import Flow
import streamlit as st

client_config = {
    "installed": {
        "client_id": st.secrets["google"]["client_id"],
        "client_secret": st.secrets["google"]["client_secret"],
        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
        "token_uri": "https://oauth2.googleapis.com/token",
        "redirect_uris": ["http://localhost"]
    }
}

flow = Flow.from_client_config(client_config, SCOPES)
creds = flow.run_local_server(port=0)

                creds = flow.run_local_server(port=0)
            except FileNotFoundError:
                st.error("credentials.json が見つかりません。")
                return None
            except Exception as e:
                st.error(f"Google認証中にエラーが発生しました: {e}")
                return None

        with open(TOKEN_PATH, 'wb') as token:
            pickle.dump(creds, token)
    return creds

def add_event_to_calendar(service, calendar_id, event_data):
    event = service.events().insert(calendarId=calendar_id, body=event_data).execute()
    return event.get('htmlLink')
