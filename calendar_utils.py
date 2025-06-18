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
                flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
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
