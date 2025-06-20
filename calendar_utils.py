import pickle
import os
import streamlit as st
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from config import SCOPES, TOKEN_PATH, TOKEN_DIR

def authenticate_google():
    creds = None
    os.makedirs(TOKEN_DIR, exist_ok=True)

    if os.path.exists(TOKEN_PATH):
        with open(TOKEN_PATH, "rb") as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            try:
                client_config = {
                    "installed": {
                        "client_id": st.secrets["google"]["client_id"],
                        "client_secret": st.secrets["google"]["client_secret"],
                        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                        "token_uri": "https://oauth2.googleapis.com/token",
                        "redirect_uris": ["http://localhost"]
                    }
                }
                flow = InstalledAppFlow.from_client_config(client_config, SCOPES)
                creds = flow.run_local_server(port=0)
                with open(TOKEN_PATH, "wb") as token:
                    pickle.dump(creds, token)
            except Exception as e:
                st.error(f"Google認証に失敗しました: {e}")
                return None
    return creds

def add_event_to_calendar(service, calendar_id, event_data):
    event = service.events().insert(calendarId=calendar_id, body=event_data).execute()
    return event.get("htmlLink")

def delete_events_in_range(service, calendar_id, start_date, end_date, keyword=None):
    deleted = 0
    try:
        events_result = service.events().list(
            calendarId=calendar_id,
            timeMin=start_date.isoformat() + 'Z',
            timeMax=end_date.isoformat() + 'Z',
            singleEvents=True,
            orderBy='startTime'
        ).execute()
        events = events_result.get('items', [])
        for event in events:
            if keyword and keyword not in event.get('summary', ''):
                continue
            try:
                service.events().delete(calendarId=calendar_id, eventId=event['id']).execute()
                deleted += 1
            except Exception as e:
                st.warning(f"削除失敗: {event.get('summary', 'No Title')} → {e}")
    except Exception as e:
        st.error(f"イベント取得エラー: {e}")
    return deleted
