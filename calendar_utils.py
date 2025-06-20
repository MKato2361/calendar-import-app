import pickle
import os
import streamlit as st
import getpass
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from config import SCOPES

def authenticate_google():
    creds = None

    # ユーザーごとにトークンを分けて保存
    user_id = getpass.getuser()
    token_dir = "tokens"
    os.makedirs(token_dir, exist_ok=True)
    token_path = os.path.join(token_dir, f"token_{user_id}.pickle")

    # 既存トークンを読み込み
    if os.path.exists(token_path):
        with open(token_path, "rb") as token:
            creds = pickle.load(token)

    # トークンが無効または存在しない場合は認証フロー
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            try:
                # secrets.tomlからクライアントID等を取得
                client_config = {
                    "installed": {
                        "client_id": st.secrets["google"]["client_id"],
                        "client_secret": st.secrets["google"]["client_secret"],
                        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                        "token_uri": "https://oauth2.googleapis.com/token",
                        "redirect_uris": ["http://localhost"]
                    }
                }

                # ローカルサーバでブラウザを開き認証（最も安定）
                flow = InstalledAppFlow.from_client_config(client_config, SCOPES)
                creds = flow.run_local_server(port=0, open_browser=True)

                # トークン保存
                with open(token_path, "wb") as token:
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

