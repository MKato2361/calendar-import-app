import pickle
import os
import streamlit as st
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

SCOPES = ["https://www.googleapis.com/auth/calendar"]

def authenticate_google():
    # Use st.session_state to store credentials for each user's session
    if 'credentials' not in st.session_state:
        st.session_state['credentials'] = None

    creds = st.session_state['credentials']
    
    # トークンが保存されていれば読み込む (今回はsession_stateに直接保存するため、ファイルからの読み込みは初回のみ考慮)
    # 実際には、本番環境ではユーザーIDと紐付けてデータベース等に保存することが推奨されます
    # ここでは、開発環境での簡易的なマルチユーザー対応としてsession_stateを使用します

    # 認証フローの開始
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                st.session_state['credentials'] = creds # Refreshされたクレデンシャルを保存
            except Exception as e:
                st.error(f"トークンのリフレッシュに失敗しました。再認証してください: {e}")
                st.session_state['credentials'] = None # 無効なクレデンシャルをクリア
                creds = None
        else:
            try:
                client_config = {
                    "installed": {
                        "client_id": st.secrets["google"]["client_id"],
                        "client_secret": st.secrets["google"]["client_secret"],
                        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                        "token_uri": "https://oauth2.googleapis.com/token",
                        "redirect_uris": ["urn:ietf:wg:oauth:2.0:oob"]  # コンソール認証用
                    }
                }
                flow = Flow.from_client_config(client_config, SCOPES)
                flow.redirect_uri = "urn:ietf:wg:oauth:2.0:oob"
                auth_url, _ = flow.authorization_url(prompt='consent')

                st.info("以下のURLをブラウザで開いて、表示されたコードをここに貼り付けてください：")
                st.write(auth_url)
                code = st.text_input("認証コードを貼り付けてください:")

                if code:
                    flow.fetch_token(code=code)
                    creds = flow.credentials
                    st.session_state['credentials'] = creds # 新しいクレデンシャルを保存
                    st.rerun() # 認証完了後、画面を更新して再認証ループを抜ける
            except Exception as e:
                st.error(f"Google認証に失敗しました: {e}")
                st.session_state['credentials'] = None
                return None
    
    return creds

def add_event_to_calendar(service, calendar_id, event_data):
    event = service.events().insert(calendarId=calendar_id, body=event_data).execute()
    return event.get("htmlLink")
