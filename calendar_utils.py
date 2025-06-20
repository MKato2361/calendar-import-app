import pickle
import os
import streamlit as st
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from datetime import datetime

SCOPES = ["https://www.googleapis.com/auth/calendar"]
TOKEN_FILE = "token.pickle" # 認証トークンを保存するファイル名

def authenticate_google():
    creds = None
    
    # 1. まず現在のセッションの認証情報がst.session_stateにあるか確認します
    if 'credentials' in st.session_state and st.session_state['credentials'] and st.session_state['credentials'].valid:
        creds = st.session_state['credentials']
        return creds

    # 2. session_stateにない場合、token.pickleから永続化された認証情報を読み込もうとします
    if os.path.exists(TOKEN_FILE):
        try:
            with open(TOKEN_FILE, "rb") as token:
                creds = pickle.load(token)
            # 読み込んだ認証情報をsession_stateに保存し、このセッションで再利用できるようにします
            st.session_state['credentials'] = creds 
        except Exception as e:
            st.warning(f"既存のトークンファイルの読み込みに失敗しました: {e}。再認証してください。")
            creds = None

    # 3. 認証情報が有効でない、または期限切れの場合、更新または再認証を行います
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            # トークンが期限切れでリフレッシュトークンがある場合、トークンをリフレッシュします
            try:
                creds.refresh(Request())
                # リフレッシュされた認証情報をファイルとsession_stateに保存します
                with open(TOKEN_FILE, "wb") as token:
                    pickle.dump(creds, token)
                st.session_state['credentials'] = creds
                st.info("認証トークンを更新しました。")
                st.rerun() # トークン更新後、アプリを再実行して変更を反映
            except Exception as e:
                st.error(f"トークンのリフレッシュに失敗しました。再認証してください: {e}")
                st.session_state['credentials'] = None
                creds = None
        else: # 有効な認証情報がない場合、新しい認証フローを開始します
            try:
                client_config = {
                    "installed": {
                        "client_id": st.secrets["google"]["client_id"],
                        "client_secret": st.secrets["google"]["client_secret"],
                        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                        "token_uri": "https://oauth2.googleapis.com/token",
                        "redirect_uris": ["urn:ietf:wg:oauth:2.0:oob"] # コンソール認証用
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
                    # 新しい認証情報をファイルとsession_stateに保存します
                    with open(TOKEN_FILE, "wb") as token:
                        pickle.dump(creds, token)
                    st.session_state['credentials'] = creds
                    st.success("Google認証が完了しました！")
                    st.rerun() # 認証成功後、アプリを再読み込み
            except Exception as e:
                st.error(f"Google認証に失敗しました: {e}")
                st.session_state['credentials'] = None
                return None
    
    return creds

def add_event_to_calendar(service, calendar_id, event_data):
    """
    Googleカレンダーにイベントを追加します。
    """
    event = service.events().insert(calendarId=calendar_id, body=event_data).execute()
    return event.get("htmlLink")

def delete_events_from_calendar(service, calendar_id, start_date: datetime, end_date: datetime):
    """
    指定された期間内のGoogleカレンダーイベントを削除します。
    """
    # 期間の終了日を1日進めて、終了日を含むようにします
    # Google Calendar APIのtimeMaxは排他的なので、指定日の終わりまで含めるにはその日の終わりを指定
    end_date_inclusive = end_date.replace(hour=23, minute=59, second=59, microsecond=999999)
    
    # ISO 8601形式に変換し、UTC時間として扱います
    time_min = start_date.isoformat() + 'Z'
    time_max = end_date_inclusive.isoformat() + 'Z'

    deleted_count = 0
    page_token = None

    with st.spinner(f"{start_date.strftime('%Y/%m/%d')}から{end_date.strftime('%Y/%m/%d')}までのイベントを検索中..."):
        while True:
            events_result = service.events().list(
                calendarId=calendar_id,
                timeMin=time_min,
                timeMax=time_max,
                singleEvents=True,
                orderBy='startTime',
                pageToken=page_token
            ).execute()
            events = events_result.get('items', [])

            if not events:
                break # イベントがなければループを終了

            for event in events:
                try:
                    service.events().delete(calendarId=calendar_id, eventId=event['id']).execute()
                    deleted_count += 1
                except Exception as e:
                    st.warning(f"イベント '{event.get('summary', '不明なイベント')}' の削除に失敗しました: {e}")
            
            page_token = events_result.get('nextPageToken')
            if not page_token:
                break # 次のページがなければループを終了
    
    return deleted_count


