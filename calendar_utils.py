import pickle
import os
import streamlit as st
import hashlib
CREDENTIALS_FILE = "credentials.json"
SCOPES = ['https://www.googleapis.com/auth/calendar']
from google_auth_oauthlib.flow import InstalledAppFlow
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from datetime import datetime, timedelta, timezone 

SCOPES = ["https://www.googleapis.com/auth/calendar"]

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
    JST_OFFSET = timedelta(hours=9)

    start_dt_jst = start_date.replace(hour=0, minute=0, second=0, microsecond=0)
    end_dt_jst = end_date.replace(hour=23, minute=59, second=59, microsecond=999999)

    time_min_utc = (start_dt_jst - JST_OFFSET).isoformat(timespec='microseconds') + 'Z'
    time_max_utc = (end_dt_jst - JST_OFFSET).isoformat(timespec='microseconds') + 'Z'

    # st.write(f"検索期間 (UTC): {time_min_utc} から {time_max_utc}") # デバッグ用

    deleted_count = 0
    all_events_to_delete = []
    page_token = None

    # Step 1: 削除対象イベントをすべてリストアップ
    with st.spinner(f"{start_date.strftime('%Y/%m/%d')}から{end_date.strftime('%Y/%m/%d')}までの削除対象イベントを検索中..."):
        while True:
            try:
                events_result = service.events().list(
                    calendarId=calendar_id,
                    timeMin=time_min_utc,
                    timeMax=time_max_utc,
                    singleEvents=True,
                    orderBy='startTime',
                    pageToken=page_token
                ).execute()
                events = events_result.get('items', [])
                all_events_to_delete.extend(events) # 取得したイベントをリストに追加

                page_token = events_result.get('nextPageToken')
                if not page_token:
                    break 
            except Exception as e:
                st.error(f"イベントの検索中にエラーが発生しました: {e}")
                return 0 # エラーが発生したら処理を中断

    total_events = len(all_events_to_delete)
    
    # 削除対象イベントがない場合、ここでリターン
    if total_events == 0:
        return 0

    # Step 2: 取得したイベントを削除（プログレスバー表示）
    progress_bar = st.progress(0)
    
    for i, event in enumerate(all_events_to_delete):
        event_summary = event.get('summary', '不明なイベント')
        try:
            service.events().delete(calendarId=calendar_id, eventId=event['id']).execute()
            deleted_count += 1
        except Exception as e:
            st.warning(f"イベント '{event_summary}' の削除に失敗しました: {e}")
        
        # プログレスバーを更新
        progress_bar.progress((i + 1) / total_events)
    
    return deleted_count

from google_auth_oauthlib.flow import InstalledAppFlow

def get_user_token_file():
    user_info = st.session_state.get("user_id", st.session_state.get("run_id", str(datetime.now())))
    user_hash = hashlib.sha256(user_info.encode()).hexdigest()
    return f"token_{user_hash}.pickle"

def authenticate_google():
    creds = None
    token_file = get_user_token_file()

    if 'credentials' in st.session_state and st.session_state['credentials'] and st.session_state['credentials'].valid:
        creds = st.session_state['credentials']
        return creds

    if os.path.exists(token_file):
        try:
            with open(token_file, "rb") as token:
                creds = pickle.load(token)
            st.session_state['credentials'] = creds
        except Exception as e:
            st.warning("トークンの読み込みに失敗しました。再認証が必要です。")

    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file(
            CREDENTIALS_FILE, SCOPES)
        creds = flow.run_local_server(port=0)

        with open(token_file, "wb") as token:
            pickle.dump(creds, token)

        st.session_state['credentials'] = creds

    return creds
