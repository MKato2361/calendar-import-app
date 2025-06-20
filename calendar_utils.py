import pickle
import os
import streamlit as st
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from datetime import datetime, timezone # timezoneをインポート

SCOPES = ["https://www.googleapis.com/auth/calendar"]

def authenticate_google():
    creds = None
    
    # 1. まず現在のセッションの認証情報がst.session_stateにあるか確認します
    if 'credentials' in st.session_state and st.session_state['credentials'] and st.session_state['credentials'].valid:
        creds = st.session_state['credentials']
        return creds

    # 認証情報がない場合、または期限切れの場合に認証フローを開始します
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            # トークンが期限切れでリフレッシュトークンがある場合、トークンをリフレッシュします
            try:
                creds.refresh(Request())
                # リフレッシュされた認証情報をsession_stateに保存します
                st.session_state['credentials'] = creds
                st.info("認証トークンを更新しました。")
                # st.rerun() # トークン更新後、アプリを再実行して変更を反映（不要な場合があるためコメントアウト）
            except Exception as e:
                st.error(f"トークンのリフレッシュに失敗しました。再認証してください: {e}")
                st.session_state['credentials'] = None
                creds = None
        else: # 有効な認証情報がない、またはリフレッシュトークンがない場合、新しい認証フローを開始します
            st.warning("Googleカレンダーにアクセスするには認証が必要です。")
            try:
                client_config = {
                    "installed": {
                        "client_id": st.secrets["google"]["client_id"],
                        "client_secret": st.secrets["google"]["client_secret"],
                        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                        "token_uri": "https://oauth2.googleapis.com/token",
                        "redirect_uris": ["urn:ietf:wg:oauth:2.0:oob", "http://localhost"] # ローカルテスト用も含む
                    }
                }
                flow = Flow.from_client_config(client_config, SCOPES)
                flow.redirect_uri = "urn:ietf:wg:oauth:2.0:oob" # OOB (Out-of-Band)方式を維持

                auth_url, _ = flow.authorization_url(prompt='consent')

                st.info("以下のURLをブラウザで開いて、表示されたコードをここに貼り付けてください：")
                st.markdown(f"**[Google認証ページを開く]({auth_url})**") # リンクとして表示
                code = st.text_input("認証コードを貼り付けてください:", key="auth_code_input")

                if code:
                    try:
                        flow.fetch_token(code=code)
                        creds = flow.credentials
                        # 新しい認証情報をsession_stateに保存します
                        st.session_state['credentials'] = creds
                        st.success("Google認証が完了しました！")
                        st.rerun() # 認証成功後、アプリを再読み込み
                    except Exception as token_e:
                        st.error(f"認証コードの検証に失敗しました。コードが正しいか確認してください。: {token_e}")
                        st.session_state['credentials'] = None
                        creds = None
            except Exception as e:
                st.error(f"Google認証フローの開始に失敗しました: {e}")
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
    # タイムゾーン情報を付与 (ローカルタイムゾーンを想定し、システムデフォルトのタイムゾーンで解釈)
    # 明示的にJST (UTC+9) として扱うか、タイムゾーンを考慮しないシンプルなISOフォーマットにするか選択
    # ここでは、タイムゾーンを付与しないisoformat()にZを付与する形で、APIがUTCとして解釈するのを期待します。
    # より厳密には、datetimeオブジェクトをタイムゾーンアウェアにしてからisoformat()する必要があります。
    
    # イベント登録時にタイムゾーンを'Asia/Tokyo'に設定しているので、削除時も同様に扱います。
    # pytzなどのライブラリを使用するとより正確ですが、ここでは標準ライブラリで対応します。
    
    # start_date, end_dateはStreamlitのdate_inputから来ているため、datetimeオブジェクトに変換時に時間部分が0:0:0になっている。
    # そのため、正確な範囲をカバーするために以下のように変換する。

    # イベントのstart/endがdatetimeの場合（時間指定イベント）
    # API呼び出しのtimeMin/timeMaxはISO 8601形式でなければならない。
    # 日本時間で指定された日付範囲をUTCに変換してAPIに渡す
    
    # 開始時刻（午前0時）と終了時刻（午後11時59分59秒999999）を日本時間で設定
    start_dt_jst = start_date.replace(hour=0, minute=0, second=0, microsecond=0)
    end_dt_jst = end_date.replace(hour=23, minute=59, second=59, microsecond=999999)

    # タイムゾーン情報を持たないdatetimeオブジェクトをUTCに変換
    # 実際のタイムゾーン変換はより複雑ですが、ここでは簡易的に9時間引くことでUTCに変換
    # これは厳密ではありませんが、多くのケースで機能します。
    # 正確なタイムゾーン変換には `pytz` ライブラリの利用が推奨されます。
    # 参考: https://developers.google.com/calendar/api/v3/reference/events/list
    
    # タイムゾーンを考慮したISO形式の文字列を作成
    # イベント登録時に'Asia/Tokyo'を使っているので、検索もそのタイムゾーンを考慮すべき
    # しかし、listメソッドのtimeMin/timeMaxはRFC3339形式（UTC）を要求する
    # なので、JSTのdatetimeをUTCのdatetimeに変換する必要がある
    
    # JSTのオフセット
    JST_OFFSET = timedelta(hours=9)

    # JSTのdatetimeオブジェクトをUTCに変換
    time_min_utc = (start_dt_jst - JST_OFFSET).isoformat() + 'Z'
    time_max_utc = (end_dt_jst - JST_OFFSET).isoformat() + 'Z'

    st.write(f"検索期間 (UTC): {time_min_utc} から {time_max_utc}") # デバッグ用

    deleted_count = 0
    page_token = None

    with st.spinner(f"{start_date.strftime('%Y/%m/%d')}から{end_date.strftime('%Y/%m/%d')}までのイベントを検索中..."):
        while True:
            try:
                events_result = service.events().list(
                    calendarId=calendar_id,
                    timeMin=time_min_utc,
                    timeMax=time_max_utc,
                    singleEvents=True,
                    orderBy='startTime',
                    pageToken=page_token
                    # showDeleted=False はデフォルトなので不要
                    # status='confirmed' はデフォルトでconfirmedのみを返すため不要。
                    # showHiddenEvents=True # 隠れたイベントも対象にする場合はコメント解除
                ).execute()
                events = events_result.get('items', [])

                if not events:
                    break # イベントがなければループを終了

                for event in events:
                    # デバッグ情報
                    event_summary = event.get('summary', '不明なイベント')
                    event_start = event['start'].get('dateTime', event['start'].get('date'))
                    st.info(f"削除対象イベント: {event_summary} (開始: {event_start})")

                    try:
                        service.events().delete(calendarId=calendar_id, eventId=event['id']).execute()
                        deleted_count += 1
                        st.success(f"イベント '{event_summary}' を削除しました。")
                    except Exception as e:
                        st.warning(f"イベント '{event_summary}' の削除に失敗しました: {e}")
                
                page_token = events_result.get('nextPageToken')
                if not page_token:
                    break # 次のページがなければループを終了
            except Exception as e:
                st.error(f"イベントの検索中にエラーが発生しました: {e}")
                break # エラーが発生したらループを終了
    
    return deleted_count
