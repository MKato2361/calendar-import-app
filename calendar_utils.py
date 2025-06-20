import pickle
import os
import streamlit as st
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from datetime import datetime
import json # Firebase configをパースするために必要

# Firebase関連のインポート
# Canvas環境ではfirebase_adminのSDKが利用可能
from firebase_admin import credentials, firestore, auth, initialize_app

SCOPES = ["https://www.googleapis.com/auth/calendar"]

# Firebaseを初期化（一度だけ実行）
# Canvas環境で提供される__firebase_config, __app_id, __initial_auth_tokenを使用
def initialize_firebase():
    # StreamlitのセッションステートにFirebase初期化済みフラグがあるか確認
    if 'firebase_initialized' not in st.session_state or not st.session_state.firebase_initialized:
        try:
            # st.secretsからFirebase設定を読み込む
            # __firebase_configはJSON文字列またはDictとして提供されることを想定
            firebase_config = {}
            if "__firebase_config" in st.secrets:
                # secrets.tomlでディクショナリとして定義されている場合
                firebase_config = st.secrets["__firebase_config"]
            elif "FIREBASE_CONFIG" in os.environ:
                # 環境変数としてJSON文字列で定義されている場合
                firebase_config = json.loads(os.environ["FIREBASE_CONFIG"])
            else:
                st.error("Firebaseの設定が見つかりません。Streamlit Secretsまたは環境変数に'__firebase_config'を設定してください。")
                st.session_state['firebase_initialized'] = False
                return None, None, False

            # credentials.Certificateはdictを受け取るため、直接渡す
            cred = credentials.Certificate(firebase_config)
            initialize_app(cred)
            st.session_state['firebase_initialized'] = True
            st.session_state['db'] = firestore.client()
            st.session_state['auth'] = auth
            st.info("Firebaseを初期化しました。")
        except Exception as e:
            st.error(f"Firebaseの初期化に失敗しました: {e}")
            st.session_state['firebase_initialized'] = False # 初期化失敗フラグ
            return None, None, False
    
    return st.session_state.get('db'), st.session_state.get('auth'), st.session_state.get('firebase_initialized')


def authenticate_google_with_firestore(db, firebase_auth):
    """
    Firebase Authenticationを通じてユーザーを認証し、
    FirestoreからGoogle Calendar APIの認証情報を読み書きします。
    """
    creds = None
    user_id = None # 認証されたユーザーのID

    # Firebase認証の実行または既存のユーザーIDの取得
    if 'firebase_user_id' not in st.session_state:
        try:
            # Canvas環境から提供される初期認証トークンを使用
            if "__initial_auth_token" in st.secrets:
                token = st.secrets["__initial_auth_token"]
                user = firebase_auth.sign_in_with_custom_token(token)
                user_id = user.uid
                st.session_state['firebase_user_id'] = user_id
                st.info(f"Firebaseでカスタム認証済み: UID {user_id}")
            else:
                # 初期トークンがない場合、匿名認証を試みる
                user = firebase_auth.sign_in_anonymously()
                user_id = user.uid
                st.session_state['firebase_user_id'] = user_id
                st.info(f"Firebaseで匿名認証済み: UID {user_id}")
        except Exception as e:
            st.error(f"Firebase認証に失敗しました: {e}")
            return None, None
    else:
        user_id = st.session_state['firebase_user_id']
        st.write(f"現在のFirebaseユーザーID: `{user_id}`") # ユーザーIDをUIに表示
        st.info(f"FirebaseセッションにUID {user_id} があります。")


    # Google認証情報のFirestoreパス
    # __app_idはCanvas環境で提供されるアプリケーションID
    # もし__app_idが利用できない場合はデフォルト値を使用
    app_id = st.secrets.get("__app_id", "default-app-id")
    # Firestoreのパスは、/artifacts/{appId}/users/{userId}/google_creds/calendar_api_token
    creds_doc_ref = db.collection('artifacts').document(app_id).collection('users').document(user_id).collection('google_creds').document('calendar_api_token')

    # 1. まず現在のStreamlitセッションの認証情報（メモリ上）があるか確認します
    if 'google_credentials' in st.session_state and st.session_state['google_credentials'] and st.session_state['google_credentials'].valid:
        creds = st.session_state['google_credentials']
        return creds, user_id

    # 2. session_stateにない場合、Firestoreから永続化された認証情報を読み込もうとします
    try:
        doc = creds_doc_ref.get()
        if doc.exists:
            pickled_creds = doc.to_dict().get('token')
            if pickled_creds:
                creds = pickle.loads(pickled_creds)
                # 読み込んだ認証情報が有効な場合は、session_stateに保存し再利用
                if creds.valid:
                    st.session_state['google_credentials'] = creds 
                    st.info("FirestoreからGoogle認証情報を読み込みました。")
                else:
                    st.warning("Firestoreの認証情報が有効ではありません。リフレッシュを試みます。")
                    creds = None # 無効な場合はリフレッシュまたは再認証へ
        else:
            st.info("FirestoreにGoogle認証情報が見つかりませんでした。")
    except Exception as e:
        st.warning(f"Firestoreからの認証情報読み込みに失敗しました: {e}。再認証してください。")
        creds = None

    # 3. 認証情報が有効でない、または期限切れの場合、更新または再認証を行います
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            # トークンが期限切れでリフレッシュトークンがある場合、トークンをリフレッシュします
            try:
                creds.refresh(Request())
                # リフレッシュされた認証情報をFirestoreとsession_stateに保存します
                # IMPORTANT: In a production environment, sensitive information like refresh tokens should be encrypted.
                creds_doc_ref.set({'token': pickle.dumps(creds)})
                st.session_state['google_credentials'] = creds
                st.success("Google認証トークンを更新し、Firestoreに保存しました。")
                st.rerun() # トークン更新後、アプリを再実行して変更を反映
            except Exception as e:
                st.error(f"トークンのリフレッシュに失敗しました。再認証してください: {e}")
                st.session_state['google_credentials'] = None
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
                    # 新しい認証情報をFirestoreとsession_stateに保存します
                    # IMPORTANT: In a production environment, sensitive information like refresh tokens should be encrypted.
                    creds_doc_ref.set({'token': pickle.dumps(creds)})
                    st.session_state['google_credentials'] = creds
                    st.success("Google認証が完了しました！Firestoreに保存しました。")
                    st.rerun() # 認証成功後、アプリを再読み込み
            except Exception as e:
                st.error(f"Google認証に失敗しました: {e}")
                st.session_state['google_credentials'] = None
                return None, None
    
    return creds, user_id

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

