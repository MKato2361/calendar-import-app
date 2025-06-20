import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
from excel_parser import process_excel_files
# 認証関数をauthenticate_google_with_firestoreに変更
from calendar_utils import initialize_firebase, authenticate_google_with_firestore, add_event_to_calendar, delete_events_from_calendar
from googleapiclient.discovery import build
import json # Firebase configをパースするために必要

st.set_page_config(page_title="Googleカレンダー登録・削除ツール", layout="wide")
st.title("📅 Googleカレンダー一括イベント登録・削除")

# Firebaseの初期化
# アプリケーションの最初に一度だけ実行されるようにする
db, firebase_auth, firebase_initialized = initialize_firebase()

# Firebaseの初期化が成功したかを確認し、失敗していればアプリを停止
if not firebase_initialized:
    st.error("Firebaseの初期化に失敗したため、アプリケーションは動作しません。Firebaseの設定を確認してください。")
    st.stop()


# ファイルアップロードとイベント設定、イベント削除のタブを作成
tabs = st.tabs(["1. ファイルのアップロード", "2. イベントの登録", "3. イベントの削除"])

with tabs[0]:
    st.header("ファイルをアップロード")
    uploaded_files = st.file_uploader("Excelファイルを選択（複数可）", type=["xlsx"], accept_multiple_files=True)

    description_columns_pool = set()
    if uploaded_files:
        for file in uploaded_files:
            try:
                df_temp = pd.read_excel(file, engine="openpyxl")
                df_temp.columns = [str(c).strip() for c in df_temp.columns]
                description_columns_pool.update(df_temp.columns)
            except Exception as e:
                st.warning(f"{file.name} の読み込みに失敗しました: {e}")
        # セッションステートにdescription_columns_poolを保存
        st.session_state['description_columns_pool'] = list(description_columns_pool)
        # アップロードされたファイルをセッションステートに保存
        # Streamlitのファイルアップローダーはファイルオブジェクトを返すため、そのまま渡せる
        st.session_state['uploaded_files'] = uploaded_files 
    
    # セッションステートにdescription_columns_poolがない場合の初期化
    if 'description_columns_pool' not in st.session_state:
        st.session_state['description_columns_pool'] = [] 
    
    # 以前アップロードされたファイルがあればそれを表示
    if st.session_state.get('uploaded_files'):
        st.subheader("アップロード済みのファイル:")
        for f in st.session_state['uploaded_files']:
            st.write(f"- {f.name}")


with tabs[1]:
    st.header("イベントを登録")
    # アップロードファイルがセッションステートにない場合、処理を停止
    if not st.session_state.get('uploaded_files'):
        st.info("先に「1. ファイルのアップロード」タブでExcelファイルをアップロードしてください。")
        st.stop()

    # イベント設定
    st.subheader("📝 イベント設定")
    all_day_event = st.checkbox("終日イベントとして登録", value=False)
    private_event = st.checkbox("非公開イベントとして登録", value=True)
    
    # セッションステートからdescription_columns_poolを取得
    description_columns = st.multiselect(
        "説明欄に含める列（複数選択可）", 
        st.session_state.get('description_columns_pool', [])
    )

    # Google認証
    st.subheader("🔐 Google認証")
    # Firebaseのdbとauthオブジェクトをauthenticate_google_with_firestoreに渡す
    creds, user_id = authenticate_google_with_firestore(db, firebase_auth)

    if creds:
        # 認証されたユーザーIDを表示
        # st.write(f"現在のユーザーID (Firebase): `{user_id}`") # calendar_utilsで表示されるためコメントアウト
        try:
            service = build("calendar", "v3", credentials=creds)
            calendar_list = service.calendarList().list().execute()
            calendar_options = {cal['summary']: cal['id'] for cal in calendar_list['items']}
            
            if not calendar_options:
                st.error("利用可能なカレンダーが見つかりませんでした。Googleカレンダーの設定を確認してください。")
                st.stop()

            # selectboxのkeyをタブごとにユニークにする
            selected_calendar_name = st.selectbox("登録先カレンダーを選択", list(calendar_options.keys()), key="reg_calendar_select")
            calendar_id = calendar_options[selected_calendar_name]

            # データ処理と登録
            st.subheader("➡️ イベント登録")
            if st.button("Googleカレンダーに登録する", key="register_events_button"):
                with st.spinner("イベントデータを処理中..."):
                    df = process_excel_files(st.session_state['uploaded_files'], description_columns, all_day_event, private_event)
                    if df.empty:
                        st.warning("有効なイベントデータがありません。")
                    else:
                        st.info(f"{len(df)} 件のイベントを登録します。")
                        progress = st.progress(0)
                        successful_registrations = 0
                        for i, row in df.iterrows():
                            try:
                                if row['All Day Event'] == "True":
                                    # 終日イベントの場合、日付のみを使用
                                    start_date_str = datetime.strptime(row['Start Date'], "%Y/%m/%d").strftime("%Y-%m-%d")
                                    end_date_str = datetime.strptime(row['End Date'], "%Y/%m/%d").strftime("%Y-%m-%d")

                                    event_data = {
                                        'summary': row['Subject'],
                                        'location': row['Location'] if pd.notna(row['Location']) else '',
                                        'description': row['Description'] if pd.notna(row['Description']) else '',
                                        'start': {'date': start_date_str},
                                        'end': {'date': end_date_str},
                                        'transparency': 'transparent' if row['Private'] == "True" else 'opaque'
                                    }
                                else:
                                    # 時間指定イベントの場合、日付と時間を使用
                                    start_dt_str = f"{row['Start Date']} {row['Start Time']}"
                                    end_dt_str = f"{row['End Date']} {row['End Time']}"

                                    start = datetime.strptime(start_dt_str, "%Y/%m/%d %H:%M").isoformat()
                                    end = datetime.strptime(end_dt_str, "%Y/%m/%d %H:%M").isoformat()

                                    event_data = {
                                        'summary': row['Subject'],
                                        'location': row['Location'] if pd.notna(row['Location']) else '',
                                        'description': row['Description'] if pd.notna(row['Description']) else '',
                                        'start': {'dateTime': start, 'timeZone': 'Asia/Tokyo'},
                                        'end': {'dateTime': end, 'timeZone': 'Asia/Tokyo'},
                                        'transparency': 'transparent' if row['Private'] == "True" else 'opaque'
                                    }
                                add_event_to_calendar(service, calendar_id, event_data)
                                successful_registrations += 1
                            except Exception as e:
                                st.error(f"{row['Subject']} の登録に失敗しました: {e}")
                            progress.progress((i + 1) / len(df))

                        st.success(f"✅ {successful_registrations} 件のイベント登録が完了しました！")
        except Exception as e:
            st.error(f"カレンダーサービスの取得またはカレンダーリストの取得に失敗しました: {e}")
            st.warning("Google認証の状態を確認するか、ページをリロードしてください。")
    else:
        st.warning("Google認証が完了していません。")

with tabs[2]:
    st.header("イベントを削除")

    # Google認証 (削除機能も認証が必要)
    st.subheader("🔐 Google認証")
    # Firebaseのdbとauthオブジェクトをauthenticate_google_with_firestoreに渡す
    creds_del, user_id_del = authenticate_google_with_firestore(db, firebase_auth)

    if creds_del:
        # 認証されたユーザーIDを表示
        # st.write(f"現在のユーザーID (Firebase): `{user_id_del}`") # calendar_utilsで表示されるためコメントアウト
        try:
            service_del = build("calendar", "v3", credentials=creds_del)
            calendar_list_del = service_del.calendarList().list().execute()
            calendar_options_del = {cal['summary']: cal['id'] for cal in calendar_list_del['items']}

            if not calendar_options_del:
                st.error("利用可能なカレンダーが見つかりませんでした。Googleカレンダーの設定を確認してください。")
                st.stop()

            # selectboxのkeyをタブごとにユニークにする
            selected_calendar_name_del = st.selectbox("削除対象カレンダーを選択", list(calendar_options_del.keys()), key="del_calendar_select")
            calendar_id_del = calendar_options_del[selected_calendar_name_del]

            st.subheader("🗓️ 削除期間の選択")
            today = date.today()
            # デフォルトで過去30日間のイベントを対象にする
            default_start_date = today - timedelta(days=30)
            default_end_date = today

            delete_start_date = st.date_input("削除開始日", value=default_start_date, key="delete_start_date")
            delete_end_date = st.date_input("削除終了日", value=default_end_date, key="delete_end_date")

            if delete_start_date > delete_end_date:
                st.error("削除開始日は終了日より前に設定してください。")
            else:
                st.subheader("🗑️ 削除実行")
                # 削除実行ボタンの確認ダイアログ
                if st.button("選択期間のイベントを削除する", key="delete_events_trigger_button"):
                    st.warning(f"「**{selected_calendar_name_del}**」カレンダーから")
                    st.warning(f"**{delete_start_date.strftime('%Y年%m月%d日')}**から**{delete_end_date.strftime('%Y年%m月%d日')}**までの")
                    st.warning("**全てのイベントを削除します。この操作は元に戻せません。よろしいですか？**")
                    
                    # ユーザーに確認を促すための別のボタンを表示
                    if st.button("はい、削除を実行します", key="confirm_delete_action_button"):
                        deleted_count = delete_events_from_calendar(
                            service_del, calendar_id_del, 
                            datetime.combine(delete_start_date, datetime.min.time()),
                            datetime.combine(delete_end_date, datetime.max.time()) # 日付の終わりまで含める
                        )
                        if deleted_count > 0:
                            st.success(f"✅ {deleted_count} 件のイベントが削除されました。")
                        else:
                            st.info("指定された期間内に削除するイベントは見つかりませんでした。")
                    else:
                        st.info("削除はキャンセルされました。")

        except Exception as e:
            st.error(f"カレンダーサービスの取得またはカレンダーリストの取得に失敗しました: {e}")
            st.warning("Google認証の状態を確認するか、ページをリロードしてください。")
    else:
        st.warning("Google認証が完了していません。")
