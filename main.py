import streamlit as st
import pandas as pd
from datetime import datetime
from excel_parser import process_excel_files
from calendar_utils import authenticate_google, add_event_to_calendar
from config import SCOPES # SCOPESはconfigから読み込まれますが、calendar_utilsでも定義しているので注意
from googleapiclient.discovery import build

st.set_page_config(page_title="Googleカレンダー登録ツール", layout="wide")
st.title("📅 Googleカレンダー一括イベント登録")

# ファイルアップロード
tabs = st.tabs(["1. ファイルのアップロード", "2. イベントの設定・登録"])
with tabs[0]:
    uploaded_files = st.file_uploader("Excelファイルを選択（複数可）", type=["xlsx"], accept_multiple_files=True)

    description_columns_pool = set()
    if uploaded_files:
        for file in uploaded_files:
            try:
                # ファイルがアップロードされたら、その都度カラムプールを更新
                df_temp = pd.read_excel(file, engine="openpyxl")
                df_temp.columns = [str(c).strip() for c in df_temp.columns]
                description_columns_pool.update(df_temp.columns)
            except Exception as e:
                st.warning(f"{file.name} の読み込みに失敗しました: {e}")
        # セッションステートにdescription_columns_poolを保存
        st.session_state['description_columns_pool'] = list(description_columns_pool)
    elif 'description_columns_pool' not in st.session_state:
        st.session_state['description_columns_pool'] = [] # 初期化

with tabs[1]:
    # uploaded_filesはtabs[0]で設定されるため、tabs[1]で利用可能
    # ただし、Streamlitの再実行で値がリセットされる可能性があるため、session_stateに入れるのがより堅牢
    if 'uploaded_files_data' not in st.session_state:
        st.session_state['uploaded_files_data'] = []

    if uploaded_files:
        # アップロードされたファイルをセッションステートに保存（ファイルの実体ではなく、必要な情報のみ）
        # ただし、ファイルオブジェクト自体をセッションステートに保存するとエラーになる可能性があるので注意
        # ここではファイルの内容を処理する `process_excel_files` がファイルオブジェクトを受け取るため、
        # `uploaded_files` を直接使う。ただし、実際のファイルパスなどが必要な場合は工夫が必要。
        pass
    else:
        # ファイルがアップロードされていない場合は、以前のセッションステートから取得を試みる
        # 今回のロジックでは、ファイルアップローダーが常に最新のuploaded_filesを返すため、
        # この部分は直接uploaded_filesを参照する形でも問題ないことが多い
        if not st.session_state.get('uploaded_files_data'): # uploaded_files_dataが空なら停止
            st.info("先にExcelファイルをアップロードしてください。")
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
    creds = authenticate_google()

    if creds:
        try:
            service = build("calendar", "v3", credentials=creds)
            calendar_list = service.calendarList().list().execute()
            calendar_options = {cal['summary']: cal['id'] for cal in calendar_list['items']}
            selected_calendar_name = st.selectbox("登録先カレンダーを選択", list(calendar_options.keys()))
            calendar_id = calendar_options[selected_calendar_name]

            # データ処理と登録
            st.subheader("➡️ イベント登録")
            if st.button("Googleカレンダーに登録する"):
                with st.spinner("イベントデータを処理中..."):
                    # uploaded_filesを直接渡す
                    df = process_excel_files(uploaded_files, description_columns, all_day_event, private_event)
                    if df.empty:
                        st.warning("有効なイベントデータがありません。")
                        # st.stop() # ここで停止すると、成功メッセージが表示されないためコメントアウト
                    else:
                        st.success(f"{len(df)} 件のイベントを登録します。")
                        progress = st.progress(0)
                        successful_registrations = 0
                        for i, row in df.iterrows():
                            try:
                                if row['All Day Event'] == "True":
                                    start_date = datetime.strptime(row['Start Date'], "%Y/%m/%d").strftime("%Y-%m-%d")
                                    end_date = datetime.strptime(row['End Date'], "%Y/%m/%d").strftime("%Y-%m-%d")

                                    event_data = {
                                        'summary': row['Subject'],
                                        'location': row['Location'] if pd.notna(row['Location']) else '',
                                        'description': row['Description'] if pd.notna(row['Description']) else '',
                                        'start': {'date': start_date},
                                        'end': {'date': end_date},
                                        'transparency': 'transparent' if row['Private'] == "True" else 'opaque'
                                    }
                                else:
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
        # st.stop() # 認証が完了していない場合でも、ユーザーに操作を継続させるためにコメントアウト
