import streamlit as st
import pandas as pd
from datetime import datetime, time
from excel_parser import process_excel_files
from calendar_utils import authenticate_google, add_event_to_calendar, delete_events_in_range
from config import SCOPES
from googleapiclient.discovery import build

st.set_page_config(page_title="Googleカレンダー登録ツール", layout="wide")
st.title("📅 Googleカレンダー一括イベント登録")

# タブ構成
tabs = st.tabs(["1. ファイルのアップロード", "2. イベントの設定・登録", "3. イベント一括削除"])

# ファイルアップロード
with tabs[0]:
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

# イベント登録
with tabs[1]:
    if not uploaded_files:
        st.info("先にExcelファイルをアップロードしてください。")
        st.stop()

    st.subheader("📌 イベント設定")
    all_day_event = st.checkbox("終日イベントとして登録", value=False)
    private_event = st.checkbox("非公開イベントとして登録", value=True)
    description_columns = st.multiselect("説明欄に含める列（複数選択可）", sorted(description_columns_pool))

    st.subheader("🔐 Google認証")
    creds = authenticate_google()

    if creds:
        service = build("calendar", "v3", credentials=creds)
        calendar_list = service.calendarList().list().execute()
        calendar_options = {
            cal['summary']: cal['id']
            for cal in calendar_list['items']
            if cal.get('accessRole') in ['owner', 'writer']
            and not 'holiday@group.v.calendar.google.com' in cal['id']
        }
        selected_calendar_name = st.selectbox("登録先カレンダーを選択", list(calendar_options.keys()))
        calendar_id = calendar_options[selected_calendar_name]

        st.subheader("📤 イベント登録")
        if st.button("Googleカレンダーに登録する"):
            with st.spinner("イベントデータを処理中..."):
                df = process_excel_files(uploaded_files, description_columns, all_day_event, private_event)
                if df.empty:
                    st.warning("有効なイベントデータがありません。")
                    st.stop()

                st.success(f"{len(df)} 件のイベントを登録します。")
                progress = st.progress(0)
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
                        link = add_event_to_calendar(service, calendar_id, event_data)
                    except Exception as e:
                        st.error(f"{row['Subject']} の登録に失敗しました: {e}")
                    progress.progress((i + 1) / len(df))

                st.success("✅ 登録が完了しました！")
    else:
        st.stop()

# イベント削除（キーワード付き）
with tabs[2]:
    st.subheader("🗑 イベント一括削除")
    creds = authenticate_google()
    if creds:
        service = build("calendar", "v3", credentials=creds)
        calendar_list = service.calendarList().list().execute()
        calendar_options = {
            cal['summary']: cal['id']
            for cal in calendar_list['items']
            if cal.get('accessRole') in ['owner', 'writer']
            and not 'holiday@group.v.calendar.google.com' in cal['id']
        }

        selected_calendar_name = st.selectbox("削除対象カレンダーを選択", list(calendar_options.keys()), key="del_calendar")
        calendar_id = calendar_options[selected_calendar_name]

        col1, col2 = st.columns(2)
        with col1:
            start_date = st.date_input("削除開始日", value=datetime.today())
        with col2:
            end_date = st.date_input("削除終了日", value=datetime.today())

        keyword = st.text_input("件名に含まれるキーワード（空欄で全件対象）", "")

        if st.button("⚠️ この条件でイベントを削除"):
            deleted = delete_events_in_range(
                service,
                calendar_id,
                datetime.combine(start_date, time.min),
                datetime.combine(end_date, time.max),
                keyword=keyword if keyword.strip() else None
            )
            st.success(f"{deleted} 件のイベントを削除しました。")
