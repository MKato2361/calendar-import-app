import streamlit as st
import pandas as pd
import datetime
from excel_parser import process_excel_files
from calendar_utils import authenticate_google, add_event_to_calendar, delete_events_in_range
from googleapiclient.discovery import build

st.set_page_config(page_title="Googleカレンダー一括イベント登録", layout="wide")
st.title("📅 Googleカレンダー一括イベント登録")

tabs = st.tabs(["📁 ファイルのアップロード", "📌 イベントの設定・登録", "🗑 イベントの削除"])

# 認証・共通処理
with tabs[1], tabs[2]:
    creds = authenticate_google()
    service = build("calendar", "v3", credentials=creds) if creds else None
    calendar_id = None
    calendar_options = {}
    if service:
        calendars = service.calendarList().list().execute().get("items", [])
        calendar_options = {cal["summary"]: cal["id"] for cal in calendars if cal.get("accessRole") in ["owner", "writer"]}

# タブ①: ファイルアップロード
with tabs[0]:
    uploaded_files = st.file_uploader("Excelファイルを選択 (複数可)", type=["xlsx"], accept_multiple_files=True)

# タブ②: イベントの設定・登録
with tabs[1]:
    if uploaded_files:
        all_columns = set()
        for f in uploaded_files:
            df = pd.read_excel(f, engine="openpyxl")
            all_columns.update(df.columns)
        description_cols = st.multiselect("説明欄に含める列（複数選択可）", sorted(all_columns))
    else:
        description_cols = []

    st.checkbox("終日イベントとして登録", key="allday")
    st.checkbox("非公開イベントとして登録", key="private", value=True)

    if service and calendar_options:
        calendar_name = st.selectbox("登録先カレンダーを選択", list(calendar_options.keys()), key="regcal")
        calendar_id = calendar_options[calendar_name]

        if st.button("📤 Googleカレンダーに登録する"):
            df = process_excel_files(uploaded_files, description_cols, st.session_state.allday, st.session_state.private)
            success = 0
            for _, row in df.iterrows():
                try:
                    event = {
                        "summary": row["Subject"],
                        "location": row["Location"],
                        "description": row["Description"],
                        "start": {},
                        "end": {},
                        "transparency": "transparent" if row["Private"] == "True" else "opaque",
                    }
                    if row["All Day Event"] == "True":
                        event["start"] = {"date": row["Start Date"].replace("/", "-")}
                        end_date = datetime.datetime.strptime(row["End Date"], "%Y/%m/%d") + datetime.timedelta(days=1)
                        event["end"] = {"date": end_date.strftime("%Y-%m-%d")}
                    else:
                        event["start"] = {"dateTime": row["Start Date"] + "T" + row["Start Time"] + ":00", "timeZone": "Asia/Tokyo"}
                        event["end"] = {"dateTime": row["End Date"] + "T" + row["End Time"] + ":00", "timeZone": "Asia/Tokyo"}
                    add_event_to_calendar(service, calendar_id, event)
                    success += 1
                except Exception as e:
                    st.error(f"{row['Subject']} の登録に失敗しました: {e}")
            st.success(f"{success} 件のイベントを登録しました。")

# タブ③: イベントの削除
with tabs[2]:
    if service and calendar_options:
        del_calendar_name = st.selectbox("削除対象カレンダーを選択", list(calendar_options.keys()), key="delcal")
        del_calendar_id = calendar_options[del_calendar_name]

        col1, col2 = st.columns(2)
        with col1:
            del_start = st.date_input("削除開始日", datetime.date.today())
        with col2:
            del_end = st.date_input("削除終了日", datetime.date.today())

        if st.button("⚠️ この期間のイベントをすべて削除"):
            deleted = delete_events_in_range(
                service,
                del_calendar_id,
                datetime.datetime.combine(del_start, datetime.time.min),
                datetime.datetime.combine(del_end, datetime.time.max)
            )
            st.success(f"{deleted} 件のイベントを削除しました。")
