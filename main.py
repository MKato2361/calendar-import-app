import streamlit as st
import pandas as pd
import datetime
from excel_parser import process_excel_files
from calendar_utils import authenticate_google, add_event_to_calendar, delete_events_in_range
from googleapiclient.discovery import build

st.set_page_config(page_title="Googleカレンダー一括イベント登録", layout="wide")
st.title("📅 Googleカレンダー一括イベント登録")

# ファイルアップロード
uploaded_files = st.file_uploader("Excelファイルを選択 (複数可)", type=["xlsx"], accept_multiple_files=True)

# イベント設定
st.header("📌 イベント設定")
all_day = st.checkbox("終日イベントとして登録")
private_event = st.checkbox("非公開イベントとして登録", value=True)

description_cols = []
if uploaded_files:
    all_columns = set()
    for f in uploaded_files:
        df = pd.read_excel(f, engine="openpyxl")
        all_columns.update(df.columns)
    description_cols = st.multiselect("説明欄に含める列（複数選択可）", sorted(all_columns))

# Google認証
st.header("🔐 Google認証")
creds = authenticate_google()
service = build("calendar", "v3", credentials=creds) if creds else None

# カレンダー選択
calendar_id = None
calendar_options = {}
if service:
    calendars = service.calendarList().list().execute().get("items", [])
    calendar_options = {cal["summary"]: cal["id"] for cal in calendars if cal.get("accessRole") in ["owner", "writer"]}
    calendar_name = st.selectbox("登録先カレンダーを選択", list(calendar_options.keys()) if calendar_options else [])
    calendar_id = calendar_options.get(calendar_name)

# イベント登録処理
if uploaded_files and calendar_id:
    if st.button("✅ イベント登録を実行"):
        df = process_excel_files(uploaded_files, description_cols, all_day, private_event)
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

# 削除モード
st.header("🗑 Googleカレンダーイベント削除")
with st.expander("削除対象の期間を指定"):
    del_calendar = st.selectbox("削除対象カレンダーを選択", list(calendar_options.keys()) if calendar_options else [], key="delcal")
    start = st.date_input("削除開始日", value=datetime.date.today())
    end = st.date_input("削除終了日", value=datetime.date.today() + datetime.timedelta(days=7))
    if st.button("⚠️ この期間のイベントをすべて削除") and del_calendar:
        del_id = calendar_options.get(del_calendar)
        deleted = delete_events_in_range(service, del_id, datetime.datetime.combine(start, datetime.time.min), datetime.datetime.combine(end, datetime.time.max))
        st.success(f"{deleted} 件のイベントを削除しました。")
