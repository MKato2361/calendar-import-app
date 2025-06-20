# excel_parser.py

import pandas as pd
import re
import datetime

def clean_mng_num(value):
    if pd.isna(value):
        return ""
    return str(value).replace("HK", "").replace("-", "").strip()

def format_description_value(val):
    if pd.isna(val):
        return ""
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    return str(val)

def find_closest_column(columns, keywords):
    for kw in keywords:
        for col in columns:
            if kw.lower() in str(col).lower():
                return col
    return None

def process_excel_files(uploaded_files, description_columns, all_day_event, private_event):
    dataframes = []

    for uploaded_file in uploaded_files:
        try:
            df = pd.read_excel(uploaded_file, engine="openpyxl")
            df.columns = [str(c).strip() for c in df.columns]
            mng_col = find_closest_column(df.columns, ["管理番号"])
            if mng_col:
                df["管理番号"] = df[mng_col].apply(clean_mng_num)
                dataframes.append(df)
        except Exception as e:
            print(f"ファイルの読み込みエラー: {uploaded_file.name} → {e}")

    if not dataframes:
        return pd.DataFrame()

    merged_df = dataframes[0]
    for df_to_merge in dataframes[1:]:
        merged_df = pd.merge(merged_df, df_to_merge, on="管理番号", how="outer")

    merged_df.drop_duplicates(subset="管理番号", inplace=True)

    name_col = find_closest_column(merged_df.columns, ["物件名"])
    start_col = find_closest_column(merged_df.columns, ["予定開始"])
    end_col = find_closest_column(merged_df.columns, ["予定終了"])
    addr_col = find_closest_column(merged_df.columns, ["住所", "所在地"])

    if not all([name_col, start_col, end_col]):
        return pd.DataFrame()

    merged_df = merged_df.dropna(subset=[start_col, end_col])

    output = []
    for _, row in merged_df.iterrows():
        mng = clean_mng_num(row["管理番号"])
        name = row.get(name_col, "")
        subj = f"{mng}{name}"

        try:
            start = pd.to_datetime(row[start_col])
            end = pd.to_datetime(row[end_col])
        except Exception:
            continue

        location = row.get(addr_col, "")
        if isinstance(location, str) and "北海道札幌市" in location:
            location = location.replace("北海道札幌市", "")

        description_parts = []
        for col in description_columns:
            if col in row:
                description_parts.append(format_description_value(row[col]))
        description = " / ".join(description_parts)

        output.append({
            "Subject": subj,
            "Start Date": start.strftime("%Y/%m/%d"),
            "Start Time": start.strftime("%H:%M"),
            "End Date": end.strftime("%Y/%m/%d"),
            "End Time": end.strftime("%H:%M"),
            "All Day Event": "True" if all_day_event else "False",
            "Description": description,
            "Location": location,
            "Private": "True" if private_event else "False"
        })

    return pd.DataFrame(output)
