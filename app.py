import html
import json
import os
import re
import tempfile
import unicodedata
import zipfile
from datetime import date, datetime, time
from io import BytesIO

import pandas as pd
import streamlit as st


REQUIRED_COLS = [
    "Job",
    "MK",
    "ISO",
    "Operation Description1",
    "OperationCode",
    "CustomFieldName",
    "CustomFieldValue",
    "ItemCode",
    "StepOrder",
    "BomVersionId",
]

DATE_FIELD = "DateTermine"
DEFAULT_TIME = time(15, 0)
AUTO_FILL_FIELDS = {"diametre", "materiel", "employe1", "sch", "type", "posoudurecorrige"}
SOUD_GRID_FIELDS = [
    "Diametre",
    "Type",
    "PoSoudureCorrige",
    "Materiel",
    "SCH",
    "Employe_1",
]
MODE_NEW = "Nouveau document"
MODE_CONTINUE = "Continuer (reprendre un fichier non termine)"
RECENT_DIR_NAME = ".recent_sessions"
RECENT_INDEX_NAME = "index.json"
RECENT_LIMIT = 30


def has_value(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return False
    return str(val).strip() != ""


def load_excel(file_or_path):
    xls = pd.ExcelFile(file_or_path, engine="openpyxl")
    sheet_name = xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=sheet_name, engine="openpyxl")
    return df, sheet_name


def normalize_text(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip().lower()
    text = unicodedata.normalize("NFKD", text)
    return "".join(ch for ch in text if not unicodedata.combining(ch))


def normalize_key(value):
    text = normalize_text(value)
    return re.sub(r"[\s_]+", "", text)


def get_now_quebec():
    try:
        from zoneinfo import ZoneInfo

        return datetime.now(ZoneInfo("America/Toronto"))
    except Exception:
        return datetime.now()


def parse_diameter_value(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    text = str(value).strip().replace(",", ".")
    if not text:
        return None
    fraction_match = re.match(r"^\s*(\d+(?:\.\d+)?)\s*/\s*(\d+(?:\.\d+)?)\s*$", text)
    if fraction_match:
        num = float(fraction_match.group(1))
        den = float(fraction_match.group(2))
        if den == 0:
            return None
        return num / den
    number_match = re.search(r"[-+]?\d*\.?\d+", text)
    if number_match:
        try:
            return float(number_match.group(0))
        except ValueError:
            return None
    return None


def format_number(value):
    if value is None:
        return ""
    if abs(value - int(round(value))) < 1e-9:
        return str(int(round(value)))
    text = f"{value:.6f}".rstrip("0").rstrip(".")
    return text


def compute_posoudurecorrige(diam_text, type_text):
    if not diam_text or not type_text:
        return None
    type_code = str(type_text).strip().upper()
    diam_str = str(diam_text).strip()
    if type_code in {"BW", "SW", "THRD"}:
        return diam_str
    if type_code in {"SOB", "LET", "OLET"}:
        diam_val = parse_diameter_value(diam_str)
        if diam_val is None:
            return None
        return format_number(diam_val * 3)
    if type_code == "FIL":
        diam_val = parse_diameter_value(diam_str)
        if diam_val is None:
            return None
        if 0 <= diam_val <= 10:
            return "4"
        if 11 <= diam_val <= 24:
            return "12"
        if 25 <= diam_val <= 48:
            return "24"
        return None
    if type_code == "TW":
        diam_val = parse_diameter_value(diam_str)
        if diam_val is None:
            return None
        return format_number(diam_val * 0.5)
    return None


def inject_styles():
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Urbanist:wght@400;600;700&family=Space+Mono:wght@400;700&display=swap');
        :root {
            --bg1: #0c111a;
            --bg2: #121b2a;
            --card: rgba(18, 24, 34, 0.92);
            --accent: #f6c453;
            --accent-2: #58d3c7;
            --text: #f5f7fb;
            --muted: #9aa4b5;
            --border: rgba(255, 255, 255, 0.08);
        }
        html, body, [class*="css"] {
            font-family: 'Urbanist', sans-serif;
            color: var(--text);
        }
        .stApp {
            background:
                radial-gradient(1000px 600px at 8% -10%, rgba(246, 196, 83, 0.12), transparent 60%),
                radial-gradient(900px 500px at 92% -20%, rgba(88, 211, 199, 0.12), transparent 60%),
                linear-gradient(180deg, var(--bg1), var(--bg2));
        }
        html[data-theme="light"] {
            --bg1: #f4f6fb;
            --bg2: #e7edf7;
            --card: rgba(255, 255, 255, 0.92);
            --accent: #d37a00;
            --accent-2: #0f8f7f;
            --text: #0e1525;
            --muted: #5e6b82;
            --border: rgba(12, 16, 24, 0.12);
        }
        body[data-theme="light"] .stApp,
        html[data-theme="light"] .stApp {
            background:
                radial-gradient(900px 520px at 8% -10%, rgba(211, 122, 0, 0.14), transparent 60%),
                radial-gradient(900px 520px at 92% -20%, rgba(15, 143, 127, 0.14), transparent 60%),
                linear-gradient(180deg, var(--bg1), var(--bg2));
        }
        @media (prefers-color-scheme: light) {
            :root {
                --bg1: #f4f6fb;
                --bg2: #e7edf7;
                --card: rgba(255, 255, 255, 0.92);
                --accent: #d37a00;
                --accent-2: #0f8f7f;
                --text: #0e1525;
                --muted: #5e6b82;
                --border: rgba(12, 16, 24, 0.12);
            }
            .stApp {
                background:
                    radial-gradient(900px 520px at 8% -10%, rgba(211, 122, 0, 0.14), transparent 60%),
                    radial-gradient(900px 520px at 92% -20%, rgba(15, 143, 127, 0.14), transparent 60%),
                    linear-gradient(180deg, var(--bg1), var(--bg2));
            }
        }
        .block-container {
            max-width: 1180px;
            padding-top: 1.2rem;
            padding-bottom: 4rem;
        }
        .grid-tight div[data-testid="stVerticalBlock"] {
            gap: 0.25rem;
        }
        .grid-tight div[data-testid="stHorizontalBlock"] {
            gap: 0.4rem;
            align-items: start;
        }
        .grid-tight div[data-testid="stTextInput"] {
            margin-bottom: 0.1rem;
        }
        .grid-tight .joint-tag {
            height: 2.05rem;
            display: inline-flex;
            align-items: center;
            margin-bottom: 0.1rem;
        }
        .grid-tight .mini-label {
            margin-bottom: 0.15rem;
        }
        .hero {
            background: linear-gradient(135deg, rgba(244, 185, 66, 0.12), rgba(79, 209, 197, 0.08));
            border: 1px solid var(--border);
            border-radius: 18px;
            padding: 1rem 1.2rem;
            box-shadow: 0 18px 40px rgba(0, 0, 0, 0.35);
            margin-bottom: 1.2rem;
        }
        .hero-title {
            font-size: 1.9rem;
            font-weight: 700;
            letter-spacing: 0.4px;
            margin: 0;
        }
        .hero-sub {
            color: var(--muted);
            margin-top: 0.3rem;
        }
        .op-header {
            margin-top: 1.4rem;
            margin-bottom: 0.4rem;
            font-size: 1.1rem;
            font-weight: 700;
            color: var(--accent);
        }
        .field-header {
            display: inline-block;
            margin: 0.4rem 0 0.6rem;
            padding: 0.35rem 0.7rem;
            background: rgba(255, 255, 255, 0.05);
            border: 1px solid rgba(255, 255, 255, 0.08);
            border-radius: 999px;
            font-weight: 600;
            color: #f0f3ff;
        }
        .joint-tag {
            font-family: 'Space Mono', monospace;
            font-size: 0.85rem;
            background: rgba(79, 209, 197, 0.12);
            color: #d7fef7;
            padding: 0 0.5rem;
            border-radius: 10px;
            border: 1px solid rgba(79, 209, 197, 0.35);
            display: inline-flex;
            align-items: center;
            height: 2.05rem;
            margin-bottom: 0.2rem;
        }
        .mini-label {
            font-size: 0.75rem;
            text-transform: uppercase;
            letter-spacing: 0.08em;
            color: var(--muted);
            margin-bottom: 0.25rem;
        }
        div[data-testid="stHorizontalBlock"] {
            gap: 0.4rem;
            align-items: center;
        }
        .stTextInput input,
        .stDateInput input,
        .stTimeInput input,
        .stSelectbox select {
            background: rgba(12, 16, 24, 0.9);
            border: 1px solid rgba(255, 255, 255, 0.08);
            border-radius: 12px;
            height: 2.05rem;
            color: var(--text);
        }
        .stTextInput input:disabled {
            background: rgba(12, 16, 24, 0.9);
            color: #d7fef7;
            border: 1px solid rgba(255, 255, 255, 0.08);
            font-family: 'Space Mono', monospace;
            text-align: center;
        }
        div[data-testid="stTextInput"] {
            margin-bottom: 0.2rem;
        }
        .stTimeInput div[data-baseweb="select"] > div,
        .stSelectbox div[data-baseweb="select"] > div {
            background: rgba(12, 16, 24, 0.9);
            border: 1px solid rgba(255, 255, 255, 0.08);
            border-radius: 12px;
            color: var(--text);
        }
        .stTimeInput div[data-baseweb="select"] svg,
        .stSelectbox div[data-baseweb="select"] svg {
            color: var(--muted);
        }
        .stTextInput input:focus,
        .stDateInput input:focus,
        .stTimeInput input:focus,
        .stSelectbox select:focus {
            border-color: var(--accent);
            box-shadow: 0 0 0 2px rgba(244, 185, 66, 0.2);
        }
        html[data-theme="light"] .stTextInput input,
        html[data-theme="light"] .stDateInput input,
        html[data-theme="light"] .stTimeInput input,
        html[data-theme="light"] .stSelectbox select,
        body[data-theme="light"] .stTextInput input,
        body[data-theme="light"] .stDateInput input,
        body[data-theme="light"] .stTimeInput input,
        body[data-theme="light"] .stSelectbox select {
            background: rgba(255, 255, 255, 0.96);
            border: 1px solid rgba(12, 16, 24, 0.14);
            color: #0e1525;
        }
        html[data-theme="light"] .stTextInput input:focus,
        html[data-theme="light"] .stDateInput input:focus,
        html[data-theme="light"] .stTimeInput input:focus,
        html[data-theme="light"] .stSelectbox select:focus,
        body[data-theme="light"] .stTextInput input:focus,
        body[data-theme="light"] .stDateInput input:focus,
        body[data-theme="light"] .stTimeInput input:focus,
        body[data-theme="light"] .stSelectbox select:focus {
            border-color: var(--accent);
            box-shadow: 0 0 0 2px rgba(211, 122, 0, 0.2);
        }
        .stButton > button {
            border-radius: 12px;
            border: 1px solid rgba(255, 255, 255, 0.15);
            background: rgba(255, 255, 255, 0.06);
            color: var(--text);
            padding: 0.5rem 1.1rem;
            font-weight: 600;
        }
        .stButton > button:hover {
            border-color: var(--accent-2);
            box-shadow: 0 0 0 2px rgba(79, 209, 197, 0.2);
        }
        .stAlert {
            border-radius: 12px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def natural_sort_key_joint(value, orig_index):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return (1, orig_index)
    text = str(value).strip()
    match = re.match(r"^([A-Za-z]+)\s*(\d+)$", text)
    if match:
        return (0, match.group(1).lower(), int(match.group(2)))
    match = re.match(r"^([A-Za-z]+).*?(\d+)", text)
    if match:
        return (0, match.group(1).lower(), int(match.group(2)))
    return (1, orig_index)


def clean_df(df, drop_filled=True):
    df_clean = df.copy()
    df_clean["_orig_index"] = df_clean.index

    mask_filled = df_clean["CustomFieldValue"].apply(has_value)

    target_fields = {"diametre", "materiel", "posoudurecorrige", "sch", "type"}
    op_norm = df_clean["OperationCode"].apply(normalize_text)
    field_norm = df_clean["CustomFieldName"].apply(normalize_text)
    mask_ass = (op_norm == "ass") & (field_norm.isin(target_fields))

    if drop_filled:
        df_clean = df_clean[~(mask_filled | mask_ass)].copy()
    else:
        df_clean = df_clean[~mask_ass].copy()
    return df_clean


def job_has_missing(df_full, job_value):
    if df_full is None:
        return False
    df_check = clean_df(df_full, drop_filled=False)
    df_job = df_check[df_check["Job"] == job_value]
    if df_job.empty:
        return False
    return (~df_job["CustomFieldValue"].apply(has_value)).any()


def passthrough_df(df):
    df_copy = df.copy()
    df_copy["_orig_index"] = df_copy.index
    return df_copy


def format_datetime(date_value, time_value):
    month = date_value.strftime("%b")
    hour = time_value.hour
    minute = time_value.minute
    ampm = "AM" if hour < 12 else "PM"
    hour12 = hour % 12
    if hour12 == 0:
        hour12 = 12
    return f"{month} {date_value.day:02d} {date_value.year}  {hour12}:{minute:02d}{ampm}"


def parse_genius_datetime(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    text = str(value).strip()
    match = re.match(
        r"^([A-Za-z]{3})\s+(\d{1,2})\s+(\d{4})\s{2}(\d{1,2}):(\d{2})(AM|PM)$",
        text,
    )
    if not match:
        return None
    month_abbr = match.group(1).title()
    month_map = {
        "Jan": 1,
        "Feb": 2,
        "Mar": 3,
        "Apr": 4,
        "May": 5,
        "Jun": 6,
        "Jul": 7,
        "Aug": 8,
        "Sep": 9,
        "Oct": 10,
        "Nov": 11,
        "Dec": 12,
    }
    month = month_map.get(month_abbr)
    if not month:
        return None
    day = int(match.group(2))
    year = int(match.group(3))
    hour = int(match.group(4))
    minute = int(match.group(5))
    ampm = match.group(6)
    if ampm == "PM" and hour != 12:
        hour += 12
    if ampm == "AM" and hour == 12:
        hour = 0
    return date(year, month, day), time(hour, minute)


def apply_updates(df_full, updates):
    if not updates:
        return df_full
    df_updated = df_full.copy()
    for idx, value in updates.items():
        if value is None:
            df_updated.at[idx, "CustomFieldValue"] = ""
        else:
            df_updated.at[idx, "CustomFieldValue"] = value
    return df_updated


def save_to_disk(df_clean, path, sheet_name, original_columns):
    if not path:
        raise ValueError("Missing save path.")
    directory = os.path.dirname(path)
    if directory:
        os.makedirs(directory, exist_ok=True)
    export_df = df_clean.drop(columns=["_orig_index"], errors="ignore")
    if original_columns:
        export_df = export_df[original_columns]
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        export_df.to_excel(writer, sheet_name=sheet_name, index=False)


def export_bytes(df_clean, sheet_name, original_columns):
    export_df = df_clean.drop(columns=["_orig_index"], errors="ignore")
    if original_columns:
        export_df = export_df[original_columns]
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export_df.to_excel(writer, sheet_name=sheet_name, index=False)
    return output.getvalue()


def export_genius_bytes(df_clean, sheet_name, original_columns):
    genius_df = df_clean[df_clean["CustomFieldValue"].apply(has_value)].copy()
    genius_df = genius_df.drop(columns=["_orig_index"], errors="ignore")
    if original_columns:
        genius_df = genius_df[original_columns]
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        genius_df.to_excel(writer, sheet_name=sheet_name, index=False)
    return output.getvalue()


def export_genius_package(
    df_full,
    sheet_name,
    original_columns,
    base_name,
    total_rows_per_file=500,
):
    genius_df = df_full[df_full["CustomFieldValue"].apply(has_value)].copy()
    genius_df = genius_df.drop(columns=["_orig_index"], errors="ignore")
    if original_columns:
        genius_df = genius_df[original_columns]
    total_rows = len(genius_df)
    data_rows_per_file = max(1, total_rows_per_file - 1)
    if total_rows <= data_rows_per_file:
        data = export_bytes(genius_df, sheet_name, original_columns)
        name = f"genius_{base_name}" if base_name else "genius_export.xlsx"
        return data, name, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    base_root = os.path.splitext(base_name or "genius_export.xlsx")[0]
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        part = 1
        for start in range(0, total_rows, data_rows_per_file):
            chunk = genius_df.iloc[start : start + data_rows_per_file]
            chunk_bytes = export_bytes(chunk, sheet_name, original_columns)
            chunk_name = f"{base_root}_GENIUS_{part:02d}.xlsx"
            zipf.writestr(chunk_name, chunk_bytes)
            part += 1
    zip_name = f"{base_root}_GENIUS_split.zip"
    return zip_buffer.getvalue(), zip_name, "application/zip"


def unique_in_order(values):
    seen = set()
    ordered = []
    for val in values:
        if val is None or (isinstance(val, float) and pd.isna(val)):
            continue
        if val not in seen:
            seen.add(val)
            ordered.append(val)
    return ordered


def sorted_group(df_group):
    df_sorted = df_group.copy()
    df_sorted["_sort_key"] = df_sorted.apply(
        lambda row: natural_sort_key_joint(
            row["Operation Description1"], row["_orig_index"]
        ),
        axis=1,
    )
    df_sorted = df_sorted.sort_values(by="_sort_key", kind="mergesort")
    return df_sorted.drop(columns=["_sort_key"])


def build_ui(df_view, selected_job):
    updates = {}
    df_job = df_view[df_view["Job"] == selected_job].copy()
    if df_job.empty:
        st.info("Aucune ligne pour ce Job.")
        return updates
    total_rows = len(df_job)
    filled_rows = df_job["CustomFieldValue"].apply(has_value).sum()
    if total_rows:
        st.progress(
            filled_rows / total_rows,
            text=f"Progression: {filled_rows}/{total_rows} valeurs remplies",
        )

    op_codes = unique_in_order(df_job["OperationCode"].tolist())
    def find_row_index(df_field, field_name, joint_raw):
        if joint_raw is None or (isinstance(joint_raw, float) and pd.isna(joint_raw)):
            joint_mask = df_field["Operation Description1"].isna()
        else:
            joint_mask = df_field["Operation Description1"] == joint_raw
        field_mask = df_field["CustomFieldName"] == field_name
        match = df_field[joint_mask & field_mask]
        if match.empty:
            return None, None
        return match.index[0], match.iloc[0]["CustomFieldValue"]

    def get_state_value(idx):
        if idx is None:
            return None
        key = f"grid_{idx}"
        return st.session_state.get(key)

    def get_value_for(idx, fallback):
        if idx is None:
            return fallback
        if idx in updates:
            return updates[idx]
        state_val = get_state_value(idx)
        if has_value(state_val):
            return state_val
        return fallback

    for op_code in op_codes:
        st.markdown(
            f"<div class='op-header'>Operation {html.escape(str(op_code))}</div>",
            unsafe_allow_html=True,
        )
        df_op = df_job[df_job["OperationCode"] == op_code].copy()

        date_mask = df_op["CustomFieldName"].apply(
            lambda v: normalize_text(v) == normalize_text(DATE_FIELD)
        )
        df_date = df_op[date_mask]
        formatted_date = None
        if not df_date.empty:
            existing_value = (
                df_date["CustomFieldValue"]
                .dropna()
                .astype(str)
                .map(str.strip)
            )
            existing_value = existing_value[existing_value != ""].head(1)
            now = get_now_quebec().replace(second=0, microsecond=0)
            default_date = now.date()
            default_time = now.time()
            if not existing_value.empty:
                parsed = parse_genius_datetime(existing_value.iloc[0])
                if parsed:
                    default_date, default_time = parsed
            date_key = f"dt_date_{selected_job}_{op_code}"
            time_key = f"dt_time_{selected_job}_{op_code}"
            date_col, time_col = st.columns([2, 1])
            date_col.markdown("<div class='mini-label'>Date</div>", unsafe_allow_html=True)
            date_value = date_col.date_input(
                f"DateTermine - date ({op_code})",
                value=default_date,
                key=date_key,
                label_visibility="collapsed",
            )
            time_col.markdown("<div class='mini-label'>Heure</div>", unsafe_allow_html=True)
            time_value = time_col.time_input(
                f"DateTermine - heure ({op_code})",
                value=default_time,
                key=time_key,
                label_visibility="collapsed",
            )
            formatted_date = format_datetime(date_value, time_value)

        other_fields = df_op[~date_mask]
        field_names = unique_in_order(other_fields["CustomFieldName"].tolist())
        op_norm = normalize_text(op_code)
        grid_field_keys = set()
        if op_norm == "soud" and not other_fields.empty:
            field_map = {}
            for name in field_names:
                key = normalize_key(name)
                if key not in field_map:
                    field_map[key] = name
            grid_fields = [
                field_map[key]
                for key in (normalize_key(name) for name in SOUD_GRID_FIELDS)
                if key in field_map
            ]
            grid_field_keys = {normalize_key(name) for name in grid_fields}
            grid_df = other_fields[
                other_fields["CustomFieldName"].apply(
                    lambda v: normalize_key(v) in grid_field_keys
                )
            ]
            if grid_fields and not grid_df.empty:
                st.markdown(
                    "<div class='field-header'>SOUD - Saisie par joint</div>",
                    unsafe_allow_html=True,
                )
                joint_rows = []
                seen = set()
                for idx, row in grid_df.iterrows():
                    joint_raw = row["Operation Description1"]
                    joint_label = (
                        "(Sans joint)"
                        if joint_raw is None
                        or (isinstance(joint_raw, float) and pd.isna(joint_raw))
                        else str(joint_raw)
                    )
                    joint_key = joint_label
                    if joint_key in seen:
                        continue
                    seen.add(joint_key)
                    joint_rows.append(
                        {
                            "raw": joint_raw,
                            "label": joint_label,
                            "sort_key": natural_sort_key_joint(
                                joint_raw, row["_orig_index"]
                            ),
                        }
                    )
                joint_rows.sort(key=lambda item: item["sort_key"])

                pos_field_name = field_map.get("posoudurecorrige", "")
                diam_field_name = field_map.get("diametre", "")
                type_field_name = field_map.get("type", "")
                joint_index_map = {}
                for joint in joint_rows:
                    raw_key = joint["raw"]
                    diam_idx, diam_val = (
                        find_row_index(grid_df, diam_field_name, raw_key)
                        if diam_field_name
                        else (None, None)
                    )
                    type_idx, type_val = (
                        find_row_index(grid_df, type_field_name, raw_key)
                        if type_field_name
                        else (None, None)
                    )
                    pos_idx, pos_val = (
                        find_row_index(grid_df, pos_field_name, raw_key)
                        if pos_field_name
                        else (None, None)
                    )
                    joint_index_map[raw_key] = {
                        "diam_idx": diam_idx,
                        "diam_val": diam_val,
                        "type_idx": type_idx,
                        "type_val": type_val,
                        "pos_idx": pos_idx,
                        "pos_val": pos_val,
                    }

                st.markdown("<div class='grid-tight'>", unsafe_allow_html=True)
                grid_cols = st.columns([0.9] + [1] * len(grid_fields))
                with grid_cols[0]:
                    st.markdown(
                        "<div class='mini-label'>Joint</div>",
                        unsafe_allow_html=True,
                    )
                    for row_pos, joint in enumerate(joint_rows):
                        st.text_input(
                            f"joint_{selected_job}_{op_code}_{row_pos}",
                            value=str(joint["label"]),
                            key=f"joint_{selected_job}_{op_code}_{row_pos}",
                            disabled=True,
                            label_visibility="collapsed",
                        )

                for col_idx, field_name in enumerate(grid_fields):
                    with grid_cols[col_idx + 1]:
                        st.markdown(
                            f"<div class='mini-label'>{html.escape(str(field_name))}</div>",
                            unsafe_allow_html=True,
                        )
                        seed_value = None
                        prev_key = None
                        for row_pos, joint in enumerate(joint_rows):
                            row_idx, current_value = find_row_index(
                                grid_df, field_name, joint["raw"]
                            )
                            if row_idx is None:
                                st.text_input(
                                    f"{field_name} {joint['label']} missing",
                                    value="",
                                    disabled=True,
                                    label_visibility="collapsed",
                                    placeholder="N/A",
                                )
                                continue
                            key = f"grid_{row_idx}"
                            current_value = get_value_for(row_idx, current_value)
                            display_value = "" if not has_value(
                                current_value
                            ) else str(current_value)
                            if normalize_key(field_name) == "posoudurecorrige":
                                idx_info = joint_index_map.get(joint["raw"], {})
                                diam_idx = idx_info.get("diam_idx")
                                type_idx = idx_info.get("type_idx")
                                pos_idx = idx_info.get("pos_idx")
                                current_diam = get_value_for(
                                    diam_idx, idx_info.get("diam_val")
                                )
                                current_type = get_value_for(
                                    type_idx, idx_info.get("type_val")
                                )
                                current_pos = get_value_for(
                                    pos_idx, idx_info.get("pos_val")
                                )
                                if has_value(current_diam) and has_value(current_type):
                                    diam_text = str(current_diam).strip()
                                    type_text = str(current_type).strip().upper()
                                    calc_key = f"pos_calc_{row_idx}"
                                    prev = st.session_state.get(calc_key)
                                    should_compute = False
                                    if prev is None:
                                        if not has_value(current_pos):
                                            should_compute = True
                                        else:
                                            st.session_state[calc_key] = {
                                                "diam": diam_text,
                                                "type": type_text,
                                            }
                                    elif (
                                        prev.get("diam") != diam_text
                                        or prev.get("type") != type_text
                                    ):
                                        should_compute = True
                                    if should_compute:
                                        computed = compute_posoudurecorrige(
                                            diam_text, type_text
                                        )
                                        if computed is not None:
                                            display_value = str(computed)
                                            current_value = display_value
                                            st.session_state[key] = display_value
                                            updates[row_idx] = display_value
                                            st.session_state[calc_key] = {
                                                "diam": diam_text,
                                                "type": type_text,
                                            }
                            if row_pos > 0 and seed_value and not has_value(
                                current_value
                            ):
                                existing_state = st.session_state.get(key, "")
                                if not has_value(existing_state):
                                    st.session_state[key] = seed_value
                                    updates[row_idx] = seed_value
                            new_value = st.text_input(
                                f"{field_name} {row_idx}",
                                value=display_value,
                                key=key,
                                label_visibility="collapsed",
                                placeholder="A remplir",
                            )
                            if new_value != display_value:
                                updates[row_idx] = new_value
                            if row_pos == 0:
                                prev_key = f"prev_{key}"
                                prev_value = st.session_state.get(
                                    prev_key, display_value
                                )
                                if new_value.strip() and new_value.strip() != prev_value:
                                    seed_value = new_value.strip()
                        if prev_key:
                            st.session_state[prev_key] = seed_value or st.session_state.get(
                                prev_key, ""
                            )
                st.markdown("</div>", unsafe_allow_html=True)
        for field_name in field_names:
            if op_norm == "soud" and normalize_key(field_name) in grid_field_keys:
                continue
            st.markdown(
                f"<div class='field-header'>{html.escape(str(field_name))}</div>",
                unsafe_allow_html=True,
            )
            df_field = other_fields[other_fields["CustomFieldName"] == field_name]
            df_field = sorted_group(df_field)
            auto_fill = normalize_key(field_name) in AUTO_FILL_FIELDS
            rows = list(df_field.iterrows())
            if not rows:
                continue
            first_idx, first_row = rows[0]
            first_joint = first_row["Operation Description1"]
            if first_joint is None or (
                isinstance(first_joint, float) and pd.isna(first_joint)
            ):
                first_joint = "(Sans joint)"
            first_current_raw = first_row["CustomFieldValue"]
            first_display = "" if not has_value(first_current_raw) else str(
                first_current_raw
            )
            first_key = f"val_{first_idx}"
            col_left, col_right = st.columns([0.9, 5.1])
            col_left.markdown(
                f"<span class='joint-tag'>{html.escape(str(first_joint))}</span>",
                unsafe_allow_html=True,
            )
            first_new = col_right.text_input(
                f"{field_name} {first_idx}",
                value=first_display,
                key=first_key,
                label_visibility="collapsed",
                placeholder="Saisir une valeur",
            )
            if first_new != first_display:
                updates[first_idx] = first_new
            seed_value = None
            if auto_fill:
                first_new_stripped = first_new.strip()
                prev_key = f"prev_{first_key}"
                prev_value = st.session_state.get(prev_key, first_display)
                if first_new_stripped and first_new_stripped != prev_value:
                    seed_value = first_new_stripped

            for idx, row in rows[1:]:
                joint_label = row["Operation Description1"]
                if joint_label is None or (
                    isinstance(joint_label, float) and pd.isna(joint_label)
                ):
                    joint_label = "(Sans joint)"
                current_value = row["CustomFieldValue"]
                display_value = "" if not has_value(current_value) else str(
                    current_value
                )
                key = f"val_{idx}"
                if auto_fill and seed_value and not has_value(current_value):
                    existing_state = st.session_state.get(key, "")
                    if not has_value(existing_state):
                        st.session_state[key] = seed_value
                        updates[idx] = seed_value
                col_left, col_right = st.columns([0.9, 5.1])
                col_left.markdown(
                    f"<span class='joint-tag'>{html.escape(str(joint_label))}</span>",
                    unsafe_allow_html=True,
                )
                new_value = col_right.text_input(
                    f"{field_name} {idx}",
                    value=display_value,
                    key=key,
                    label_visibility="collapsed",
                    placeholder="Saisir une valeur",
                )
                if new_value != display_value:
                    updates[idx] = new_value
            if auto_fill:
                st.session_state[prev_key] = first_new
        has_data = False
        has_employe = False
        if not other_fields.empty:
            for idx, row in other_fields.iterrows():
                value = updates.get(idx, row["CustomFieldValue"])
                if has_value(value):
                    has_data = True
                    field_key = normalize_key(row["CustomFieldName"])
                    if field_key == "employe1":
                        has_employe = True
                    if has_data and has_employe:
                        break
        require_employe = normalize_text(op_code) == "soud"
        allow_date = has_data and (has_employe if require_employe else True)
        if formatted_date and allow_date:
            for idx in df_date.index:
                if df_view.at[idx, "CustomFieldValue"] != formatted_date:
                    updates[idx] = formatted_date
        st.divider()
    return updates


def try_tk_open_file():
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.update()
        path = filedialog.askopenfilename(
            title="Open Excel file",
            filetypes=[("Excel files", "*.xlsx")],
        )
        root.destroy()
        return path if path else None
    except Exception:
        return None


def try_tk_save_file(default_name):
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.update()
        path = filedialog.asksaveasfilename(
            title="Save Excel file",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel files", "*.xlsx")],
        )
        root.destroy()
        return path if path else None
    except Exception:
        return None


def normalize_save_path(path, default_name):
    if not path:
        return ""
    path = os.path.expanduser(path)
    if os.path.isdir(path):
        path = os.path.join(path, default_name)
    if not path.lower().endswith(".xlsx"):
        path += ".xlsx"
    return path


def build_extracteur_path(path):
    if not path:
        return ""
    base, ext = os.path.splitext(path)
    if base.endswith("_EXTRACTEUR"):
        return f"{base}{ext or '.xlsx'}"
    if not ext:
        ext = ".xlsx"
    return f"{base}_EXTRACTEUR{ext}"


def build_extracteur_name(filename):
    if not filename:
        return "export_EXTRACTEUR.xlsx"
    base, ext = os.path.splitext(filename)
    if not ext:
        ext = ".xlsx"
    if base.endswith("_EXTRACTEUR"):
        return f"{base}{ext}"
    return f"{base}_EXTRACTEUR{ext}"


def get_recent_dir():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    for root in (base_dir, os.getcwd(), tempfile.gettempdir()):
        try:
            path = os.path.join(root, RECENT_DIR_NAME)
            os.makedirs(path, exist_ok=True)
            return path
        except Exception:
            continue
    return ""


def load_recent_index():
    directory = get_recent_dir()
    if not directory:
        return []
    index_path = os.path.join(directory, RECENT_INDEX_NAME)
    if not os.path.exists(index_path):
        return []
    try:
        with open(index_path, "r", encoding="utf-8") as handle:
            data = json.load(handle)
        if isinstance(data, list):
            return data
        return []
    except Exception:
        return []


def save_recent_index(items):
    directory = get_recent_dir()
    if not directory:
        return
    index_path = os.path.join(directory, RECENT_INDEX_NAME)
    try:
        with open(index_path, "w", encoding="utf-8") as handle:
            json.dump(items, handle, indent=2)
    except Exception:
        return


def sanitize_filename(name):
    if not name:
        return "session"
    name = re.sub(r"[^A-Za-z0-9._-]+", "_", name)
    name = name.strip("._")
    return name or "session"


def recent_key(label, meta):
    label_key = normalize_text(label)
    meta = meta or {}
    project_key = normalize_text(meta.get("project_line", ""))
    creator_key = normalize_text(meta.get("creator", ""))
    return f"{label_key}::{project_key}::{creator_key}"


def next_session_id(items):
    max_id = 0
    for item in items:
        raw = str(item.get("session_id", "")).lstrip("S")
        if raw.isdigit():
            max_id = max(max_id, int(raw))
    return f"S{max_id + 1:03d}"


def save_recent_snapshot(df_full, sheet_name, original_columns, label, meta=None):
    directory = get_recent_dir()
    if not directory:
        return None
    safe_label = sanitize_filename(label)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{stamp}_{safe_label}.xlsx"
    path = os.path.join(directory, filename)
    try:
        save_to_disk(df_full, path, sheet_name, original_columns)
    except Exception:
        return None
    items = load_recent_index()
    key = recent_key(label, meta)
    existing_id = None
    filtered = []
    for item in items:
        if item.get("key") == key:
            existing_id = item.get("session_id")
            continue
        filtered.append(item)
    items = filtered
    item = {
        "session_id": existing_id or next_session_id(items),
        "label": label,
        "path": path,
        "saved_at": stamp,
        "key": key,
    }
    if isinstance(meta, dict):
        item["project_line"] = meta.get("project_line", "")
        item["creator"] = meta.get("creator", "")
        item["added_date"] = meta.get("added_date", "")
    items.insert(0, item)
    save_recent_index(items[:RECENT_LIMIT])
    return path


def init_session_state():
    defaults = {
        "df_full": None,
        "df_view": None,
        "sheet_name": None,
        "save_path": None,
        "auto_save_path": None,
        "selected_job": None,
        "pending_job": None,
        "updates": {},
        "original_columns": None,
        "loaded_source_id": None,
        "job_changed": False,
        "mode": None,
        "hide_filled_rows": True,
        "recent_path": None,
        "recent_session": None,
        "file_meta": {},
        "file_meta_source": None,
        "meta_project_line": "",
        "meta_creator": "",
        "meta_date": "",
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def main():
    st.set_page_config(page_title="Qualifab Genius Input", layout="wide")
    inject_styles()
    init_session_state()
    st.markdown(
        """
        <div class="hero">
            <div class="hero-title">Qualifab Data Entry Booster for poor Genius</div>
            <div class="hero-sub">Choisissez un mode, chargez un fichier Excel, puis renseignez les valeurs.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    mode = st.radio(
        "Mode de travail",
        [MODE_NEW, MODE_CONTINUE],
    )
    if st.session_state["mode"] != mode:
        st.session_state["mode"] = mode
        st.session_state["df_full"] = None
        st.session_state["df_view"] = None
        st.session_state["loaded_source_id"] = None
        st.session_state["selected_job"] = None
        st.session_state["updates"] = {}
        st.session_state["save_path"] = None
        st.session_state["auto_save_path"] = None
        st.session_state["hide_filled_rows"] = mode == MODE_NEW

    file_or_path = None
    source_id = None
    save_path = st.session_state.get("save_path")

    if mode == MODE_NEW:
        st.subheader("1) Charger un fichier brut")
        input_path = st.session_state.get("new_input_path")
        uploaded = st.file_uploader(
            "Uploader un fichier .xlsx",
            type=["xlsx"],
            key="uploader_new",
        )
        if input_path:
            file_or_path = input_path
            source_id = f"path::{os.path.abspath(input_path)}"
        elif uploaded is not None:
            file_or_path = uploaded
            source_id = f"upload::{uploaded.name}::{uploaded.size}"
        else:
            st.info("Selectionnez un fichier Excel pour commencer.")

        if file_or_path is not None:
            st.subheader("2) Choisir l'emplacement de sauvegarde")
            default_name = (
                os.path.basename(file_or_path)
                if isinstance(file_or_path, str)
                else uploaded.name
            )
            default_path = os.path.join(os.getcwd(), default_name)
            path_input = st.text_input(
                "Chemin complet du fichier de travail (.xlsx)",
                value=st.session_state.get("save_path") or default_path,
            )
            save_path = normalize_save_path(path_input, default_name)
            st.session_state["save_path"] = save_path
            st.session_state["auto_save_path"] = build_extracteur_path(save_path)
            st.caption(f"Dossier par defaut: {os.getcwd()}")

    else:
        st.subheader("0) Reouvrir une session recente")
        recent_sessions = load_recent_index()
        if recent_sessions:
            recent_rows = []
            options = []
            session_map = {}
            for session in recent_sessions:
                session_id = session.get("session_id") or "S000"
                label = session.get("label") or "session"
                stamp = session.get("saved_at") or ""
                proj = session.get("project_line", "")
                creator = session.get("creator", "")
                added = session.get("added_date", "")
                recent_rows.append(
                    {
                        "ID": session_id,
                        "Date": stamp,
                "Numero de projet": proj,
                        "Createur": creator,
                        "Ajoute": added,
                        "Fichier": label,
                    }
                )
                display = f"[{session_id}] {stamp} - {label}"
                options.append(display)
                session_map[display] = session
            st.dataframe(pd.DataFrame(recent_rows), use_container_width=True)
            choice = st.selectbox(
                "Sessions recentes (serveur)",
                options,
                key="recent_select",
            )
            if st.button("Charger la session selectionnee"):
                selected = session_map.get(choice)
                if selected and selected.get("path"):
                    st.session_state["cont_input_path"] = selected.get("path")
                    st.session_state["save_path"] = selected.get("path")
                    st.session_state["auto_save_path"] = build_extracteur_path(
                        selected.get("path")
                    )
                    st.session_state["meta_project_line"] = selected.get(
                        "project_line", ""
                    )
                    st.session_state["meta_creator"] = selected.get("creator", "")
                    st.session_state["meta_date"] = selected.get("added_date", "")
                    st.session_state["recent_session"] = None
                    st.rerun()
                else:
                    st.warning("Session introuvable sur le serveur.")
        else:
            st.caption("Aucune session recente disponible.")

        st.subheader("1) Reprendre un fichier existant")
        input_path = st.session_state.get("cont_input_path")
        uploaded = st.file_uploader(
            "Uploader un fichier .xlsx",
            type=["xlsx"],
            key="uploader_continue",
        )
        if input_path:
            file_or_path = input_path
            source_id = f"path::{os.path.abspath(input_path)}"
            save_path = input_path
            st.session_state["save_path"] = save_path
        elif uploaded is not None:
            file_or_path = uploaded
            source_id = f"upload::{uploaded.name}::{uploaded.size}"
            default_name = uploaded.name
            default_path = os.path.join(os.getcwd(), default_name)
            path_input = st.text_input(
                "Chemin de sauvegarde (.xlsx)",
                value=st.session_state.get("save_path") or default_path,
            )
            save_path = normalize_save_path(path_input, default_name)
            st.session_state["save_path"] = save_path
            st.session_state["auto_save_path"] = build_extracteur_path(save_path)
            st.caption(
                "Impossible de recuperer le chemin original depuis l'upload."
            )
        else:
            st.info("Choisissez un fichier ou une session recente pour continuer.")

    if file_or_path is not None:
        if st.session_state.get("file_meta_source") != source_id:
            st.session_state["meta_project_line"] = ""
            st.session_state["meta_creator"] = ""
            st.session_state["meta_date"] = get_now_quebec().date().isoformat()
            st.session_state["file_meta_source"] = source_id

        meta_col1, meta_col2 = st.columns(2)
        meta_col1.text_input("Numero de projet", key="meta_project_line")
        meta_col2.text_input("Createur", key="meta_creator")
        if not st.session_state.get("meta_date"):
            st.session_state["meta_date"] = get_now_quebec().date().isoformat()
        st.caption(f"Date: {st.session_state['meta_date']}")
        st.session_state["file_meta"] = {
            "project_line": st.session_state.get("meta_project_line", ""),
            "creator": st.session_state.get("meta_creator", ""),
            "added_date": st.session_state.get("meta_date", ""),
        }
        if (
            not st.session_state["meta_project_line"].strip()
            or not st.session_state["meta_creator"].strip()
        ):
            st.warning(
                "Veuillez remplir la ligne de projet et le createur avant de continuer."
            )
            return

    if file_or_path is not None and save_path:
        if st.session_state["loaded_source_id"] != source_id:
            try:
                df_raw, sheet_name = load_excel(file_or_path)
            except Exception as exc:
                st.error(f"Erreur lors du chargement: {exc}")
                return
            missing = [c for c in REQUIRED_COLS if c not in df_raw.columns]
            if missing:
                st.error(f"Colonnes manquantes: {', '.join(missing)}")
                return
            if mode == MODE_NEW:
                df_full = clean_df(df_raw, drop_filled=True)
            else:
                df_full = df_raw.copy()
                df_full["_orig_index"] = df_full.index
            st.session_state["df_full"] = df_full
            st.session_state["sheet_name"] = sheet_name
            st.session_state["original_columns"] = list(df_raw.columns)
            st.session_state["loaded_source_id"] = source_id
            st.session_state["selected_job"] = None
            st.session_state["updates"] = {}
            st.session_state["auto_save_path"] = build_extracteur_path(
                st.session_state["save_path"]
            )
            save_recent_snapshot(
                st.session_state["df_full"],
                st.session_state["sheet_name"],
                st.session_state["original_columns"],
                os.path.basename(st.session_state.get("save_path") or "session"),
                meta=st.session_state.get("file_meta"),
            )

        st.success(f"Fichier de travail: {st.session_state['save_path']}")
        st.info(f"Sauvegarde auto: {st.session_state['auto_save_path']}")

    df_full = st.session_state.get("df_full")
    if df_full is None:
        return

    if mode == MODE_CONTINUE:
        df_view = df_full
    else:
        df_view = clean_df(df_full, drop_filled=st.session_state["hide_filled_rows"])
    st.session_state["df_view"] = df_view

    jobs = unique_in_order(df_view["Job"].tolist())
    if not jobs:
        st.warning("Aucun Job disponible.")
        return

    if st.session_state["selected_job"] in jobs:
        default_index = jobs.index(st.session_state["selected_job"])
    else:
        default_index = 0

    if (
        st.session_state["selected_job"] is None
        or st.session_state["selected_job"] not in jobs
    ):
        st.session_state["selected_job"] = jobs[default_index]

    def on_job_change():
        new_value = st.session_state["job_select"]
        st.session_state["pending_job"] = new_value

    st.selectbox(
        "Job",
        jobs,
        index=jobs.index(st.session_state["selected_job"]),
        key="job_select",
        on_change=on_job_change,
    )

    current_job = st.session_state["selected_job"]
    pending_job = st.session_state.get("pending_job")
    if pending_job and pending_job != current_job:
        if job_has_missing(df_full, current_job):
            st.warning(
                "Des champs ne sont pas remplis pour ce Job. Voulez-vous vraiment changer?"
            )
            col_stay, col_leave = st.columns(2)
            if col_stay.button("Rester sur ce Job"):
                st.session_state["pending_job"] = None
                st.rerun()
            if col_leave.button("Changer quand meme"):
                st.session_state["selected_job"] = pending_job
                st.session_state["pending_job"] = None
                st.session_state["job_changed"] = True
                st.rerun()
        else:
            st.session_state["selected_job"] = pending_job
            st.session_state["pending_job"] = None
            st.session_state["job_changed"] = True

    selected_job = st.session_state["selected_job"]
    if mode != MODE_CONTINUE:
        st.checkbox(
            "Masquer les lignes deja remplies",
            key="hide_filled_rows",
            help="Decoche pour eviter que les lignes disparaissent pendant la saisie.",
        )

    updates = build_ui(df_view, selected_job)
    st.session_state["updates"] = updates
    df_updated = apply_updates(df_full, updates)
    st.session_state["df_full"] = df_updated
    if mode == MODE_CONTINUE:
        st.session_state["df_view"] = df_updated
    else:
        st.session_state["df_view"] = clean_df(
            df_updated, drop_filled=st.session_state["hide_filled_rows"]
        )

    if st.session_state.get("job_changed"):
        auto_path = st.session_state.get("auto_save_path")
        if auto_path:
            if mode == MODE_NEW:
                df_to_save = clean_df(st.session_state["df_full"], drop_filled=True)
            else:
                df_to_save = st.session_state["df_full"]
            save_to_disk(
                df_to_save,
                auto_path,
                st.session_state["sheet_name"],
                st.session_state["original_columns"],
            )
            save_recent_snapshot(
                st.session_state["df_full"],
                st.session_state["sheet_name"],
                st.session_state["original_columns"],
                os.path.basename(st.session_state.get("save_path") or "session"),
                meta=st.session_state.get("file_meta"),
            )
            st.success("Sauvegarde automatique effectuee.")
        else:
            st.warning("Chemin de sauvegarde auto manquant.")
        st.session_state["job_changed"] = False

    col_save, col_export, col_genius = st.columns(3)
    if col_save.button("Sauvegarder maintenant"):
        try:
            save_to_disk(
                st.session_state["df_full"],
                st.session_state["save_path"],
                st.session_state["sheet_name"],
                st.session_state["original_columns"],
            )
            save_recent_snapshot(
                st.session_state["df_full"],
                st.session_state["sheet_name"],
                st.session_state["original_columns"],
                os.path.basename(st.session_state.get("save_path") or "session"),
                meta=st.session_state.get("file_meta"),
            )
            st.success("Sauvegarde terminee.")
        except Exception as exc:
            st.error(f"Erreur de sauvegarde: {exc}")

    export_data = export_bytes(
        st.session_state["df_full"],
        st.session_state["sheet_name"],
        st.session_state["original_columns"],
    )
    if st.session_state.get("save_path"):
        export_base = os.path.basename(st.session_state["save_path"])
    else:
        export_base = "export.xlsx"
    export_name = f"NONTERMINE_{export_base}"
    col_export.download_button(
        "Exporter Excel (toutes les lignes)",
        data=export_data,
        file_name=export_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    genius_source = st.session_state["df_full"]
    genius_data, genius_name, genius_mime = export_genius_package(
        genius_source,
        st.session_state["sheet_name"],
        st.session_state["original_columns"],
        export_name,
        total_rows_per_file=500,
    )
    col_genius.download_button(
        "Exporter Genius",
        data=genius_data,
        file_name=genius_name,
        mime=genius_mime,
    )



if __name__ == "__main__":
    main()
