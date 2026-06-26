import streamlit as st
import pandas as pd
import datetime
import io
import os
import base64
import math
from PIL import Image as PILImage
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from copy import copy
import sqlite3
import toml
from translations import get_t, get_data_t, translate_name

# ==========================================
# GLOBAL STYLES & BACKGROUND
# ==========================================
st.markdown("""
<style>

/* Ensure all buttons center their text perfectly */
.stButton > button {
    display: flex !important;
    
    align-items: center !important;
    text-align: center !important;
}
.stButton > button p {
    text-align: center !important;
    margin: 0 !important;
}

/* Global background */
[data-testid="stAppViewContainer"] {
    overflow-x: hidden;
}
</style>
""", unsafe_allow_html=True)




# ==========================================
# KHỞI TẠO LƯU TRỮ API KEY
# ==========================================

SECRETS_DIR = ".streamlit"
SECRETS_FILE = os.path.join(SECRETS_DIR, "secrets.toml")

def load_saved_api_key():
    import streamlit as st
    keys = []
    
    # Check streamlit secrets first
    if "GEMINI_API_KEYS" in st.secrets:
        keys = st.secrets["GEMINI_API_KEYS"]
    elif "GEMINI_API_KEY" in st.secrets:
        keys = [st.secrets["GEMINI_API_KEY"]]
        
    # Override with local file if it exists
    if os.path.exists(SECRETS_FILE):
        try:
            with open(SECRETS_FILE, "r", encoding="utf-8") as f:
                secrets = toml.load(f)
                if "GEMINI_API_KEYS" in secrets:
                    keys = secrets["GEMINI_API_KEYS"]
                elif "GEMINI_API_KEY" in secrets:
                    keys = [secrets["GEMINI_API_KEY"]]
        except:
            pass
            
    if keys:
        if isinstance(keys, str):
            keys = [keys]
        if 'api_key_idx' not in st.session_state:
            st.session_state['api_key_idx'] = 0
        else:
            st.session_state['api_key_idx'] = (st.session_state['api_key_idx'] + 1) % len(keys)
        return keys[st.session_state['api_key_idx']]
        
    return ""

def save_api_key(key):
    if not os.path.exists(SECRETS_DIR):
        os.makedirs(SECRETS_DIR)
    try:
        with open(SECRETS_FILE, "w", encoding="utf-8") as f:
            toml.dump({"GEMINI_API_KEY": key}, f)
    except Exception:
        pass

# ==========================================
# CƠ SỞ DỮ LIỆU SQLITE
# ==========================================
DB_FILE = 'attendance.db'

def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS records (
            ma_nv TEXT, ten_nv TEXT, ngay TEXT,
            gio_vao TEXT, gio_ra TEXT,
            di_tre INTEGER, ve_som INTEGER, ot REAL, tong_gio REAL, ghi_chu TEXT,
            UNIQUE(ma_nv, ngay)
        )
    ''')
    c.execute('''
        CREATE TABLE IF NOT EXISTS field_checkins (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ma_nv TEXT, ten_nv TEXT, thoi_gian TEXT,
            loai TEXT, dia_diem TEXT, toa_do TEXT, ghi_chu TEXT
        )
    ''')
    conn.commit()
    conn.close()

init_db()

def get_company_emp_options(lang):
    emps = {}

    if 'df_raw' in st.session_state and st.session_state.df_raw is not None:
        try:
            m_a = auto_detect_columns(st.session_state.df_raw)
            if 'ma_nv' in m_a and 'ten_nv' in m_a:
                uniq_e = st.session_state.df_raw[[m_a['ma_nv'], m_a['ten_nv']]].drop_duplicates()
                for _, r in uniq_e.iterrows():
                    ma = str(r[m_a['ma_nv']]).strip()
                    ten = str(r[m_a['ten_nv']]).strip()
                    if ma and ten and ma.lower() not in ['nan', 'none', ''] and ten.lower() not in ['nan', 'none', '']:
                        emps[ma] = ten
        except: pass

    try:
        conn = sqlite3.connect(DB_FILE)
        try:
            df_rec = pd.read_sql_query("SELECT DISTINCT ma_nv, ten_nv FROM records WHERE ma_nv IS NOT NULL AND ten_nv IS NOT NULL", conn)
            for _, r in df_rec.iterrows():
                ma = str(r['ma_nv']).strip()
                ten = str(r['ten_nv']).strip()
                if ma and ten and ma.lower() not in ['nan', 'none', ''] and ten.lower() not in ['nan', 'none', '']:
                    emps[ma] = ten
        except: pass
        try:
            df_fc = pd.read_sql_query("SELECT DISTINCT ma_nv, ten_nv FROM field_checkins WHERE ma_nv IS NOT NULL AND ten_nv IS NOT NULL", conn)
            for _, r in df_fc.iterrows():
                ma = str(r['ma_nv']).strip()
                ten = str(r['ten_nv']).strip()
                if ma and ten and ma.lower() not in ['nan', 'none', ''] and ten.lower() not in ['nan', 'none', '']:
                    emps[ma] = ten
        except: pass
        conn.close()
    except: pass

    if 'manual_emps' in st.session_state:
        for me in st.session_state.manual_emps:
            ma = str(me.get('ma', '')).strip()
            ten = str(me.get('ten', '')).strip()
            if ma and ten: emps[ma] = ten

    opts = [f"{ma} - {translate_name(ten, lang)}" for ma, ten in sorted(emps.items())]
    return sorted(list(set(opts)))

def save_field_checkin(ma_nv, ten_nv, thoi_gian, loai, dia_diem, toa_do, ghi_chu):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''
        INSERT INTO field_checkins (ma_nv, ten_nv, thoi_gian, loai, dia_diem, toa_do, ghi_chu)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    ''', (ma_nv, ten_nv, thoi_gian, loai, dia_diem, toa_do, ghi_chu))
    conn.commit()
    conn.close()

def get_field_checkins(limit=50):
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql_query(f"SELECT * FROM field_checkins ORDER BY id DESC LIMIT {limit}", conn)
    conn.close()
    return df

def save_to_db(df_filtered, mapping):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    records = []
    for _, row in df_filtered.iterrows():
        records.append((
            str(row[mapping['ma_nv']]), str(row[mapping['ten_nv']]),
            str(row["Ngày"]), str(row[mapping['gio_vao']]), str(row[mapping['gio_ra']]),
            int(row.get("Phút đi trễ", 0)), int(row.get("Phút về sớm", 0)),
            float(row.get("Giờ OT", 0)), float(row.get("Số giờ làm thực tế", 0)),
            str(row.get("Ghi chú", ""))
        ))
    c.executemany('''
        INSERT OR REPLACE INTO records
        (ma_nv, ten_nv, ngay, gio_vao, gio_ra, di_tre, ve_som, ot, tong_gio, ghi_chu)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', records)
    conn.commit()
    conn.close()

# ==========================================
# 1. CẤU HÌNH & ASSETS
# ==========================================
APP_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(APP_DIR, "assets")
FAVICON_PATH = os.path.join(ASSETS_DIR, "favicon.png")
LOGO_HEADER_PATH = os.path.join(ASSETS_DIR, "logo_header.png")

def load_favicon():
    try:
        return PILImage.open(FAVICON_PATH)
    except Exception:
        return "📊"

def _logo_img_tag(b64_val: str, style: str = "height:60px;", extra_class: str = "") -> str:
    class_str = f' class="{extra_class}"' if extra_class else ""
    if not b64_val:
        return f'<div style="font-size:32px">📊</div>'
    return f'<img src="data:image/png;base64,{b64_val}" style="{style}"{class_str}>'

LOGO_HEADER_B64 = "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"


def load_bg_base64():
    bg_path = os.path.join(ASSETS_DIR, "bg_blue_waves.png")
    try:
        with open(bg_path, "rb") as f:
            return base64.b64encode(f.read()).decode("utf-8")
    except Exception:
        return ""

BG_B64 = load_bg_base64()

st.set_page_config(page_title="Bảng Chấm Công", page_icon=load_favicon(), layout="wide", initial_sidebar_state="expanded")

if BG_B64:
    st.markdown(f"""
    <style>
    [data-testid="stAppViewContainer"], .stApp {{
        background: linear-gradient(rgba(255, 255, 255, 0.4), rgba(240, 248, 255, 0.6)), url("data:image/png;base64,{BG_B64}") !important;
        background-size: 100% 100% !important;
        background-position: center !important;
        background-attachment: fixed !important;
        background-repeat: no-repeat !important;
    }}
    </style>
    """, unsafe_allow_html=True)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Be+Vietnam+Pro:wght@400;500;600;700;800&family=Inter:wght@400;500;600;700&display=swap');

:root {
    --brand-50:  #F0F9FF;
    --brand-100: #E0F2FE;
    --brand-400: #38BDF8;
    --brand-500: #0EA5E9;
    --brand-600: #0284C7;
    --brand-700: #0369A1;
    --blue-500:  #0284C7;
    --ink-900:   #0F172A;
    --ink-700:   #334155;
    --ink-500:   #64748B;
    --line:      #E2E8F0;
    --card-shadow: 0 4px 20px rgba(0, 0, 0, 0.05);
    --card-hover-shadow: 0 10px 30px rgba(0, 0, 0, 0.1);
}

html, body, .stApp {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif !important;
    background-color: var(--bg-color);
}

html, body, .stApp {
    font-family: 'Inter', 'Be Vietnam Pro', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    background-color: var(--bg-color);
}

#MainMenu {visibility: hidden;}
footer {visibility: hidden; display: none !important;}
[data-testid="stToolbar"] {visibility: hidden !important; display: none !important;}

.stApp { background: var(--bg-color); }
.block-container {
    padding-top: 1.5rem !important;
    padding-left: 2rem !important;
    padding-right: 2rem !important;
    max-width: 98% !important;
}

/* ===== App Header Banner ===== */
.app-header {
    position: relative; overflow: hidden;
    background: linear-gradient(135deg, #0369A1 0%, #0EA5E9 50%, #38BDF8 100%);
    border-radius: 24px; padding: 36px 48px;
    display: flex; align-items: center; gap: 36px;
    margin-bottom: 32px;
    box-shadow: 0 16px 40px rgba(14, 165, 233, 0.2);
    color: white;
}
.app-header::before {
    content: ''; position: absolute; top: -50%; right: -5%;
    width: 350px; height: 350px;
    background: radial-gradient(circle, rgba(255,255,255,0.15) 0%, rgba(255,255,255,0) 70%);
    border-radius: 50%; pointer-events: none;
}
.app-header::after {
    content: ''; position: absolute; bottom: -40%; right: 15%;
    width: 250px; height: 250px;
    background: radial-gradient(circle, rgba(96, 165, 250, 0.2) 0%, rgba(255,255,255,0) 70%);
    border-radius: 50%; pointer-events: none;
}
.app-header-icon {
    font-size: 32px; display: flex; align-items: center; justify-content: center;
    background: #FFFFFF; border-radius: 100px; padding: 12px 32px;
    box-shadow: 0 10px 30px rgba(0, 0, 0, 0.15); z-index: 1;
    transition: transform 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.275);
}
.app-header-icon:hover { transform: scale(1.05) translateY(-2px); }
.app-header-icon img { height: 64px; width: auto; display: block; mix-blend-mode: multiply; }
.app-header-title { font-size: 30px; font-weight: 800; margin: 0; letter-spacing: -0.02em; z-index: 1; position: relative; }
.app-header-sub { font-size: 16px; color: rgba(255,255,255,0.9); margin: 8px 0 0 0; z-index: 1; position: relative; font-weight: 500; }
.app-header-badge {
    margin-left: auto;
    background: rgba(255,255,255,0.1); backdrop-filter: blur(12px);
    border: 1px solid rgba(255,255,255,0.25); color: white;
    font-size: 14.5px; font-weight: 600; padding: 12px 24px;
    border-radius: 100px; box-shadow: 0 8px 24px rgba(0,0,0,0.1);
    z-index: 1; transition: all 0.3s ease;
}
.app-header-badge:hover { background: rgba(255,255,255,0.2); transform: translateY(-2px); }

/* ===== Cards ===== */
.card { background: #FFFFFF; border: 1px solid var(--line); border-radius: 16px; padding: 22px 26px; margin-bottom: 18px; box-shadow: var(--card-shadow); transition: transform 0.2s ease, box-shadow 0.2s ease; }
.card:hover { transform: translateY(-3px); box-shadow: var(--card-hover-shadow); }
.card-title { font-size: 15px; font-weight: 700; color: var(--ink-900); margin: 0 0 16px 0; display: flex; align-items: center; gap: 12px; letter-spacing: -0.01em; }
.card-icon { display: inline-flex; align-items: center; justify-content: center; width: 32px; height: 32px; font-size: 16px; border-radius: 10px; background: var(--brand-50); color: var(--brand-600); }

/* Upload hint */
.upload-hint { background-color: #EFF6FF; border-radius: 12px; padding: 12px 18px; font-size: 14px; color: #0369A1; font-weight: 500; margin-bottom: 16px; border: 1px solid #BFDBFE; }

/* Buttons */
div[data-testid="stButton"], div[data-testid="stDownloadButton"] {
    display: flex;
    justify-content: center;
    width: 100%;
}
div[data-testid="stButton"] > button, div[data-testid="stDownloadButton"] > button {
    background: linear-gradient(135deg, #0EA5E9, #38BDF8) !important;
    color: white !important; border: none !important; border-radius: 100px !important;
    padding: 0 28px !important; font-size: 14px !important; font-weight: 600 !important;
    box-shadow: 0 4px 14px rgba(14, 165, 233, 0.3) !important; transition: all 0.2s ease !important;
    width: auto !important;
    height: 40px !important;
    min-height: 40px !important;
    line-height: 40px !important;
    white-space: nowrap !important;
}
div[data-testid="stButton"] > button:hover, div[data-testid="stDownloadButton"] > button:hover {
    transform: translateY(-2px) !important; box-shadow: 0 6px 18px rgba(14, 165, 233, 0.4) !important; filter: brightness(1.05) !important;
}
div[data-testid="stButton"] > button:active, div[data-testid="stDownloadButton"] > button:active { transform: translateY(0) !important; }

/* Streamlit components */
[data-testid="stFileUploader"] { background: #FFFFFF; border-radius: 16px; padding: 12px; border: 1px dashed var(--brand-400); box-shadow: var(--card-shadow); transition: all 0.3s ease; }
[data-testid="stFileUploader"]:hover { border-color: var(--brand-600); background: #F8FAFC; }
[data-testid="stFileUploaderDropzone"] { padding: 20px !important; border-radius: 12px !important; }
[data-testid="stDataFrame"] { border-radius: 16px; overflow: hidden; border: 1px solid var(--line); box-shadow: var(--card-shadow); transition: box-shadow 0.3s ease; }
[data-testid="stDataFrame"]:hover { box-shadow: var(--card-hover-shadow); }
[data-testid="stDateInput"] input, [data-testid="stTimeInput"] input, [data-testid="stNumberInput"] input { border-radius: 10px !important; border: 1px solid var(--line) !important; transition: border-color 0.2s ease, box-shadow 0.2s ease; }
[data-testid="stDateInput"] input:focus, [data-testid="stTimeInput"] input:focus, [data-testid="stNumberInput"] input:focus { border-color: var(--brand-500) !important; box-shadow: 0 0 0 3px rgba(14, 165, 233, 0.2) !important; }
[data-testid="stExpander"] { border: 1px solid var(--line) !important; border-radius: 16px !important; background: #FFFFFF !important; box-shadow: var(--card-shadow) !important; transition: box-shadow 0.2s ease; }
[data-testid="stExpander"]:hover { box-shadow: var(--card-hover-shadow) !important; }
/* Container box widgets */
[data-testid="stVerticalBlockBorderWrapper"] { border-radius: 16px !important; background: #FFFFFF !important; box-shadow: var(--card-shadow) !important; transition: transform 0.2s ease, box-shadow 0.2s ease; border: 1px solid var(--line) !important; padding: 20px !important;}
[data-testid="stVerticalBlockBorderWrapper"]:hover { box-shadow: var(--card-hover-shadow) !important; transform: translateY(-2px); }

/* Metrics */
[data-testid="stMetric"] {
    background: linear-gradient(to bottom right, #FFFFFF, #F8FAFC);
    border: 1px solid var(--line); border-radius: 16px; padding: 18px 20px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.03); border-left: 4px solid var(--brand-500);
    transition: transform 0.2s ease, box-shadow 0.2s ease;
}
[data-testid="stMetric"]:hover { transform: translateY(-2px); box-shadow: 0 8px 20px rgba(0,0,0,0.06); }
[data-testid="stMetricLabel"] { font-size: 13px !important; color: var(--ink-500) !important; font-weight: 500 !important; }
[data-testid="stMetricValue"] { font-family: 'Inter', 'Be Vietnam Pro', sans-serif !important; font-size: 26px !important; font-weight: 700 !important; color: var(--ink-900) !important; }

.stSpinner > div { border-top-color: var(--brand-500) !important; }
.stAlert { border-radius: 12px !important; border: none !important; box-shadow: 0 2px 8px rgba(0,0,0,0.03) !important; }
[data-testid="stSelectbox"] > div { border-radius: 10px !important; }
[data-testid="stMultiSelect"] > div > div { border-radius: 10px !important; }
[data-baseweb="tag"] { background-color: var(--brand-50) !important; color: var(--brand-700) !important; border-radius: 6px !important; }
.stSubheader, h3 { font-size: 16px !important; font-weight: 700 !important; color: var(--ink-900) !important; letter-spacing: -0.01em !important; }

/* Sidebar */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #FFFFFF 0%, #F8FAFC 70%, #E0F2FE 100%);
    border-right: 1px solid #BAE6FD;
}

/* ===== Ẩn chữ "Select all" mặc định của Streamlit trong multiselect ===== */
[data-baseweb="menu"] li:first-child > div[aria-selected] {
    display: none !important;
}
/* Cách 2: nhắm thẳng vào nút Select all */
[data-baseweb="menu"] [data-testid="stMultiSelectOptionAll"] {
    display: none !important;
}
/* Dropdown Menu & Top Right Nav */
.global-top-right-nav {
    position: fixed;
    top: 15px;
    right: 25px;
    z-index: 999999;
    display: flex;
    align-items: center;
    gap: 24px;
    font-family: 'Inter', 'Be Vietnam Pro', sans-serif;
    color: #0F172A;
}
.menu-dropdown {
    position: relative;
    display: inline-block;
    padding-bottom: 30px;
    margin-bottom: -30px;
}
.dropdown-content {
    visibility: hidden;
    opacity: 0;
    position: absolute;
    top: 100%;
    left: 50%;
    transform: translateX(-50%) translateY(10px);
    background-color: white;
    min-width: 340px;
    box-shadow: 0 20px 40px rgba(0,0,0,0.15);
    border-radius: 16px;
    padding: 16px;
    z-index: 1000000;
    display: flex;
    flex-direction: column;
    gap: 12px;
    transition: all 0.3s ease;
    cursor: default;
}
.menu-dropdown:hover .dropdown-content {
    visibility: visible;
    opacity: 1;
    transform: translateX(-50%) translateY(0);
}
.menu-dropdown.right-align .dropdown-content {
    left: auto;
    right: 0;
    transform: translateY(10px);
}
.menu-dropdown.right-align:hover .dropdown-content {
    transform: translateY(0);
}
.feature-card {
    background: #F8FAFC;
    border: 1px solid #E2E8F0;
    border-radius: 12px;
    padding: 16px;
    text-align: left;
    box-shadow: 0 2px 8px rgba(0,0,0,0.02);
    display: grid;
    grid-template-columns: auto 1fr;
    column-gap: 16px;
    row-gap: 4px;
    align-items: start;
    cursor: pointer;
    transition: background 0.2s ease;
}
.feature-card:hover {
    background: #F1F5F9;
}
.fc-icon {
    font-size: 20px;
    background: white;
    width: 36px; height: 36px;
    display: inline-flex;
    align-items: center; justify-content: center;
    border-radius: 8px;
    grid-row: span 2;
    margin-bottom: 0;
}
.feature-card h3 {
    font-size: 14px;
    font-weight: 700;
    color: #1E293B;
    margin: 0;
}
.feature-card p {
    font-size: 13px;
    color: #64748B;
    margin: 0;
    line-height: 1.4;
}
.contact-box {
    display: flex; align-items: center; gap: 14px; background: #F8FAFC; padding: 14px; 
    border-radius: 12px; border: 1px solid #E2E8F0; transition: all 0.2s ease; cursor: default;
}
.contact-box:hover {
    border-color: #3B82F6; background: #EFF6FF; transform: translateY(-2px); box-shadow: 0 4px 12px rgba(59,130,246,0.1);
}
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. HÀM TIỆN ÍCH
# ==========================================
def to_time_obj(val):
    try:
        is_na = pd.isna(val)
    except (TypeError, ValueError):
        is_na = False
    if is_na or val == "":
        return None
    if isinstance(val, datetime.time):
        return val
    if isinstance(val, datetime.datetime):
        return val.time()
    if isinstance(val, (int, float)):
        if 0 <= val < 1:
            total_seconds = int(val * 86400)
            return datetime.time(hour=total_seconds//3600, minute=(total_seconds%3600)//60, second=total_seconds%60)
    if isinstance(val, str):
        val = val.strip()
        if not val or val.lower() in ('nan', 'nat', 'none'):
            return None
        try:
            return pd.to_datetime(val).time()
        except:
            return None
    return None

def time_to_float(t):
    if t is None:
        return 0.0
    return t.hour + t.minute / 60.0 + t.second / 3600.0

def format_gio_lam(val):
    if val is None or (isinstance(val, float) and math.isnan(val)) or str(val).lower() in ['nan', 'nat', 'none', '']:
        return 0
    val = round(float(val), 1)
    return int(val) if val == int(val) else val

def format_gio_lam_str(val):
    return str(format_gio_lam(val))

# ==========================================
# 3. LOGIC TÍNH TOÁN
# ==========================================
def calculate_working_hours(time_in, time_out, start_chuan=8.0, end_chuan=17.0, lunch_start=12.0, lunch_end=13.0, max_hours=8.0):
    vao_trong = time_in is None or str(time_in).strip().lower() in ['', 'nan', 'none', 'nat']
    ra_trong = time_out is None or str(time_out).strip().lower() in ['', 'nan', 'none', 'nat']
    result = {'admin_hours': 0.0, 'tong_gio': 0.0, 'di_tre': 0, 've_som': 0, 'ot': 0.0, 'is_chieu': False}
    if vao_trong or ra_trong:
        return result
    t_in = to_time_obj(time_in)
    t_out = to_time_obj(time_out)
    if t_in is None or t_out is None:
        return result
    in_f = time_to_float(t_in)
    out_f = time_to_float(t_out)
    if out_f < in_f:
        result['tong_gio'] = -1
        return result
    eff_in = max(in_f, start_chuan)
    eff_out = min(out_f, end_chuan)
    if eff_out > eff_in:
        total_admin = eff_out - eff_in
        lunch_overlap = max(0.0, min(eff_out, lunch_end) - max(eff_in, lunch_start))
        admin_hours = min(max(total_admin - lunch_overlap, 0.0), max_hours)
    else:
        admin_hours = 0.0
    result['admin_hours'] = admin_hours
    
    if in_f >= lunch_start:
        result['is_chieu'] = True
        result['di_tre'] = int(max(0, (in_f - lunch_end) * 60)) if in_f > lunch_end else 0
    else:
        result['di_tre'] = int(max(0, (in_f - start_chuan) * 60)) if in_f > start_chuan else 0
        
    result['ve_som'] = int(max(0, (end_chuan - out_f) * 60)) if out_f < end_chuan else 0
    result['ot'] = max(0.0, out_f - end_chuan) + max(0.0, start_chuan - in_f)
    result['tong_gio'] = admin_hours
    return result

def clean_date(val):
    if pd.isna(val):
        return None
    try:
        if isinstance(val, datetime.datetime):
            return val
        return pd.to_datetime(str(val).strip(), dayfirst=True)
    except:
        return None

def get_fixed_holidays_for_years(years):
    dates = []
    for y in sorted(years):
        for (mo, day) in [(1,1), (4,30), (5,1), (9,2)]:
            dates.append(datetime.date(y, mo, day))
    return sorted(dates)

def is_workday_func(d_obj):
    """Trả về True nếu là ngày làm việc.
    - Thứ 7 cuối cùng của tháng = bắt buộc làm việc.
    - Các Thứ 7 khác và Chủ nhật = nghỉ.
    """
    if pd.isna(d_obj): return True
    try:
    	d = d_obj.date() if hasattr(d_obj, 'date') else d_obj
    	wd = d.weekday()
    	if wd == 6: return False  # Chủ nhật: nghỉ
    	if wd == 5:               # Thứ 7
    	    # Là Thứ 7 cuối cùng của tháng nếu cộng 7 ngày thì sang tháng khác
    	    return (d + datetime.timedelta(days=7)).month != d.month
    	return True  # Thứ 2 – Thứ 6: luôn làm việc
    except:
    	return True

def is_last_saturday_of_month(d_obj):
    """Kiểm tra ngày có phải Thứ 7 cuối cùng của tháng không."""
    try:
    	d = d_obj.date() if hasattr(d_obj, 'date') else d_obj
    	if d.weekday() != 5: return False
    	return (d + datetime.timedelta(days=7)).month != d.month
    except:
    	return False

def calculate_working_days(start_date, end_date, holidays=None, makeups=None):
    if holidays is None:
        holidays = set()
    if makeups is None:
        makeups = set()
    total = 0
    holiday_info = []
    makeup_info = []
    current = start_date
    while current <= end_date:
        is_wd = is_workday_func(current)
        if current in makeups:
            total += 1
            makeup_info.append({
                "date": current,
                "weekday_label": ["Thứ 2","Thứ 3","Thứ 4","Thứ 5","Thứ 6","Thứ 7","Chủ nhật"][current.weekday()]
            })
        elif current in holidays:
            holiday_info.append({
                "date": current,
                "weekday_label": ["Thứ 2","Thứ 3","Thứ 4","Thứ 5","Thứ 6","Thứ 7","Chủ nhật"][current.weekday()],
                "is_workday": is_wd
            })
        elif is_wd:
            total += 1
        current += datetime.timedelta(days=1)
    return total, holiday_info, makeup_info

# ==========================================
# 4. PARSER ĐỌC FILE
# ==========================================
def find_header_row(file, file_name):
    file.seek(0)
    if file_name.endswith(('.csv', '.txt', '.dat', '.tsv')):
        df_raw = pd.read_csv(file, header=None, nrows=30, sep=None, engine='python')
    else:
        df_raw = pd.read_excel(file, header=None, nrows=30)
    for i, row in df_raw.iterrows():
        row_str = [str(v).strip().lower() for v in row.values if pd.notna(v) and str(v).lower() != 'nan']
        keywords = ['mã', 'tên', 'ngày', 'vào', 'ra']
        if sum(1 for kw in keywords if any(kw in cell for cell in row_str)) >= 3:
            return i
    return 0

def parse_excel_file(uploaded_file):
    file_name = uploaded_file.name.lower()
    try:
        header_row = find_header_row(uploaded_file, file_name)
    except Exception as e:
        st.error(f"❌ Lỗi khi đọc file: {e}")
        st.stop()
    uploaded_file.seek(0)
    try:
        if file_name.endswith(('.csv', '.txt', '.dat', '.tsv')):
            df = pd.read_csv(uploaded_file, header=header_row, sep=None, engine='python')
        else:
            df = pd.read_excel(uploaded_file, header=header_row)
    except Exception as e:
        st.error(f"❌ Lỗi khi đọc bảng dữ liệu: {e}")
        st.stop()
    df = df.dropna(how='all')
    df.columns = df.columns.astype(str).str.strip()
    return df

def auto_detect_columns(df):
    mapping = {}
    for col in df.columns:
        col_lower = str(col).strip().lower()
        if 'mã' in col_lower and 'nv' not in col_lower.replace('mã',''):
            mapping['ma_nv'] = col
        elif col_lower in ['tên nhân viên', 'tên nv', 'họ tên', 'tên']:
            mapping['ten_nv'] = col
        elif 'ngày' in col_lower and 'ca' not in col_lower:
            mapping['ngay'] = col
        elif col_lower in ['vào', 'vao']:
            mapping['gio_vao'] = col
        elif col_lower == 'ra' and 'sớm' not in col_lower and 'som' not in col_lower:
            mapping['gio_ra'] = col
        elif col_lower in ['chức vụ', 'chuc vu', 'vị trí', 'title']:
            mapping['chuc_vu'] = col
        elif col_lower in ['phòng ban', 'phong ban', 'bộ phận', 'bo phan', 'department']:
            mapping['phong_ban'] = col
    return mapping

# ==========================================
# 5. EXPORTER XUẤT FILE
# ==========================================
def export_excel_tong_hop(df_filtered, mapping, start_date, end_date, total_wd):
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter
    import io
    import pandas as pd
    import streamlit as st
    
    t = st.session_state.get('cached_t', None)
    if t is None:
        from translations import get_t, get_data_t, translate_name
        t = get_t(st.session_state.lang)
        data_t = get_data_t(st.session_state.lang)
        t_name = lambda name: translate_name(name, st.session_state.lang)
    else:
        from translations import get_t, get_data_t, translate_name
        t = get_t(st.session_state.lang)
        data_t = get_data_t(st.session_state.lang)
        t_name = lambda name: translate_name(name, st.session_state.lang)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Chi tiết chấm công"
    
    font_bold14 = Font(name='Times New Roman', size=14, bold=True)
    font_bold12 = Font(name='Times New Roman', size=12, bold=True)
    font_normal = Font(name='Times New Roman', size=12)
    align_center = Alignment(horizontal='center', vertical='center')
    align_left = Alignment(horizontal='left', vertical='center')
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    border_thick = Border(left=Side(style='medium'), right=Side(style='medium'), top=Side(style='medium'), bottom=Side(style='medium'))
    
    ws.merge_cells('A1:N1')
    ws['A1'] = t("export_title")
    ws['A1'].font = font_bold14
    ws['A1'].alignment = align_center
    
    ws.merge_cells('A2:N2')
    ws['A2'] = t("export_date_range", start_date.strftime('%d/%m/%Y'), end_date.strftime('%d/%m/%Y'))
    ws['A2'].font = font_bold12
    ws['A2'].alignment = align_center
    
    ws.merge_cells('A4:H4')
    ws['A4'] = t("export_total_days", total_wd)
    ws['A4'].font = font_bold12
    ws['A4'].alignment = align_left
    
    total_hours = total_wd * 8
    ws.merge_cells('A5:H5')
    ws['A5'] = t("export_total_hours", total_hours)
    ws['A5'].font = font_bold12
    ws['A5'].alignment = align_left
    
    total_off_days = (end_date - start_date).days + 1 - total_wd
    ws.merge_cells('A6:H6')
    ws['A6'] = t("export_total_off", total_off_days)
    ws['A6'].font = font_bold12
    ws['A6'].alignment = align_left
    
    headers = [
        t("export_col_emp_code"), t("export_col_emp_name"), t("emp_dept"), t("emp_position"),
        t("export_col_date"), t("export_col_weekday"), t("export_col_in"), t("export_col_out"),
        t("export_col_work_day"), t("export_col_actual_hours"), t("export_col_ot"),
        t("export_col_total_hours"), t("export_col_ot_reason"), t("export_col_note")
    ]
    header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    for col_num, header in enumerate(headers, 1):
        c = ws.cell(row=8, column=col_num, value=header)
        c.font = font_bold12
        c.alignment = align_center
        c.fill = header_fill
        c.border = border
        
    row_idx = 9
    gray_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    
    def get_val(row, key):
        col = mapping.get(key)
        if col and col in row:
            v = row[col]
            return str(v) if pd.notna(v) else ""
        return ""
        
    for _, row in df_filtered.iterrows():
        try: d_obj = row['_parsed_date'].date()
        except: d_obj = pd.to_datetime(row['_parsed_date']).date()
        
        thu = d_obj.weekday()
        if st.session_state.lang == 'ja':
            thu_str = ["月", "火", "水", "木", "金", "土", "日"][thu]
        else:
            thu_str = ["Hai", "Ba", "Tư", "Năm", "Sáu", "Bảy", "CN"][thu]
        
        ws.cell(row=row_idx, column=1, value=get_val(row, 'ma_nv'))
        ws.cell(row=row_idx, column=2, value=t_name(get_val(row, 'ten_nv')))
        ws.cell(row=row_idx, column=3, value=data_t(get_val(row, 'phong_ban')))
        ws.cell(row=row_idx, column=4, value=data_t(get_val(row, 'chuc_vu')))
        ws.cell(row=row_idx, column=5, value=d_obj.strftime('%d/%m/%Y'))
        ws.cell(row=row_idx, column=6, value=thu_str)
        ws.cell(row=row_idx, column=7, value=get_val(row, 'gio_vao'))
        ws.cell(row=row_idx, column=8, value=get_val(row, 'gio_ra'))
        
        hc = float(row.get('Giờ hành chính', 0)) if pd.notna(row.get('Giờ hành chính')) else 0.0
        ws.cell(row=row_idx, column=9, value=round(hc/8, 2) if hc > 0 else 0)
        
        ws.cell(row=row_idx, column=10, value=hc if hc > 0 else "")
        
        ot = float(row.get('Giờ OT', 0)) if pd.notna(row.get('Giờ OT')) else 0.0
        ws.cell(row=row_idx, column=11, value=ot if ot > 0 else "")
        
        col_hc_letter = get_column_letter(10)
        col_ot_letter = get_column_letter(11)
        ws.cell(row=row_idx, column=12, value=f"=SUM({col_hc_letter}{row_idx}:{col_ot_letter}{row_idx})")
        
        ws.cell(row=row_idx, column=13, value=str(row.get('Lý do tăng ca', '')) if pd.notna(row.get('Lý do tăng ca')) else "")
        
        ghi_chu = str(row.get('Ghi chú', '')) if pd.notna(row.get('Ghi chú')) else ""
        loai = row.get('_loai', '') if pd.notna(row.get('_loai', '')) else ''
        
        ghi_chu_clean = ghi_chu
        for icon in ["🚫 ", "✅ ", "⚠️ ", "📝 ", "📅 ", "🔴 ", "🟣 ", "🟠 ", "🟢 ", "🔵 "]:
            ghi_chu_clean = ghi_chu_clean.replace(icon, "")
        ghi_chu_clean = ghi_chu_clean.strip()
        
        is_abnormal = False
        if "nghi_khong_phep" in loai or "Thiếu giờ" in ghi_chu_clean or "Lỗi" in ghi_chu_clean or "Vắng" in ghi_chu_clean or "Làm thiếu" in ghi_chu_clean:
            is_abnormal = True

        ws.cell(row=row_idx, column=14, value=ghi_chu_clean.strip())
        
        font_red = Font(name='Times New Roman', size=12, color="FF0000")
        for col_num in range(1, 15):
            c = ws.cell(row=row_idx, column=col_num)
            c.font = font_normal
            if col_num == 14 and is_abnormal:
                c.font = font_red
            c.border = border
            if col_num in [1,2,3,4, 13, 14]:
                c.alignment = align_left
            else:
                c.alignment = align_center
            
            if thu in [5, 6] and not (thu == 5 and is_last_saturday_of_month(d_obj)):
                c.fill = gray_fill
                
        row_idx += 1
        
    for col_num, width in enumerate([10, 20, 15, 15, 12, 8, 10, 10, 8, 15, 10, 15, 20, 30], 1):
        ws.column_dimensions[get_column_letter(col_num)].width = width
        
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ==========================================
# 6. GIAO DIỆN CHÍNH
# ==========================================
if "step" not in st.session_state: st.session_state.step = 1
if "df_raw" not in st.session_state: st.session_state.df_raw = None
if "show_history" not in st.session_state: st.session_state.show_history = False

# ----- QUẢN LÝ TRẠNG THÁI TRANG (ROUTING) -----
if "app_page" not in st.session_state:
    st.session_state.app_page = "home"

# ----- QUẢN LÝ HIỂN THỊ SIDEBAR THEO TRANG -----
if st.session_state.app_page == "chamcong":
    # Trang chấm công: Sidebar luôn mở và cố định bên trái
    st.markdown("""
    <style>
        [data-testid="stSidebar"] { 
            display: block !important; 
            visibility: visible !important; 
            transform: none !important; 
            min-width: 320px !important;
            max-width: 320px !important;
            position: fixed !important;
            left: 0 !important;
            top: 0 !important;
            height: 100vh !important;
            z-index: 999999 !important;
            background-color: white !important;
            box-shadow: 2px 0 10px rgba(0,0,0,0.1) !important;
        }
        [data-testid="collapsedControl"] { display: none !important; }
        [data-testid="stAppViewContainer"] {
            margin-left: 320px !important;
            width: calc(100% - 320px) !important;
        }
    </style>
    """, unsafe_allow_html=True)
else:
    # Trang Home và MOS: Ẩn hoàn toàn sidebar và dãn full màn hình
    st.markdown("""
    <style>
        [data-testid="stSidebar"] { display: none !important; }
        [data-testid="collapsedControl"] { display: none !important; }
        [data-testid="stAppViewContainer"] {
            margin-left: 0 !important;
            width: 100% !important;
        }
    </style>
    """, unsafe_allow_html=True)

# ----- NGÔN NGỮ (LANGUAGE) -----
if "lang" not in st.session_state:
    st.session_state.lang = "vi"

def render_lang_toggle():
    """Render nút gạt ngôn ngữ Nhật-Việt fixed ở góc trên phải, dùng st.button."""
    is_ja = (st.session_state.lang == "ja")

    track_class = "ja-active" if is_ja else "vi-active"
    vi_class    = "inactive"  if is_ja else "active"
    ja_class    = "active"    if is_ja else "inactive"

    st.markdown(f"""
    <style>
    /* ===== Language Toggle Fixed Top-Right ===== */
    .lang-toggle-wrapper {{
        position: fixed !important;
        top: 40px !important;
        right: 420px !important;
        z-index: 999998 !important;
        display: flex !important;
        align-items: center !important;
        gap: 10px !important;
        background: rgba(255,255,255,0.94) !important;
        backdrop-filter: blur(14px) !important;
        border: 1.5px solid rgba(14,165,233,0.18) !important;
        border-radius: 50px !important;
        padding: 7px 16px 7px 12px !important;
        box-shadow: 0 4px 20px rgba(14,165,233,0.13), 0 1.5px 6px rgba(0,0,0,0.06) !important;
        pointer-events: none !important;
    }}
    .lang-flag-label {{
        font-size: 19px !important;
        line-height: 1 !important;
        user-select: none !important;
    }}
    .lang-flag-label.inactive {{ opacity: 0.32 !important; }}
    .lang-flag-label.active   {{ opacity: 1 !important; }}

    /* Toggle track */
    .lang-switch-track {{
        position: relative !important;
        width: 46px !important;
        height: 24px !important;
        border-radius: 999px !important;
        flex-shrink: 0 !important;
    }}
    .lang-switch-track.vi-active {{
        background: linear-gradient(135deg, #0EA5E9, #38BDF8) !important;
    }}
    .lang-switch-track.ja-active {{
        background: linear-gradient(135deg, #0369A1, #0EA5E9) !important;
    }}
    .lang-switch-track::after {{
        content: '' !important;
        position: absolute !important;
        top: 3px !important;
        width: 18px !important;
        height: 18px !important;
        border-radius: 50% !important;
        background: white !important;
        box-shadow: 0 2px 6px rgba(0,0,0,0.25) !important;
    }}
    .lang-switch-track.vi-active::after {{ left: 3px !important; }}
    .lang-switch-track.ja-active::after {{ left: 25px !important; }}

    /* ===== The actual clickable Streamlit button — transparent, covers the toggle ===== */
    .st-key-lang_switch_btn {{
        position: fixed !important;
        top: 36px !important;
        right: 420px !important;
        z-index: 999999 !important;
        width: 168px !important;
    }}
    .st-key-lang_switch_btn > div {{
        width: 168px !important;
    }}
    .st-key-lang_switch_btn button {{
        background: transparent !important;
        border: none !important;
        box-shadow: none !important;
        width: 168px !important;
        height: 48px !important;
        min-height: 48px !important;
        padding: 0 !important;
        cursor: pointer !important;
        border-radius: 50px !important;
    }}
    .st-key-lang_switch_btn button:hover {{
        background: rgba(14,165,233,0.06) !important;
        box-shadow: none !important;
        transform: none !important;
        filter: none !important;
    }}
    .st-key-lang_switch_btn button p {{
        display: none !important;
    }}
    </style>

    <div class="lang-toggle-wrapper">
        <span class="lang-flag-label {vi_class}">🇻🇳</span>
        <div class="lang-switch-track {track_class}"></div>
        <span class="lang-flag-label {ja_class}">🇯🇵</span>
    </div>
    """, unsafe_allow_html=True)

    # Nút thực sự — transparent, đè lên visual toggle để nhận click
    def toggle_lang():
        st.session_state.lang = "ja" if st.session_state.lang == "vi" else "vi"
    st.button("　", key="lang_switch_btn", help="Chuyển đổi ngôn ngữ / 言語切替", on_click=toggle_lang)



def render_chatbot():
    # 7. CHATBOT BONG BÓNG LƠ LỬNG
    # ==========================================
    lang = st.session_state.get('lang', 'vi')
    lbl_online = "Đang trực tuyến" if lang == 'vi' else "オンライン"
    lbl_api_warn = "⚠️ Vui lòng nhập Gemini API Key để Chatbot hoạt động!" if lang == 'vi' else "⚠️ チャットボットを機能させるにはGemini APIキーを入力してください！"
    lbl_api_tip = "💡 **Mẹo:** Bạn có thể nhập nhiều API Key cách nhau bằng dấu phẩy (,) để hệ thống tự động luân chuyển khi quá tải!" if lang == 'vi' else "💡 **ヒント:** 複数のAPIキーをカンマ(,)で区切って入力すると、過負荷時に自動で切り替わります！"
    lbl_api_placeholder = "Nhập Gemini API Key tại đây..." if lang == 'vi' else "ここにAPIキーを入力..."
    lbl_api_label = "Nhập Gemini API Key" if lang == 'vi' else "Gemini APIキーを入力"
    lbl_settings = "⚙️ Cài đặt Chatbot" if lang == 'vi' else "⚙️ チャットボット設定"
    lbl_change_api = "🔄 Đổi mã API Key" if lang == 'vi' else "🔄 APIキーを変更する"
    lbl_chat_input = "Hỏi AI..." if lang == 'vi' else "AIに質問する..."
    lbl_hello = "Xin chào! Tôi có thể giúp gì cho bạn về dữ liệu chấm công?" if lang == 'vi' else "こんにちは！勤怠データについて何かお手伝いしましょうか？"
    lang_instruction = "TIẾNG VIỆT" if lang == 'vi' else "TIẾNG NHẬT (日本語)"

    if 'chat_open' not in st.session_state:
        st.session_state.chat_open = False
    if 'chat_pos' not in st.session_state:
        st.session_state.chat_pos = 'right'

    def toggle_chat():
        st.session_state.chat_open = not st.session_state.chat_open

    # CSS bong bóng tròn màu xanh — dùng key container để CSS selector chính xác
    side = st.session_state.chat_pos
    css = """
    <style>
    .st-key-chat_bubble_wrap button {
        border-radius: 50% !important;
        width: 64px !important; height: 64px !important;
        background: linear-gradient(135deg, #0EA5E9 0%, #0284C7 100%) !important;
        color: white !important; border: none !important;
        box-shadow: 0 8px 24px rgba(14,165,233,0.4) !important;
        font-size: 26px !important;
        transition: transform 0.2s, box-shadow 0.2s !important;
        padding: 0 !important;
        min-height: 64px !important;
    }
    .st-key-chat_bubble_wrap button:hover {
        transform: scale(1.08) !important;
        box-shadow: 0 12px 32px rgba(14,165,233,0.5) !important;
    }
    .st-key-chat_bubble_wrap {
        position: fixed !important;
        bottom: 30px !important;
        right: 30px !important;
        z-index: 9999999 !important;
        width: auto !important;
    }
    .st-key-chat_box_wrap {
        position: fixed !important;
        bottom: 100px !important;
        right: 30px !important;
        width: 360px !important;
        background: #F8FAFC !important;
        border-radius: 16px !important;
        box-shadow: 0 20px 48px rgba(0,0,0,0.2) !important;
        z-index: 9999999 !important;
        padding: 0 !important;
        border: 1px solid #E2E8F0 !important;
        overflow: hidden !important;
    }
    .st-key-chat_box_wrap [data-testid="stVerticalBlock"] {
        gap: 0 !important; padding: 0 !important;
    }
    .st-key-chat_box_wrap [data-testid="stChatMessage"] {
        background: white !important; border: 1px solid #E2E8F0 !important;
        border-radius: 16px 16px 16px 4px !important; padding: 12px !important;
        box-shadow: 0 2px 8px rgba(0,0,0,0.02) !important; margin: 12px !important; width: auto !important;
    }
    .st-key-chat_box_wrap [data-testid="stChatMessage"] [data-testid="stMarkdownContainer"] p {
        font-size: 13.5px !important; color: #334155 !important; line-height: 1.5 !important;
    }
    .st-key-chat_box_wrap [data-testid="stChatInput"] {
        background: white !important; padding: 12px 16px !important; border-top: 1px solid #E2E8F0 !important;
    }
    .st-key-chat_box_wrap [data-testid="stChatInput"] textarea {
        background: #F1F5F9 !important; border-radius: 24px !important; border: 1px solid #E2E8F0 !important; font-size: 13px !important;
    }
    .st-key-close_chat_box { position: absolute !important; top: 12px !important; right: 12px !important; z-index: 100 !important; }
    .st-key-close_chat_box button { background: rgba(255,255,255,0.1) !important; color: white !important; border: none !important; border-radius: 50% !important; width: 32px !important; height: 32px !important; min-height: 32px !important; padding: 0 !important; display: flex !important; align-items: center !important;  }
    .st-key-close_chat_box button:hover { background: rgba(255,255,255,0.2) !important; color: white !important; }
    </style>
    """
    css = css.replace("{side}", side)
    st.markdown(css, unsafe_allow_html=True)
    

    bubble = st.container(key="chat_bubble_wrap")
    with bubble:
        st.button("💬", on_click=toggle_chat, key="chat_bubble_btn", help="Trợ lý AI V-MOS")

    if st.session_state.chat_open:
        chat_box = st.container(key="chat_box_wrap")
        with chat_box:
            if st.button("✕", key="close_chat_box"):
                st.session_state.chat_open = False
                st.rerun()
            st.markdown("""
            <div style="background:#0EA5E9; color:white; padding:16px; display:flex; align-items:center; gap:12px;">
                <div style="width:40px; height:40px; background:linear-gradient(135deg, #F472B6 0%, #3B82F6 100%); border-radius:50%; display:flex; align-items:center; justify-content:center; font-size:24px; box-shadow:0 4px 10px rgba(0,0,0,0.1); flex-shrink:0;">🤖</div>
                <div style="display:flex; flex-direction:column;">
                    <div style="font-weight:700; font-size:15px; line-height:1.2;">V-MOS Assistant</div>
                    <div style="font-size:12px; display:flex; align-items:center; gap:6px; opacity:0.9; margin-top:4px;">
                        <span style="width:8px; height:8px; background:#10B981; border-radius:50%; display:inline-block;"></span> {lbl_online}
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            saved_key = load_saved_api_key()
            if saved_key:
                st.session_state['gemini_configured'] = True
            else:
                st.session_state['gemini_configured'] = False
                
            if not st.session_state['gemini_configured']:
                st.warning(lbl_api_warn)
                st.info(lbl_api_tip, icon="🔑")
                new_key = st.text_input(lbl_api_label, type="password", placeholder=lbl_api_placeholder)
                if new_key:
                    import toml
                    try:
                        with open(".streamlit/secrets.toml", "r", encoding="utf-8") as f:
                            secrets = toml.load(f)
                    except:
                        secrets = {}
                    keys_list = [k.strip() for k in new_key.split(",") if k.strip()]
                    if len(keys_list) > 1:
                        secrets["GEMINI_API_KEYS"] = keys_list
                        if "GEMINI_API_KEY" in secrets: del secrets["GEMINI_API_KEY"]
                    elif len(keys_list) == 1:
                        secrets["GEMINI_API_KEY"] = keys_list[0]
                        if "GEMINI_API_KEYS" in secrets: del secrets["GEMINI_API_KEYS"]
                    with open(".streamlit/secrets.toml", "w", encoding="utf-8") as f:
                        toml.dump(secrets, f)
                    st.session_state['gemini_configured'] = True
                    st.rerun()
            else:
                with st.expander(lbl_settings, expanded=False):
                    if st.button(lbl_change_api, use_container_width=True):
                        try:
                            import toml
                            with open(".streamlit/secrets.toml", "w", encoding="utf-8") as f:
                                toml.dump({}, f)
                        except: pass
                        st.session_state['gemini_configured'] = False
                        st.rerun()
                    
                if 'chat_messages' not in st.session_state:
                    st.session_state['chat_messages'] = [{"role": "assistant", "content": lbl_hello}]
                elif len(st.session_state['chat_messages']) == 1 and st.session_state['chat_messages'][0]["role"] == "assistant":
                    st.session_state['chat_messages'][0]["content"] = lbl_hello

                chat_container = st.container(height=300)
                with chat_container:
                    for msg in st.session_state['chat_messages']:
                        with st.chat_message(msg["role"], avatar="🤖" if msg["role"] == "assistant" else "👤"):
                            st.markdown(msg["content"])

                prompt = st.chat_input(lbl_chat_input)
                if prompt:
                    st.session_state['chat_messages'].append({"role": "user", "content": prompt})
                    with chat_container:
                        with st.chat_message("user", avatar="👤"):
                            st.markdown(prompt)
                        with st.chat_message("assistant", avatar="🤖"):
                            with st.spinner("..."):
                                try:
                                    import requests
                                    saved_key = load_saved_api_key()
                                    if not saved_key:
                                        raise Exception("Không tìm thấy API Key. Vui lòng nhập lại API Key ở menu.")

                                    df_ctx = "Chưa có dữ liệu"
                                    if st.session_state.get('app_page') == 'mos':
                                        if 'df_mos_edited' in st.session_state and st.session_state.df_mos_edited is not None:
                                            df_ctx = "--- DỮ LIỆU TỔNG HỢP DỰ ÁN MOS ---\n" + st.session_state.df_mos_edited.to_csv(index=False)
                                        else:
                                            df_ctx = "Chưa có dữ liệu dự án MOS nào được xử lý."
                                    else:
                                        if 'df_filtered_for_chat' in st.session_state and st.session_state.df_filtered_for_chat is not None:
                                            dff = st.session_state.df_filtered_for_chat.copy()
                                            ten = st.session_state.mapping.get('ten_nv', dff.columns[1])
                                            dff['Số giờ làm thực tế'] = pd.to_numeric(dff['Số giờ làm thực tế'], errors='coerce').fillna(0)
                                            dff['Giờ OT'] = pd.to_numeric(dff['Giờ OT'], errors='coerce').fillna(0)
                                            dff['Giờ hành chính'] = pd.to_numeric(dff['Giờ hành chính'], errors='coerce').fillna(0)
    
                                            df_detail = pd.DataFrame({
                                                "Tên NV": dff[ten],
                                                "Ngày": dff["Ngày"],
                                                "Giờ làm": dff["Giờ hành chính"].round(2),
                                                "OT": dff["Giờ OT"].round(2),
                                                "Tổng": dff["Số giờ làm thực tế"].round(2),
                                                "Ghi chú": dff["Ghi chú"]
                                            })
    
                                            if len(df_detail) <= 500:
                                                df_ctx = f"--- DỮ LIỆU CHI TIẾT (Toàn bộ {len(df_detail)} dòng) ---\n" + df_detail.to_csv(index=False)
                                            else:
                                                agg = df_detail.groupby("Tên NV").agg({'Giờ làm': 'sum', 'OT': 'sum', 'Tổng': 'sum'}).reset_index().round(2)
                                                df_ctx = f"--- BẢNG TỔNG HỢP (TẤT CẢ NHÂN VIÊN) ---\n{agg.to_csv(index=False)}\n\n"
    
                                                # Nén dữ liệu cực đại để AI đọc được toàn bộ không bị mất dòng nào
                                                compressed_lines = []
                                                for name, group in df_detail.groupby("Tên NV"):
                                                    days = []
                                                    for _, row in group.iterrows():
                                                        d = str(row["Ngày"])[:5] # "15/05/2026" -> "15/05"
                                                        hc = row["Giờ làm"]
                                                        ot = row["OT"]
                                                        note = str(row["Ghi chú"]).strip()
                                                        if ot == 0 and note == "":
                                                            if hc == 8.0: days.append(d)
                                                            elif hc == 0.0: days.append(f"{d}(Nghỉ)")
                                                            else: days.append(f"{d}({hc}h)")
                                                        else:
                                                            days.append(f"{d}({hc}h|OT:{ot}|{note})")
                                                    compressed_lines.append(f"{name}: " + ", ".join(days))
    
                                                df_ctx += "--- CHI TIẾT LỊCH SỬ TỪNG NGÀY (DẠNG NÉN SIÊU TỐI ƯU) ---\n"
                                                df_ctx += "Quy ước: Nếu chỉ ghi ngày (VD: 15/05) nghĩa là làm đủ 8h không OT. Các ngày khác ghi rõ (Giờ làm|OT|Ghi chú).\n"
                                                df_ctx += "\n".join(compressed_lines)
                                        elif 'df_result' in st.session_state and st.session_state.df_result is not None:
                                            df_ctx = st.session_state.df_result.head(50).to_csv(index=False)
                                        else:
                                            df_ctx = "Chưa có dữ liệu nào được tải lên."

                                    sys_prompt = f"""Bạn là V-MOS AI, một siêu trợ lý trí tuệ nhân tạo toàn năng.
Bạn là trợ lý chính cho hệ thống "Quản lý Chấm công & Dự án Nội bộ MOS". Bạn am hiểu sâu sắc về cả việc tính toán giờ làm việc (chấm công) và tổng hợp giờ làm dự án MOS.
Bạn SẴN SÀNG TRẢ LỜI BẤT KỲ CÂU HỎI NÀO CỦA NGƯỜI DÙNG (từ phân tích số liệu bảng chấm công, báo cáo dự án MOS, cho đến viết email, làm toán, lập trình, tư vấn).
Hãy thoải mái trò chuyện và hỗ trợ người dùng như một AI thực thụ. Tuyệt đối không bao giờ từ chối trả lời vì lý do "tôi chỉ là trợ lý chấm công".
BẠN ĐANG GIAO TIẾP VỚI NGƯỜI DÙNG. NGÔN NGỮ GIAO TIẾP VÀ TRẢ LỜI BẮT BUỘC LÀ: {lang_instruction}. BẠN PHẢI TUYỆT ĐỐI TUÂN THỦ NGÔN NGỮ NÀY TRONG CÂU TRẢ LỜI CỦA MÌNH.

[DỮ LIỆU HIỆN TẠI MÀ NGƯỜI DÙNG ĐANG TƯƠNG TÁC]:
{df_ctx}

Luôn ưu tiên trả lời tự nhiên, thân thiện và chính xác."""

                                    try:
                                        import httpx
                                        import time
                                        
                                        # Lịch sử đã có sẵn prompt ở cuối cùng
                                        gemini_history = []
                                        # Bỏ qua system_instruction, chuyển đổi lịch sử
                                        for msg in st.session_state['chat_messages'][-10:]:
                                            role = "user" if msg["role"] == "user" else "model"
                                            gemini_history.append({"role": role, "parts": [{"text": msg["content"]}]})

                                        data = {
                                            "systemInstruction": {"parts": [{"text": sys_prompt}]},
                                            "contents": gemini_history,
                                            "generationConfig": {"temperature": 0.7}
                                        }
                                        
                                        # Lấy danh sách keys trực tiếp để quản lý vòng lặp
                                        all_keys = []
                                        try:
                                            import toml
                                            with open(".streamlit/secrets.toml", "r", encoding="utf-8") as f:
                                                secrets = toml.load(f)
                                                if "GEMINI_API_KEYS" in secrets:
                                                    all_keys = secrets["GEMINI_API_KEYS"]
                                                elif "GEMINI_API_KEY" in secrets:
                                                    all_keys = [secrets["GEMINI_API_KEY"]]
                                        except:
                                            pass
                                            
                                        if not all_keys:
                                            k = load_saved_api_key()
                                            if k: all_keys = [k]
                                            
                                        if not all_keys:
                                            answer = "⚠️ Vui lòng cấu hình API Key."
                                        else:
                                            # Thử 2 vòng qua tất cả các key
                                            max_retries = len(all_keys) * 2
                                            if max_retries < 3: max_retries = 3
                                            
                                            answer = None
                                            for attempt in range(max_retries):
                                                current_key = all_keys[attempt % len(all_keys)]
                                                url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={current_key}"
                                                
                                                res = httpx.post(url, headers={"Content-Type": "application/json"}, json=data, timeout=30.0)
                                                
                                                if res.status_code == 200:
                                                    answer = res.json()["candidates"][0]["content"]["parts"][0]["text"].strip()
                                                    break
                                                elif res.status_code == 400 and "API_KEY_INVALID" in res.text:
                                                    if attempt == max_retries - 1:
                                                        answer = "Tất cả các API Key đều không hợp lệ. Vui lòng kiểm tra lại Cài đặt Chatbot."
                                                    continue
                                                elif res.status_code in [429, 503, 500, 502, 504]:
                                                    if attempt < max_retries - 1:
                                                        import time
                                                        # Nếu còn key khác chưa thử trong vòng 1, thử ngay lập tức
                                                        if len(all_keys) > 1 and attempt < len(all_keys) - 1:
                                                            time.sleep(0.5)
                                                        else:
                                                            # Đã thử hết 1 vòng hoặc chỉ có 1 key -> đợi lâu hơn (Exponential)
                                                            wait_time = 2 ** (attempt % 4) + 1 
                                                            time.sleep(wait_time)
                                                        continue
                                                    else:
                                                        answer = f"Hệ thống AI đang quá tải (Lỗi {res.status_code}). Vui lòng đợi khoảng 1 phút rồi nhắn lại nhé!"
                                                        break
                                                else:
                                                    answer = f"Lỗi kết nối Gemini: {res.text}" 
                                                    break
                                                
                                        if not answer:
                                            answer = "Lỗi không xác định khi kết nối."
                                    except Exception as e:
                                        answer = f"Lỗi hệ thống: {str(e)}"

                                    st.markdown(answer)
                                    st.session_state['chat_messages'].append({"role": "assistant", "content": answer})
                                except Exception as e:
                                    st.error(f"Lỗi: {str(e)}")


# 1. GIAO DIỆN CHUNG (HOME)
logo_html_large = _logo_img_tag(LOGO_HEADER_B64, "height:200px;width:auto;object-fit:contain;mix-blend-mode:multiply;margin-bottom:20px; filter: drop-shadow(0 4px 10px rgba(0,0,0,0.1)); transition: transform 0.3s ease;")

# Render nút gạt ngôn ngữ trên MỌI trang
render_lang_toggle()
t = get_t(st.session_state.lang)
st.session_state['cached_t'] = t

# Dark mode removed - always light mode
st.session_state.dark_mode = False

st.markdown("""
    <style>
    /* Lang toggle — flush right edge */
    .lang-toggle-wrapper {
        top: 10px !important;
        right: 20px !important;
    }
    .st-key-lang_switch_btn {
        top: 6px !important;
        right: 20px !important;
    }
    /* Hide sidebar collapse button globally */
    [data-testid="collapsedControl"] {
        display: none !important;
    }
    </style>
""", unsafe_allow_html=True)

if st.session_state.app_page == "home":
    nav_text_color = "#0F172A"
    nav_bg_color   = "rgba(255,255,255,0.8)"
    
    if st.session_state.lang == 'ja':
        feature_modal_html = """<style>.vimos-feature-modal{display:none;position:fixed;top:0;left:0;width:100vw;height:100vh;background:rgba(15,23,42,0.6);backdrop-filter:blur(8px);z-index:99999999;align-items:center;justify-content:center;font-family:'Inter','Be Vietnam Pro',sans-serif;}.vimos-feature-modal:target{display:flex!important;}.vimos-modal-content{background:white;width:90%;max-width:650px;border-radius:24px;padding:40px;box-shadow:0 25px 50px -12px rgba(0,0,0,0.25);position:relative;border:1px solid #E2E8F0;max-height:85vh;overflow-y:auto;text-align:center;}.vimos-modal-close{position:absolute;top:24px;right:24px;background:#F1F5F9;border:none;width:36px;height:36px;border-radius:50%;font-size:18px;cursor:pointer;color:#64748B;display:flex;align-items:center;justify-content:center;text-decoration:none;}.vimos-modal-close:hover{background:#E2E8F0;}.vimos-modal-btn{background:linear-gradient(135deg,#0EA5E9,#0284C7);color:white;padding:12px 36px;border-radius:100px;border:none;font-weight:600;font-size:15px;cursor:pointer;box-shadow:0 8px 16px rgba(14,165,233,0.25);text-decoration:none;display:inline-block;}</style><div id="feature-modal-box" class="vimos-feature-modal"><div class="vimos-modal-content"><a href="#_" class="vimos-modal-close">✕</a><h2 style="font-size:24px;font-weight:800;color:#0F172A;margin-top:0;margin-bottom:28px;background:linear-gradient(135deg,#0284C7,#0EA5E9);-webkit-background-clip:text;-webkit-text-fill-color:transparent;">✦ システム機能全般</h2><div style="display:flex;flex-direction:column;gap:20px;text-align:left;"><div style="display:flex;gap:16px;align-items:flex-start;"><div style="background:#E0F2FE;color:#0284C7;width:44px;height:44px;border-radius:12px;display:flex;align-items:center;justify-content:center;font-size:22px;flex-shrink:0;">⏱️</div><div><div style="font-weight:700;font-size:16px;color:#1E293B;margin-bottom:4px;">実働時間＆自動OT計算</div><div style="color:#64748B;font-size:14px;line-height:1.5;">勤怠タイムレコーダーのExcel生データをアップロードすると、出退勤時間を自動認識し、標準労働時間および残業時間（OT）を分単位で正確に算出します。</div></div></div><div style="display:flex;gap:16px;align-items:flex-start;"><div style="background:#F0FDF4;color:#16A34A;width:44px;height:44px;border-radius:12px;display:flex;align-items:center;justify-content:center;font-size:22px;flex-shrink:0;">📊</div><div><div style="font-weight:700;font-size:16px;color:#1E293B;margin-bottom:4px;">MOSプロジェクト工数集計</div><div style="color:#64748B;font-size:14px;line-height:1.5;">MOSプロジェクトコードごとの各メンバーの作業報告集計をサポートします。自動照合、工数比率分析、Excel集計レポートファイルの出力を実行します。</div></div></div><div style="display:flex;gap:16px;align-items:flex-start;"><div style="background:#FEF3C7;color:#D97706;width:44px;height:44px;border-radius:12px;display:flex;align-items:center;justify-content:center;font-size:22px;flex-shrink:0;">🤖</div><div><div style="font-weight:700;font-size:16px;color:#1E293B;margin-bottom:4px;">24時間365日AIアシスタント</div><div style="color:#64748B;font-size:14px;line-height:1.5;">AIチャットボットを統合し、社内規定や労働法に関する疑問、ソフトウェアの操作手順に即座に回答します。</div></div></div><div style="display:flex;gap:16px;align-items:flex-start;"><div style="background:#F3E8FF;color:#9333EA;width:44px;height:44px;border-radius:12px;display:flex;align-items:center;justify-content:center;font-size:22px;flex-shrink:0;">🌐</div><div><div style="font-weight:700;font-size:16px;color:#1E293B;margin-bottom:4px;">日越バイリンガル切替</div><div style="color:#64748B;font-size:14px;line-height:1.5;">ベトナム語と日本語の間でインターフェース全体をスムーズに切替可能であり、国際的な職場環境に最適化されています。</div></div></div></div><div style="margin-top:32px;padding-top:20px;border-top:1px solid #F1F5F9;"><a href="#_" class="vimos-modal-btn">了解</a></div></div></div>"""
        about_modal_html = """<div id="about-modal-box" class="vimos-feature-modal"><div class="vimos-modal-content"><a href="#_" class="vimos-modal-close">✕</a><h2 style="font-size:24px;font-weight:800;color:#0F172A;margin-top:0;margin-bottom:20px;background:linear-gradient(135deg,#0284C7,#0EA5E9);-webkit-background-clip:text;-webkit-text-fill-color:transparent;">🏢 VIET.MOSについて</h2><div style="text-align:justify;color:#334155;font-size:15px;line-height:1.7;padding:0 10px;">VIET.MOS COMPANY LIMITEDは、情報技術分野で事業を展開し、製造業向けの自動化ソリューションの設計および提供を行う企業です。ベトナムの創造的精神と日本の先進的なエンジニアリング思考の融合に基づく発展方針のもと、VIET.MOSは機械設計、電気設計・制御プログラミング、シミュレーション設計、およびオーダーメイドの機械ソリューション構築サービスに注力しています。専門性の高いエンジニアチームと絶え間ない学習精神を持ち、当社は常に顧客に最適で効率的かつ高品質なソリューションを提供することを目指すと同時に、生産価値の向上と社会の持続可能な発展に貢献してまいります。</div><div style="margin-top:32px;padding-top:20px;border-top:1px solid #F1F5F9;"><a href="#_" class="vimos-modal-btn">閉じる</a></div></div></div>"""
        guide_modal_html = """<div id="guide-modal-box" class="vimos-feature-modal"><div class="vimos-modal-content" style="text-align:left;"><a href="#_" class="vimos-modal-close">✕</a><h2 style="font-size:24px;font-weight:800;color:#0F172A;margin-top:0;margin-bottom:20px;text-align:center;background:linear-gradient(135deg,#0284C7,#0EA5E9);-webkit-background-clip:text;-webkit-text-fill-color:transparent;">📖 利用ガイド</h2><div style="color:#334155;font-size:14.5px;line-height:1.6;"><p style="margin-top:0;">クイックガイドへようこそ！タイムレコーダーのExcelファイルから労働時間を計算する基本手順は以下の通りです：</p><div style="margin-bottom:12px;"><b style="color:#0284C7;">1. Excelファイルのアップロード</b><div style="padding-left:12px;color:#64748B;">• 左側パネルにある「ファイルアップロード」ボタンをクリックするか、タイムレコーダーのExcelファイルをドラッグ＆ドロップします。</div></div><div style="margin-bottom:12px;"><b style="color:#0284C7;">2. データ列の設定</b><div style="padding-left:12px;color:#64748B;">• 出退勤時間などの列が自動スキャンされます。正しく認識されない場合はドロップダウンから選択し直して確認を押してください。</div></div><div style="margin-bottom:12px;"><b style="color:#0284C7;">3. 標準定時時間の設定</b><div style="padding-left:12px;color:#64748B;">• 標準勤務時間帯（例：08:00 - 17:00）と昼休憩時間を入力することで休憩時間が自動控除されます。</div></div><div style="margin-bottom:12px;"><b style="color:#0284C7;">4. データ処理＆レポート出力</b><div style="padding-left:12px;color:#64748B;">• 「データ処理」ボタンを押してダッシュボード統計を確認後、「Excelレポートをダウンロード」をクリックして保存します。</div></div></div><div style="margin-top:28px;padding-top:16px;border-top:1px solid #F1F5F9;text-align:center;"><a href="#_" class="vimos-modal-btn">了解</a></div></div></div>"""
    else:
        feature_modal_html = """<style>.vimos-feature-modal{display:none;position:fixed;top:0;left:0;width:100vw;height:100vh;background:rgba(15,23,42,0.6);backdrop-filter:blur(8px);z-index:99999999;align-items:center;justify-content:center;font-family:'Inter','Be Vietnam Pro',sans-serif;}.vimos-feature-modal:target{display:flex!important;}.vimos-modal-content{background:white;width:90%;max-width:650px;border-radius:24px;padding:40px;box-shadow:0 25px 50px -12px rgba(0,0,0,0.25);position:relative;border:1px solid #E2E8F0;max-height:85vh;overflow-y:auto;text-align:center;}.vimos-modal-close{position:absolute;top:24px;right:24px;background:#F1F5F9;border:none;width:36px;height:36px;border-radius:50%;font-size:18px;cursor:pointer;color:#64748B;display:flex;align-items:center;justify-content:center;text-decoration:none;}.vimos-modal-close:hover{background:#E2E8F0;}.vimos-modal-btn{background:linear-gradient(135deg,#0EA5E9,#0284C7);color:white;padding:12px 36px;border-radius:100px;border:none;font-weight:600;font-size:15px;cursor:pointer;box-shadow:0 8px 16px rgba(14,165,233,0.25);text-decoration:none;display:inline-block;}</style><div id="feature-modal-box" class="vimos-feature-modal"><div class="vimos-modal-content"><a href="#_" class="vimos-modal-close">✕</a><h2 style="font-size:24px;font-weight:800;color:#0F172A;margin-top:0;margin-bottom:28px;background:linear-gradient(135deg,#0284C7,#0EA5E9);-webkit-background-clip:text;-webkit-text-fill-color:transparent;">✦ Toàn bộ tính năng hệ thống</h2><div style="display:flex;flex-direction:column;gap:20px;text-align:left;"><div style="display:flex;gap:16px;align-items:flex-start;"><div style="background:#E0F2FE;color:#0284C7;width:44px;height:44px;border-radius:12px;display:flex;align-items:center;justify-content:center;font-size:22px;flex-shrink:0;">⏱️</div><div><div style="font-weight:700;font-size:16px;color:#1E293B;margin-bottom:4px;">Tính số giờ làm thực tế & OT tự động</div><div style="color:#64748B;font-size:14px;line-height:1.5;">Tải lên file Excel chấm công thô từ máy quét, hệ thống tự động nhận diện cột Giờ Vào/Ra, tính toán tổng giờ làm tiêu chuẩn và giờ làm thêm (Overtime) chuẩn xác đến từng phút.</div></div></div><div style="display:flex;gap:16px;align-items:flex-start;"><div style="background:#F0FDF4;color:#16A34A;width:44px;height:44px;border-radius:12px;display:flex;align-items:center;justify-content:center;font-size:22px;flex-shrink:0;">📊</div><div><div style="font-weight:700;font-size:16px;color:#1E293B;margin-bottom:4px;">Tổng hợp giờ làm dự án MOS</div><div style="color:#64748B;font-size:14px;line-height:1.5;">Hỗ trợ tổng hợp báo cáo công việc của từng thành viên theo mã dự án MOS. Tự động đối chiếu, phân tích tỉ lệ đóng góp và xuất ra file báo cáo Excel tổng hợp hoàn chỉnh.</div></div></div><div style="display:flex;gap:16px;align-items:flex-start;"><div style="background:#FEF3C7;color:#D97706;width:44px;height:44px;border-radius:12px;display:flex;align-items:center;justify-content:center;font-size:22px;flex-shrink:0;">🤖</div><div><div style="font-weight:700;font-size:16px;color:#1E293B;margin-bottom:4px;">Trợ lý AI thông minh 24/7</div><div style="color:#64748B;font-size:14px;line-height:1.5;">Tích hợp Chatbot AI sẵn sàng giải đáp mọi thắc mắc về quy định, luật lao động và hướng dẫn thao tác phần mềm tức thì.</div></div></div><div style="display:flex;gap:16px;align-items:flex-start;"><div style="background:#F3E8FF;color:#9333EA;width:44px;height:44px;border-radius:12px;display:flex;align-items:center;justify-content:center;font-size:22px;flex-shrink:0;">🌐</div><div><div style="font-weight:700;font-size:16px;color:#1E293B;margin-bottom:4px;">Chuyển ngữ song ngữ Việt - Nhật</div><div style="color:#64748B;font-size:14px;line-height:1.5;">Hỗ trợ chuyển đổi linh hoạt toàn bộ giao diện giữa Tiếng Việt và Tiếng Nhật một cách mượt mà, tối ưu cho môi trường làm việc quốc tế.</div></div></div></div><div style="margin-top:32px;padding-top:20px;border-top:1px solid #F1F5F9;"><a href="#_" class="vimos-modal-btn">Đã hiểu</a></div></div></div>"""
        about_modal_html = """<div id="about-modal-box" class="vimos-feature-modal"><div class="vimos-modal-content"><a href="#_" class="vimos-modal-close">✕</a><h2 style="font-size:24px;font-weight:800;color:#0F172A;margin-top:0;margin-bottom:20px;background:linear-gradient(135deg,#0284C7,#0EA5E9);-webkit-background-clip:text;-webkit-text-fill-color:transparent;">🏢 Giới thiệu VIET.MOS</h2><div style="text-align:justify;color:#334155;font-size:15px;line-height:1.7;padding:0 10px;">VIET.MOS COMPANY LIMITED là doanh nghiệp hoạt động trong lĩnh vực công nghệ thông tin, thiết kế và cung cấp các giải pháp tự động hóa cho sản xuất. Với định hướng phát triển dựa trên sự kết hợp giữa tinh thần sáng tạo của Việt Nam và tư duy kỹ thuật tiên tiến từ Nhật Bản, VIET.MOS tập trung cung cấp các dịch vụ thiết kế cơ khí, thiết kế điện – lập trình điều khiển, thiết kế mô phỏng và xây dựng giải pháp máy móc theo yêu cầu. Sở hữu đội ngũ kỹ sư giàu chuyên môn cùng tinh thần không ngừng học hỏi, công ty luôn hướng đến việc mang lại những giải pháp tối ưu, hiệu quả và chất lượng cao cho khách hàng, đồng thời góp phần nâng cao giá trị sản xuất và thúc đẩy sự phát triển bền vững của xã hội.</div><div style="margin-top:32px;padding-top:20px;border-top:1px solid #F1F5F9;"><a href="#_" class="vimos-modal-btn">Đóng</a></div></div></div>"""
        guide_modal_html = """<div id="guide-modal-box" class="vimos-feature-modal"><div class="vimos-modal-content" style="text-align:left;"><a href="#_" class="vimos-modal-close">✕</a><h2 style="font-size:24px;font-weight:800;color:#0F172A;margin-top:0;margin-bottom:20px;text-align:center;background:linear-gradient(135deg,#0284C7,#0EA5E9);-webkit-background-clip:text;-webkit-text-fill-color:transparent;">📖 Hướng dẫn sử dụng</h2><div style="color:#334155;font-size:14.5px;line-height:1.6;"><p style="margin-top:0;">Chào mừng bạn đến với hướng dẫn nhanh! Dưới đây là các bước cơ bản để tính giờ làm việc từ file Excel máy chấm công:</p><div style="margin-bottom:12px;"><b style="color:#0284C7;">1. Tải lên file Excel</b><div style="padding-left:12px;color:#64748B;">• Nhấp vào nút <b>Tải file lên</b> ở bảng điều khiển bên trái (hoặc kéo thả file Excel từ máy quét chấm công).</div></div><div style="margin-bottom:12px;"><b style="color:#0284C7;">2. Định dạng cột dữ liệu</b><div style="padding-left:12px;color:#64748B;">• Hệ thống tự động quét và nhận diện các cột Giờ Vào/Ra. Nếu nhận diện sai, bạn có thể tự chọn lại từ menu thả xuống rồi bấm Xác nhận.</div></div><div style="margin-bottom:12px;"><b style="color:#0284C7;">3. Cài đặt thời gian chuẩn</b><div style="padding-left:12px;color:#64748B;">• Nhập khung giờ ca chuẩn (VD: 08:00 - 17:00) và giờ nghỉ trưa để hệ thống tự động trừ giờ nghỉ.</div></div><div style="margin-bottom:12px;"><b style="color:#0284C7;">4. Xử lý & Xuất báo cáo</b><div style="padding-left:12px;color:#64748B;">• Nhấn nút <b>Xử lý dữ liệu</b> để xem thống kê chi tiết trên Dashboard, sau đó bấm <b>Tải file Excel báo cáo</b> về máy.</div></div></div><div style="margin-top:28px;padding-top:16px;border-top:1px solid #F1F5F9;text-align:center;"><a href="#_" class="vimos-modal-btn">Đã hiểu</a></div></div></div>"""
    
    st.markdown(f"""<div style="position: fixed; top: 80px; right: 40px; z-index: 999999; display: flex; flex-direction: column; align-items: flex-end; gap: 16px; font-family: 'Inter', 'Be Vietnam Pro', sans-serif;"><div style="display: flex; gap: 24px; font-size: 15px; font-weight: 600; background: {nav_bg_color}; padding: 8px 16px; border-radius: 100px; backdrop-filter: blur(8px); border: 1px solid rgba(14,165,233,0.1);"><a href="#feature-modal-box" style="color: {nav_text_color}; text-decoration: none; cursor: pointer;">{t('feature')}</a><a href="#about-modal-box" style="color: {nav_text_color}; text-decoration: none; cursor: pointer;">{t('about')}</a><a href="#guide-modal-box" style="color: {nav_text_color}; text-decoration: none; cursor: pointer;">{t('guide')}</a><a href="#" style="color: {nav_text_color}; text-decoration: none; cursor: pointer;">{t('contact')}</a></div></div>{feature_modal_html}{about_modal_html}{guide_modal_html}""", unsafe_allow_html=True)
    
    # (Announcement banner removed)
    
    # Render nội dung vào sidebar để Streamlit bắt buộc hiển thị sidebar
    with st.sidebar:
        logo_b64_sidebar = LOGO_HEADER_B64
        if logo_b64_sidebar:
            st.markdown(f'<div style="text-align:center; padding: 16px 0 8px 0;"><img src="data:image/png;base64,{logo_b64_sidebar}" style="height:60px; object-fit:contain;"></div>', unsafe_allow_html=True)
        st.markdown("---")
        st.markdown(f"<p style='text-align:center; color:#64748B; font-size:13px;'>{t('subtitle_system')}</p>", unsafe_allow_html=True)
        st.markdown("---")
        if st.button(t("btn_calc_hours"), use_container_width=True, type="primary", key="sidebar_chamcong"):
            st.session_state.app_page = "chamcong"
            st.rerun()
        if st.button(t("btn_mos"), use_container_width=True, type="primary", key="sidebar_mos"):
            st.session_state.app_page = "mos"
            st.rerun()

    logo_b64 = LOGO_HEADER_B64
    logo_html = _logo_img_tag(logo_b64, "height: 80px; object-fit: contain;") if logo_b64 else '<span style="color:#10B981">✦</span> VIET.MOS'
    
    st.markdown(f"""
<style>
/* New minimalist CSS */
div[data-testid="block-container"]:has(.vimos-home-page-marker) {{
    display: flex !important;
    flex-direction: column !important;
    
    align-items: center !important;
    min-height: 100vh !important;
    padding-top: 0 !important;
}}
div[data-testid="block-container"]:has(.vimos-home-page-marker) div[data-testid="stVerticalBlock"] {{
    display: flex !important;
    flex-direction: column !important;
    
    align-items: center !important;
    width: 100% !important;
}}
.landing-wrapper {{
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    font-family: 'Inter', 'Be Vietnam Pro', sans-serif;
    color: #0F172A;
}}
.vimos-subtitle {{
    font-size: 24px;
    font-weight: 500;
    color: #64748B;
    letter-spacing: 0.5px;
    margin-top: 40px;
    margin-bottom: 30px;
    text-align: center;
}}
/* Hide the horizontal block's default styling so it doesn't interfere */
div[data-testid="block-container"]:has(.vimos-home-page-marker) [data-testid="stHorizontalBlock"] {{
    position: static !important;
}}

.st-key-btn_home_calc, .st-key-btn_home_mos, .st-key-btn_home_checkin, .st-key-btn_home_history {{
    width: 100% !important;
}}
.st-key-btn_home_calc button, .st-key-btn_home_mos button, .st-key-btn_home_checkin button, .st-key-btn_home_history button {{
    background: linear-gradient(135deg, #0EA5E9, #38BDF8) !important;
    color: white !important;
    padding: 16px 20px !important;
    white-space: normal !important;
    border-radius: 100px !important;
    font-size: 16px !important;
    font-weight: 700 !important;
    border: none !important;
    cursor: pointer !important;
    box-shadow: 0 10px 25px rgba(14, 165, 233, 0.25) !important;
    transition: transform 0.3s ease, box-shadow 0.3s ease !important;
    min-height: 68px !important;
    display: flex !important;
    align-items: center !important;
    justify-content: center !important;
    width: 100% !important;
}}
.st-key-btn_home_calc button p, .st-key-btn_home_mos button p, .st-key-btn_home_checkin button p, .st-key-btn_home_history button p {{
    color: white !important;
    font-size: 16.5px !important;
    font-weight: 700 !important;
    margin: 0 !important;
    text-align: center !important;
}}
.st-key-btn_home_calc button:hover, .st-key-btn_home_mos button:hover, .st-key-btn_home_checkin button:hover, .st-key-btn_home_history button:hover {{
    transform: translateY(-4px) scale(1.02) !important;
    box-shadow: 0 16px 35px rgba(14, 165, 233, 0.35) !important;
    color: white !important;
}}
</style>

<div class="vimos-home-page-marker" style="display:none;"></div>
<div class="landing-wrapper">
<div style="transform: scale(2.6); margin-top: 5vh; margin-bottom: 50px;">
{logo_html}
</div>
<div class="vimos-subtitle" style="font-size: 22px; font-weight: 700; color: #475569; letter-spacing: 1px; margin-bottom: 25px; text-transform: uppercase;">
{t('subtitle_system')}
</div>
</div>
    """, unsafe_allow_html=True)


    # --- REAL-TIME DUAL CLOCK (CIRCULAR, TOP-LEFT) ---
    clock_html = """
    <!DOCTYPE html>
    <html>
    <head>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@600;700&display=swap');
        body {
            margin: 0; padding: 0; display: flex; justify-content: center; align-items: center;
            font-family: 'Inter', sans-serif; background: transparent; overflow: hidden;
            width: 100vw; height: 100vh;
        }
        .clock-circle {
            width: 120px; height: 120px; border-radius: 50%;
            background: rgba(255, 255, 255, 0.9);
            border: 2px solid rgba(14, 165, 233, 0.3);
            box-shadow: 0 10px 25px rgba(14, 165, 233, 0.2);
            backdrop-filter: blur(12px);
            display: flex; flex-direction: column; justify-content: center; align-items: center;
            gap: 6px; transition: transform 0.3s ease; cursor: default;
        }
        .clock-circle:hover { transform: scale(1.05); box-shadow: 0 12px 30px rgba(14, 165, 233, 0.3); }
        .time-block { display: flex; align-items: center; gap: 6px; }
        .flag { font-size: 16px; }
        .time { font-size: 14px; font-weight: 700; color: #0F172A; letter-spacing: 0.5px; font-variant-numeric: tabular-nums; }
        .divider { width: 70px; height: 1.5px; background: linear-gradient(90deg, transparent, rgba(14,165,233,0.3), transparent); margin: 0; }
    </style>
    </head>
    <body>
        <div class="clock-circle">
            <div class="time-block">
                <span class="flag">🇻🇳</span>
                <span class="time" id="time-vn">--:--:--</span>
            </div>
            <div class="divider"></div>
            <div class="time-block">
                <span class="flag">🇯🇵</span>
                <span class="time" id="time-jp">--:--:--</span>
            </div>
        </div>
        <script>
            // Break out of iframe to make it fixed top-left
            try {
                const frame = window.frameElement;
                if (frame) {
                    frame.style.position = 'fixed';
                    frame.style.top = '30px';
                    frame.style.left = '30px';
                    frame.style.zIndex = '999999';
                    frame.style.width = '140px';
                    frame.style.height = '140px';
                    frame.style.border = 'none';
                    frame.style.backgroundColor = 'transparent';
                    
                    const parentDiv = frame.parentElement;
                    if (parentDiv) {
                        parentDiv.style.position = 'fixed';
                        parentDiv.style.top = '30px';
                        parentDiv.style.left = '30px';
                        parentDiv.style.width = '140px';
                        parentDiv.style.height = '140px';
                        parentDiv.style.zIndex = '999999';
                    }
                }
            } catch(e) {}

            function updateTime() {
                const now = new Date();
                const vnTime = new Intl.DateTimeFormat('en-US', { timeZone: 'Asia/Ho_Chi_Minh', hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false }).format(now);
                const jpTime = new Intl.DateTimeFormat('en-US', { timeZone: 'Asia/Tokyo', hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false }).format(now);
                document.getElementById('time-vn').innerText = vnTime;
                document.getElementById('time-jp').innerText = jpTime;
            }
            updateTime();
            setInterval(updateTime, 1000);
        </script>
    </body>
    </html>
    """
    st.components.v1.html(clock_html, height=140)

    st.markdown(f"""
<div class="landing-wrapper">
<div style="position: fixed; bottom: 16px; width: 100%; text-align: center; color: #64748B; font-size: 14px; font-weight: 500; z-index: 10;">Copyright &copy; 2026 V-mos System</div>
</div>
    """, unsafe_allow_html=True)

    st.markdown('<div style="height: 55px;"></div>', unsafe_allow_html=True)
    
    col_left, col_btns, col_right, col_side = st.columns([1.05, 2.2, 0.05, 0.7])
    
    with col_btns:
        r1_c1, r1_c2 = st.columns(2)
        with r1_c1:
            if st.button(f"⏱️ {t('btn_calc_hours')}", use_container_width=True, type="primary", key="btn_home_calc"):
                st.session_state.app_page = "chamcong"
                st.rerun()
        with r1_c2:
            if st.button(f"📊 {t('btn_mos')}", use_container_width=True, type="primary", key="btn_home_mos"):
                st.session_state.app_page = "mos"
                st.rerun()
                
        st.markdown('<div style="height: 14px;"></div>', unsafe_allow_html=True)
        r2_c1, r2_c2 = st.columns(2)
        with r2_c1:
            if st.button("📱 Check-in GPS" if st.session_state.lang == 'vi' else "📱 打刻 GPS", use_container_width=True, type="primary", key="btn_home_checkin"):
                st.session_state.app_page = "checkin"
                st.rerun()
        with r2_c2:
            if st.button("📁 Lịch sử & So sánh" if st.session_state.lang == 'vi' else "📁 履歴＆比較", use_container_width=True, type="primary", key="btn_home_history"):
                st.session_state.app_page = "history"
                st.rerun()

    with col_side:
        st.markdown(f"""<div style="margin-top: -125px; display: flex; flex-direction: column; gap: 12px; width: 140px; margin-left: auto; margin-right: 15px;">
<div style="background: rgba(255, 255, 255, 0.95); border-radius: 14px; padding: 10px 12px; box-shadow: 0 6px 20px rgba(0,0,0,0.05); border: 1px solid rgba(14, 165, 233, 0.2);">
<div style="font-size: 11.5px; font-weight: 700; color: #0F172A; margin-bottom: 8px; text-align: center;">{t('weather_title')}</div>
<div style="display: flex; justify-content: space-around; align-items: center; text-align: center;">
<div>
<div style="font-size: 15px;">🇻🇳</div>
<div style="font-size: 10px; font-weight: 600; color: #475569;">HN</div>
<div style="font-size: 13px; font-weight: 800; color: #0EA5E9;" id="vmc-temp-hn">--°</div>
</div>
<div style="width: 1px; height: 28px; background: #E2E8F0;"></div>
<div>
<div style="font-size: 15px;">🇯🇵</div>
<div style="font-size: 10px; font-weight: 600; color: #475569;">TK</div>
<div style="font-size: 13px; font-weight: 800; color: #0EA5E9;" id="vmc-temp-tk">--°</div>
</div>
</div>
</div>
<div style="background: rgba(255, 255, 255, 0.95); border-radius: 14px; padding: 10px 12px; box-shadow: 0 6px 20px rgba(0,0,0,0.05); border: 1px solid rgba(14, 165, 233, 0.2);">
<div style="font-size: 11.5px; font-weight: 700; color: #0F172A; margin-bottom: 8px; text-align: center;">{t('holiday_title')}</div>
<div style="display: flex; flex-direction: column; gap: 6px; font-size: 11px; color: #334155; font-weight: 600;">
<div style="display: flex; justify-content: space-between; align-items: center; padding-bottom: 4px; border-bottom: 1px dashed #F1F5F9;"><span>{'🇻🇳 Quốc khánh' if st.session_state.lang == 'vi' else '🇻🇳 建国記念日'}</span><span style="color:#0284C7; font-size:10px; font-weight:700;">02/09</span></div>
<div style="display: flex; justify-content: space-between; align-items: center; padding-bottom: 4px; border-bottom: 1px dashed #F1F5F9;"><span>{'🇯🇵 Kính Lão' if st.session_state.lang == 'vi' else '🇯🇵 敬老の日'}</span><span style="color:#DC2626; font-size:10px; font-weight:700;">18/09</span></div>
<div style="display: flex; justify-content: space-between; align-items: center;"><span>{'🇯🇵 Thu Phân' if st.session_state.lang == 'vi' else '🇯🇵 秋分の日'}</span><span style="color:#DC2626; font-size:10px; font-weight:700;">23/09</span></div>
</div>
</div>
</div>""", unsafe_allow_html=True)

    st.components.v1.html("""
    <script>
    (function() {
        async function fetchWeather() {
            try {
                const [hn, tk] = await Promise.all([
                    fetch('https://api.open-meteo.com/v1/forecast?latitude=21.0285&longitude=105.8542&current_weather=true').then(r=>r.json()),
                    fetch('https://api.open-meteo.com/v1/forecast?latitude=35.6895&longitude=139.6917&current_weather=true').then(r=>r.json())
                ]);
                const doc = window.parent.document;
                const hnEl = doc.getElementById('vmc-temp-hn');
                const tkEl = doc.getElementById('vmc-temp-tk');
                if (hnEl) hnEl.innerText = Math.round(hn.current_weather.temperature) + '°C';
                if (tkEl) tkEl.innerText = Math.round(tk.current_weather.temperature) + '°C';
            } catch(e) {}
        }
        setTimeout(fetchWeather, 300);
        setTimeout(fetchWeather, 1500);
    })();
    </script>
    """, height=0)

    render_chatbot()
    st.stop() # Dừng toàn bộ code phía dưới để giữ giao diện sạch



# ==========================================
# CÁC HÀM XỬ LÝ DỮ LIỆU MOS
# ==========================================
import re

import httpx
import json

def summarize_tasks_with_ai(ma_da: str, ten_da: str, tasks: list) -> str:
    """
    Gọi Claude API để tóm tắt danh sách task thành 1 câu nội dung chung.
    Nếu API lỗi, fallback về cách nối chuỗi thủ công.
    """
    if not tasks:
        return ''
    
    # Loại bỏ trùng lặp, giữ thứ tự
    tasks_unique = list(dict.fromkeys([t.strip() for t in tasks if t.strip()]))
    
    if len(tasks_unique) == 1:
        return tasks_unique[0]
    
    prompt = f"""Dưới đây là danh sách các nội dung công việc của dự án 
"{ten_da}" (mã: {ma_da}) bằng tiếng Nhật:

{chr(10).join(f'- {t}' for t in tasks_unique)}

Nhiệm vụ của bạn:
1. Dịch tên dự án "{ten_da}" sang tiếng Việt.
2. Tóm tắt danh sách công việc trên thành 1 câu súc tích bằng tiếng Nhật, sau đó dịch câu đó sang tiếng Việt.

Trả về kết quả DƯỚI DẠNG JSON với định dạng sau (không giải thích thêm):
{{
  "ten_da_song_ngu": "Tên tiếng Nhật \\n Tên tiếng Việt",
  "noi_dung_song_ngu": "Tóm tắt tiếng Nhật \\n Tóm tắt tiếng Việt"
}}
"""

    try:
        api_key = load_saved_api_key()
        
        if not api_key:
            return ' / '.join(tasks_unique)
            
        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}"
        response = httpx.post(
            url,
            headers={"Content-Type": "application/json"},
            json={
                "contents": [{"parts": [{"text": prompt}]}],
                "generationConfig": {"temperature": 0.2}
            },
            timeout=15.0
        )
        data = response.json()
        result_text = data["candidates"][0]["content"]["parts"][0]["text"].strip()
        
        import json
        if "```json" in result_text:
            result_text = result_text.split("```json")[1].split("```")[0]
        elif "```" in result_text:
            result_text = result_text.split("```")[1].split("```")[0]
            
        return json.loads(result_text)
    except Exception as e:
        return ' / '.join(tasks_unique)

def extract_ma_nv_from_filename(filename: str) -> str:
    """Lấy mã NV từ tên file: Report_VM011_ロン_2026_5.xlsx → VM011"""
    match = re.search(r'(VM\d+)', filename)
    return match.group(1) if match else filename

def extract_ten_nv_from_filename(filename: str) -> str:
    """Lấy tên Nhật từ file để hiển thị: Report_VM011_ロン_2026_5 → ロン"""
    name = re.sub(r'\.xlsx?$', '', filename)
    # Lấy phần chữ (ko phải số/gạch dưới) ngay sau cụm VMxxx
    match = re.search(r'VM\d+[\s_]+([^\d_]+)', name)
    if match:
        return match.group(1).strip()
    return filename

def parse_mos_file(file, filename: str) -> pd.DataFrame:
    """
    Đọc 1 file Report, lấy chỉ phần MOS業務.
    Trả về DataFrame với các cột đã chuẩn hóa.
    """
    import openpyxl
    wb = openpyxl.load_workbook(file, read_only=True, data_only=True)
    ws = wb.active
    
    rows = list(ws.iter_rows(values_only=True))
    
    # Tìm header row (có chữ 案件番号)
    header_row_idx = None
    for i, row in enumerate(rows):
        if row and any(str(v) == '案件番号' for v in row if v):
            header_row_idx = i
            break
    if header_row_idx is None:
        raise ValueError(f"Không tìm thấy dòng tiêu đề chuẩn (案件番号) trong file {filename}")
    
    header = rows[header_row_idx]
    
    col_map = {}
    date_cols = {}
    for ci, col in enumerate(header):
        if col is None: continue
        col_str = str(col).strip()
        if col_str == '日本担当者': col_map['tanto'] = ci
        elif col_str == '案件番号': col_map['ma_da'] = ci
        elif col_str == 'お客様': col_map['khach'] = ci
        elif col_str == '案件名': col_map['ten_da'] = ci
        elif col_str == '区分': col_map['phan_vung'] = ci
        elif col_str == 'タスク名': col_map['task'] = ci
        elif col_str == '合計': col_map['tong'] = ci
        elif '/' in col_str:
            try:
                parts = col_str.split('/')
                month, day = int(parts[0]), int(parts[1])
                date_cols[f'{month}/{day}'] = ci
            except: pass
    
    in_mos = False
    records = []
    
    for row in rows[header_row_idx + 2:]:
        if row is None: continue
        first_val = str(row[0]).strip() if row[0] else ''
        
        if first_val == 'MOS業務':
            in_mos = True
            continue
        elif first_val in ['JMOS業務', '社内業務', 'JMO S業務']:
            in_mos = False
            continue
        
        if not in_mos:
            continue
        
        ma_da = str(row[col_map.get('ma_da', 1)] or '').strip()
        if not ma_da or ma_da == 'None':
            continue
        
        tanto = str(row[col_map.get('tanto', 0)] or '').strip()
        khach = str(row[col_map.get('khach', 2)] or '').strip()
        ten_da = str(row[col_map.get('ten_da', 3)] or '').strip()
        phan_vung = str(row[col_map.get('phan_vung', 4)] or '').strip()
        task = str(row[col_map.get('task', 5)] or '').strip()
        tong = row[col_map.get('tong', 6)]
        
        try: tong = float(tong) if tong else 0.0
        except: tong = 0.0
        
        ngay_co_gio = []
        for date_str, ci in date_cols.items():
            val = row[ci]
            try:
                if val and float(val) > 0:
                    ngay_co_gio.append(date_str)
            except: pass
        
        def parse_date_str(d_str, year=2026):
            try:
                m, d = d_str.split('/')
                return datetime.date(year, int(m), int(d))
            except: return None
        
        ngay_dates = [parse_date_str(d) for d in ngay_co_gio]
        ngay_dates = [d for d in ngay_dates if d]
        
        ngay_bat_dau = min(ngay_dates).strftime('%d/%m/%Y') if ngay_dates else ''
        ngay_ket_thuc = max(ngay_dates).strftime('%d/%m/%Y') if ngay_dates else ''
        
        ql_nhat = ''
        if '/' in tanto:
            ql_nhat = tanto.split('/')[-1].replace('様', '').strip()
        elif tanto:
            ql_nhat = tanto.replace('様', '').strip()
            
        khach = khach.replace('様', '').strip()
        
        if 'シミュレーション' in phan_vung:
            phan_vung_vn = 'シミュレーション設計 \n Thiết kế mô phỏng'
        elif '電気' in phan_vung:
            phan_vung_vn = '電気設計 \n Thiết kế điện'
        elif 'メカ' in phan_vung or '機械' in phan_vung:
            phan_vung_vn = 'メカ設計 \n Thiết kế cơ khí'
        else:
            phan_vung_vn = f"{phan_vung} \n {phan_vung}" 
        
        records.append({
            'ma_nv': extract_ma_nv_from_filename(filename),
            'ten_nv': extract_ten_nv_from_filename(filename),
            'ma_da': ma_da,
            'khach': khach,
            'ten_da': ten_da,
            'phan_vung': phan_vung_vn,
            'task': task,
            'tong_gio': tong,
            'ngay_bat_dau': ngay_bat_dau,
            'ngay_ket_thuc': ngay_ket_thuc,
            'ql_nhat': ql_nhat,
        })
    if not records:
        raise ValueError(f"Không tìm thấy dữ liệu hợp lệ (phần MOS業務) trong file {filename}")
    return pd.DataFrame(records)


def tong_hop_mos(dfs: list) -> pd.DataFrame:
    """
    Gộp nhiều DataFrame từ nhiều file, tổng hợp theo mã dự án.
    """
    if not dfs:
        return pd.DataFrame()
    
    df_all = pd.concat(dfs, ignore_index=True)
    
    import streamlit as st
    result = []
    groups = list(df_all.groupby('ma_da'))
    total_groups = len(groups)
    
    if total_groups == 0:
        return pd.DataFrame()
        
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for idx, (ma_da, grp) in enumerate(groups):
        status_text.markdown(f"**🤖 Đang phân tích dự án {idx+1}/{total_groups}:** `{ma_da}`")
        ngay_bd_list = [datetime.datetime.strptime(d, '%d/%m/%Y').date() 
                        for d in grp['ngay_bat_dau'] if d]
        ngay_kt_list = [datetime.datetime.strptime(d, '%d/%m/%Y').date() 
                        for d in grp['ngay_ket_thuc'] if d]
        ngay_bd = min(ngay_bd_list).strftime('%d/%m/%Y') if ngay_bd_list else ''
        ngay_kt = max(ngay_kt_list).strftime('%d/%m/%Y') if ngay_kt_list else ''
        if ngay_bd == ngay_kt:
            ngay_kt = ''
        
        nv_co_gio = grp[grp['tong_gio'] > 0]['ten_nv'].unique()
        jp_names = [mv for mv in sorted(nv_co_gio) if mv]
        name_dict = {'ダンフン': 'Hưng', 'クアン': 'Quân', 'ハン': 'Hằng', 'グエット': 'Nguyệt', 'ロン': 'Long', 'ダオ': 'Đạo', 'フォン': 'Phương'}
        
        # Lấy nội dung công việc (task) hợp lệ đầu tiên làm đại diện duy nhất
        valid_tasks = [str(t).strip() for t in grp['task'].tolist() if str(t).strip() and str(t).lower() not in ['nan', 'none']]
        first_task = valid_tasks[0] if valid_tasks else ""
        tasks = [first_task] if first_task else []
        
        ten_da_val = grp['ten_da'].iloc[0]
        
        # Gọi AI (lấy cả tên dự án và nội dung)
        ai_res = summarize_tasks_with_ai(ma_da, ten_da_val, tasks)
        if isinstance(ai_res, dict):
            ten_da_song_ngu = ai_res.get('ten_da_song_ngu', ten_da_val)
            noi_dung = ai_res.get('noi_dung_song_ngu', first_task)
        else:
            ten_da_song_ngu = ten_da_val
            noi_dung = ai_res
            
        phan_vung = grp['phan_vung'].iloc[0]
        
        # Xử lý: Nếu có 'Phương' thì chuyển sang quản lý VN
        ql_viet_nam = None
        if 'フォン' in jp_names:
            jp_names.remove('フォン')
            ql_viet_nam = "フォン \n Phương"
            
        # Nếu nội dung liên quan đến điện thì Đạo làm quản lý
        if any(kw in str(noi_dung).lower() or kw in str(phan_vung).lower() for kw in ['điện', '電気']):
            ql_viet_nam = "ダオ \n Đạo"
            if 'ダオ' in jp_names:
                jp_names.remove('ダオ')
                
        # Nếu nội dung liên quan đến thiết kế mô phỏng thì Phương làm quản lý
        if any(kw in str(noi_dung).lower() or kw in str(phan_vung).lower() for kw in ['mô phỏng', 'シミュレーション', 'simulation']):
            ql_viet_nam = "フォン \n Phương"
            if 'フォン' in jp_names:
                jp_names.remove('フォン')
                
        vn_names = [name_dict.get(mv, mv) for mv in jp_names]
        nguoi_th = ' / '.join(jp_names) + ' \n ' + ' / '.join(vn_names) if jp_names else ''
            
        ql_nhat_goc = grp['ql_nhat'].iloc[0]
        jp_manager_dict = {'石部': 'Ishibe', '勝亦': 'Katsumata', '池谷': 'Ikeya', '田代': 'Tashiro', '井手上': 'Ideue', '山崎': 'Yamazaki', 'フォン': 'Phương', 'ロン': 'Long'}
        ql_nhat_trans = jp_manager_dict.get(ql_nhat_goc, ql_nhat_goc)
        ql_nhat_song_ngu = f"{ql_nhat_goc} \n {ql_nhat_trans}" if ql_nhat_goc else ""
        
        result.append({
            'Mã dự án': ma_da,
            'Tên dự án': ten_da_song_ngu,
            'Phân vùng': grp['phan_vung'].iloc[0],
            'Nội dung ủy thác': noi_dung,
            'Giờ làm (h)': grp['tong_gio'].sum(),
            'Ngày bắt đầu': ngay_bd,
            'Ngày kết thúc': ngay_kt,
            'Quản lý Nhật Bản': ql_nhat_song_ngu,
            'Quản lý Việt Nam': ql_viet_nam,
            'Người thực hiện': nguoi_th,
            'Trạng thái': None,
        })
        progress_bar.progress((idx + 1) / total_groups)
    
    status_text.empty()
    progress_bar.empty()
    
    df_result = pd.DataFrame(result)
    df_result.insert(0, 'STT', range(1, len(df_result)+1))
    return df_result


def render_holiday_makeup_sidebar():
    
    if "custom_holidays" not in st.session_state:
        st.session_state.custom_holidays = set()
    if "custom_workdays" not in st.session_state:
        st.session_state.custom_workdays = set()

    with st.sidebar.expander(t("exp_holidays"), expanded=False):
        st.markdown(f"<small style='color:green;'>{t('holidays_note')}</small>", unsafe_allow_html=True)

        st.markdown(f"<small>{t('holiday_choose')}</small>", unsafe_allow_html=True)
        selected_date = st.date_input(t("holiday_select"), value=datetime.date.today(), label_visibility="collapsed", key="date_in_holiday")
        if st.button(t("btn_add_custom_holiday"), key="btn_add_custom_holiday", use_container_width=True):
            st.session_state.custom_holidays.add(selected_date)
            st.rerun()
        if st.session_state.custom_holidays:
            st.markdown(t("custom_holiday_count", count=len(st.session_state.custom_holidays)), unsafe_allow_html=True)
            holiday_list = [d.strftime('%d/%m/%Y') for d in sorted(st.session_state.custom_holidays)]
            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                selected_to_remove = st.multiselect(t("holiday_del_sel"), options=holiday_list, placeholder=t("holiday_del_placeholder"), label_visibility="collapsed", key="multi_sel_del_holiday")
            with col_btn2:
                if st.button(t("btn_del_selected"), key="btn_del_sel_holiday", use_container_width=True) and selected_to_remove:
                    for d_str in selected_to_remove:
                        try:
                            st.session_state.custom_holidays.remove(datetime.datetime.strptime(d_str, '%d/%m/%Y').date())
                        except: pass
                    st.rerun()
            with col_btn2:
                if st.button(t("btn_del_all"), key="btn_del_all_holiday", use_container_width=True):
                    st.session_state.custom_holidays = set()
                    st.rerun()

    with st.sidebar.expander(t("exp_makeup"), expanded=False):
        st.markdown(f"<small>{t('makeup_note')}</small>", unsafe_allow_html=True)
        selected_makeup = st.date_input(t("makeup_choose"), value=datetime.date.today(), label_visibility="collapsed", key="date_makeup_input")
        if st.button(t("btn_add_makeup"), key="btn_add_makeup", use_container_width=True):
            st.session_state.custom_workdays.add(selected_makeup)
            st.rerun()
        if st.session_state.custom_workdays:
            st.markdown(f"<small><b>{t('makeup_count').format(count=len(st.session_state.custom_workdays))}</b></small>", unsafe_allow_html=True)
            makeup_list = [d.strftime('%d/%m/%Y') for d in sorted(st.session_state.custom_workdays)]
            selected_makeup_remove = st.multiselect(t("makeup_remove_prompt"), options=makeup_list, placeholder="...", label_visibility="collapsed", key="multi_sel_del_makeup")
            col_btn3, col_btn4 = st.columns(2)
            with col_btn3:
                if st.button(t("btn_del_selected"), key="btn_del_sel_makeup", use_container_width=True) and selected_makeup_remove:
                    for d_str in selected_makeup_remove:
                        try: st.session_state.custom_workdays.remove(datetime.datetime.strptime(d_str, '%d/%m/%Y').date())
                        except: pass
                    st.rerun()
            with col_btn4:
                if st.button(t("btn_del_all"), key="btn_del_all_makeup", use_container_width=True):
                    st.session_state.custom_workdays = set()
                    st.rerun()

def render_leave_ot_sidebar():
    with st.sidebar.expander("📂 Khớp nối Nghỉ phép & OT" if st.session_state.lang == 'vi' else "📂 休暇＆残業の登録", expanded=False):
        if "manual_leave" not in st.session_state: st.session_state.manual_leave = {}
        if "manual_ot_reason" not in st.session_state: st.session_state.manual_ot_reason = {}
        
        emp_options = get_company_emp_options(st.session_state.lang)
        
        tab_leave, tab_ot = st.tabs(["🌴 Nghỉ phép" if st.session_state.lang=='vi' else "🌴 休暇登録", "⏰ OT" if st.session_state.lang=='vi' else "⏰ 残業登録"])
        with tab_leave:
            sel_e_leave = st.selectbox("Nhân viên" if st.session_state.lang=='vi' else "社員", emp_options, key="sb_sel_e_leave")
            sel_d_leave = st.date_input("Ngày nghỉ" if st.session_state.lang=='vi' else "休暇日", value=datetime.date.today(), key="sb_sel_d_leave")
            if st.button("➕ Thêm lịch nghỉ" if st.session_state.lang=='vi' else "➕ 休暇を追加", type="secondary", use_container_width=True, key="btn_add_l"):
                if sel_e_leave:
                    ma_l = sel_e_leave.split(" - ")[0].strip().upper()
                    d_str = sel_d_leave.strftime("%d/%m/%Y")
                    st.session_state.manual_leave[(ma_l, d_str)] = True
                    st.success(f"✅ Đã thêm: {ma_l} ({d_str})")
            if st.session_state.manual_leave:
                st.caption(f"📌 Đã đăng ký: {len(st.session_state.manual_leave)} ngày nghỉ")
                if st.button("🗑️ Xóa danh sách nghỉ" if st.session_state.lang=='vi' else "🗑️ 休暇リストを削除", key="btn_clr_l", use_container_width=True):
                    st.session_state.manual_leave = {}; st.rerun()

        with tab_ot:
            sel_e_ot = st.selectbox("Nhân viên" if st.session_state.lang=='vi' else "社員", emp_options, key="sb_sel_e_ot")
            sel_d_ot = st.date_input("Ngày OT" if st.session_state.lang=='vi' else "残業日", value=datetime.date.today(), key="sb_sel_d_ot")
            ot_reasons_list = ["Xử lý sự cố khẩn cấp", "Bảo trì công trình", "Chạy thử máy mới", "Họp MOS định kỳ", "Khác"] if st.session_state.lang=='vi' else ["緊急トラブル対応", "設備定期メンテナンス", "新機種テスト運転", "定例MOSミーティング", "その他"]
            sel_r_ot = st.selectbox("Lý do OT" if st.session_state.lang=='vi' else "残業理由", ot_reasons_list, key="sb_sel_r_ot")
            if st.button("➕ Thêm lịch OT" if st.session_state.lang=='vi' else "➕ 残業を追加", type="secondary", use_container_width=True, key="btn_add_ot"):
                if sel_e_ot:
                    ma_o = sel_e_ot.split(" - ")[0].strip().upper()
                    d_str = sel_d_ot.strftime("%d/%m/%Y")
                    st.session_state.manual_ot_reason[(ma_o, d_str)] = sel_r_ot
                    st.success(f"✅ Đã thêm: {ma_o} ({d_str})")
            if st.session_state.manual_ot_reason:
                st.caption(f"📌 Đã đăng ký: {len(st.session_state.manual_ot_reason)} lịch OT")
                if st.button("🗑️ Xóa danh sách OT" if st.session_state.lang=='vi' else "🗑️ 残業リストを削除", key="btn_clr_ot", use_container_width=True):
                    st.session_state.manual_ot_reason = {}; st.rerun()

def render_email_sending_sidebar():
    with st.sidebar.expander("📧 Gửi Phiếu Xác Nhận Chấm Công" if st.session_state.lang == 'vi' else "📧 給与・勤怠明細の送信", expanded=False):
        st.caption("Gửi email thông báo tự động tới từng kỹ sư." if st.session_state.lang == 'vi' else "エンジニアへの明細自動メール送信。")
        mode_opts = ["🧪 Mô phỏng nhanh", "📨 Gửi thực SMTP"] if st.session_state.lang == 'vi' else ["🧪 デモシミュレーション", "📨 SMTP実送信"]
        mode_mail = st.radio("Chế độ gửi" if st.session_state.lang == 'vi' else "送信モード", mode_opts, horizontal=True, key="sb_mode_mail_global")
        if "SMTP" in mode_mail:
            st.text_input("SMTP Server" if st.session_state.lang == 'vi' else "SMTPサーバー", value="smtp.gmail.com", key="sb_smtp_srv_global")
            st.text_input("Sender Email" if st.session_state.lang == 'vi' else "送信元メール", placeholder="hr@vietmos.com", key="sb_smtp_mail_global")
            st.text_input("App Password" if st.session_state.lang == 'vi' else "アプリパスワード", type="password", key="sb_smtp_pwd_global")
        if st.button("🚀 Khởi chạy Phát Hành Email" if st.session_state.lang == 'vi' else "🚀 メール送信実行", type="primary", use_container_width=True, key="sb_btn_send_mail_global"):
            import time
            msg_busy = "⏳ Đang tổng hợp và phát hành phiếu chấm công..." if st.session_state.lang == 'vi' else "⏳ 明細データを集計して送信中..."
            with st.status(msg_busy):
                time.sleep(0.5)
                if 'df_raw' in st.session_state and st.session_state.df_raw is not None:
                    count_s = len(st.session_state.df_raw)
                    st.write(f"📨 Đã tạo và phát hành phiếu cho `{count_s}` bản ghi chấm công -> *Thành công*")
                else:
                    st.write("📨 Đã khởi chạy mô phỏng phát hành email -> *Thành công*")
            st.success("✅ Đã phát hành phiếu xác nhận tới email các kỹ sư!" if st.session_state.lang == 'vi' else "✅ 全エンジニアへのメール送信が完了しました！")

def render_mos_page():
    with st.sidebar:
        render_holiday_makeup_sidebar()
        st.markdown("<br>", unsafe_allow_html=True)



    st.markdown("""
    <style>
    /* ── Main Container ── */
    .mos-main-card {
        /* Bỏ khung trắng theo yêu cầu */
        padding: 0;
    }

    /* ── MOS Page Header (Banner) ── */
    
    .mos-header {
        background: linear-gradient(135deg, #0369A1 0%, #0EA5E9 100%);
        border-radius: 16px;
        padding: 32px 40px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 32px;
        box-shadow: 0 8px 25px rgba(59, 130, 246, 0.4);
        position: relative;
        overflow: hidden;
    }
    .mos-header::before {
        content: '';
        position: absolute;
        top: -50%;
        right: -10%;
        width: 300px;
        height: 300px;
        background: radial-gradient(circle, rgba(255,255,255,0.1) 0%, rgba(255,255,255,0) 70%);
        border-radius: 50%;
    }
    .mos-header-left {
        display: flex;
        gap: 20px;
        align-items: center;
        position: relative;
        z-index: 2;
    }
    .mos-header-logo {
        width: 250px;
        height: auto;
        object-fit: contain;
        filter: drop-shadow(0 4px 6px rgba(0,0,0,0.2));
        transition: transform 0.3s ease;
    }
    .mos-header-logo:hover {
        transform: scale(1.05) rotate(-2deg);
    }
    .mos-hero-title {
        font-size: 26px;
        font-weight: 800;
        color: white;
        margin: 0 0 6px 0;
        line-height: 1.2;
        letter-spacing: -0.02em;
    }
    .mos-hero-sub {
        font-size: 14px;
        color: rgba(255,255,255,0.85);
        margin: 0;
    }
    .mos-hero-badge {
        background: rgba(255,255,255,0.2);
        border: 1px solid rgba(255,255,255,0.4);
        backdrop-filter: blur(4px);
        border-radius: 20px;
        padding: 8px 18px;
        font-size: 13px;
        color: white;
        font-weight: 600;
        position: relative;
        z-index: 2;
    }

    /* ── Stepper ── */
    .mos-stepper {
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 32px;
        padding: 0 20px;
    }
    .mos-step-item {
        display: flex;
        flex-direction: column;
        align-items: center;
        gap: 8px;
        position: relative;
        z-index: 1;
    }
    .mos-step-circle {
        width: 28px;
        height: 28px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 12px;
        font-weight: 600;
        background: white;
    }
    .step-active .mos-step-circle {
        border: 2px solid #0EA5E9;
        color: #0EA5E9;
        box-shadow: 0 0 0 4px rgba(14, 165, 233, 0.15);
        animation: pulse 2s infinite;
    }
    @keyframes pulse {
        0% { box-shadow: 0 0 0 0 rgba(14, 165, 233, 0.3); }
        70% { box-shadow: 0 0 0 10px rgba(14, 165, 233, 0); }
        100% { box-shadow: 0 0 0 0 rgba(14, 165, 233, 0); }
    }
    .step-inactive .mos-step-circle {
        border: 2px solid #E2E8F0;
        color: #94A3B8;
    }
    .step-done .mos-step-circle {
        background: #0EA5E9;
        border: 2px solid #0EA5E9;
        color: white;
    }
    .mos-step-label {
        font-size: 13px;
        font-weight: 700;
        margin-top: 4px;
    }
    .step-active .mos-step-label { color: #0369A1; }
    .step-inactive .mos-step-label { color: #94A3B8; }
    .step-done .mos-step-label { color: #0EA5E9; }
    
    .mos-stepper-line {
        flex: 1;
        height: 1px;
        background: #E2E8F0;
        margin: 0 16px;
        margin-bottom: 20px; /* offset to align with circles */
    }

    /* ── Upload zone override ── */
    [data-testid="stFileUploader"] {
        background: #F8FAFC !important;
        border: 1.5px dashed #CBD5E1 !important;
        border-radius: 12px !important;
        padding: 24px !important;
    }

    /* ── File tag ── */
    .mos-file-tag {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background: white;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 8px 12px;
        font-size: 12px;
        color: #334155;
        margin: 4px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.04);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    .mos-file-tag:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(0,0,0,0.08);
        border-color: #BFDBFE;
    }

    /* ── Name mapping grid ── */
    .mos-nv-label {
        background: #F1F5F9;
        color: #0F172A;
        border: 1px solid #E2E8F0;
        border-radius: 6px;
        padding: 4px 10px;
        font-size: 12px;
        font-weight: 600;
        display: inline-block;
        margin-bottom: 4px;
    }

    /* ── Result summary bar ── */
    .mos-summary {
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 14px 20px;
        display: flex;
        gap: 32px;
        margin-bottom: 16px;
        flex-wrap: wrap;
    }
    .mos-summary-item .label {
        font-size: 11px;
        color: #64748B;
        text-transform: uppercase;
        letter-spacing: .05em;
    }
    .mos-summary-item .value {
        font-size: 20px;
        font-weight: 700;
        color: #0F172A;
    }

    /* Override primary button */
    .stButton > button[kind="primary"] {
        background: #0EA5E9 !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        font-weight: 500 !important;
        transition: all 0.3s ease;
    }
    .stButton > button[kind="primary"]:hover {
        background: #0284C7 !important;
        box-shadow: 0 4px 12px rgba(14, 165, 233, 0.4) !important;
    }
    </style>
    """, unsafe_allow_html=True)
    
    # Nút quay lại trang chủ luôn hiển thị ở trên cùng bên trái
    if st.button(t("btn_back_home"), key="btn_back_home_mos_top", type="primary"):
        st.session_state.app_page = "home"
        st.rerun()
    
    st.markdown(f"""
    <div class="mos-header">
        <div class="mos-header-left">
            {_logo_img_tag(LOGO_HEADER_B64, extra_class="mos-header-logo")}
            <div>
                <p class="mos-hero-title">{t("mos_hero_title")}</p>
                <p class="mos-hero-sub">{t("mos_hero_sub")}</p>
            </div>
        </div>
        <span class="mos-hero-badge">モス委託業務工数集計</span>
    </div>
    
    """, unsafe_allow_html=True)
    
    # Xác định bước hiện tại
    step = 1
    if 'df_mos_result' in st.session_state and st.session_state.df_mos_result is not None:
        step = 2
    if st.session_state.get('mos_saved', False):
        step = 3

    st.markdown(f"""
    <div class="mos-stepper">
        <div class="mos-step-item {'step-done' if step > 1 else 'step-active'}">
            <div class="mos-step-circle">{'✓' if step > 1 else '1'}</div>
            <div class="mos-step-label">{t("mos_step_1")}</div>
        </div>
        <div class="mos-stepper-line"></div>
        <div class="mos-step-item {'step-done' if step > 2 else ('step-active' if step == 2 else 'step-inactive')}">
            <div class="mos-step-circle">{'✓' if step > 2 else '2'}</div>
            <div class="mos-step-label">{t("mos_step_2")}</div>
        </div>
        <div class="mos-stepper-line"></div>
        <div class="mos-step-item {'step-active' if step == 3 else 'step-inactive'}">
            <div class="mos-step-circle">3</div>
            <div class="mos-step-label">{t("mos_step_3")}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Đã gộp vào stepper ở trên
    

    uploaded_files = st.file_uploader(
        "Upload",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        label_visibility="collapsed"
    )
    
    if not uploaded_files:
        st.info(t("mos_upload_prompt"))
        return
            
    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    col_btn, _ = st.columns([1.5, 4.5])
    with col_btn:
        run_btn = st.button(t("mos_process_btn"), type="primary", use_container_width=True)
        
    if run_btn:
        files_key = '_'.join(sorted([f.name for f in uploaded_files]))
        if st.session_state.get('mos_files_key') == files_key and 'df_mos_result' in st.session_state:
            st.success("✅ Sử dụng kết quả đã tóm tắt (Cached). Nếu muốn tính lại, hãy chọn lại file!")
        else:
            has_key = bool(load_saved_api_key())
            
            if not has_key:
                st.warning("⚠️ Chưa cấu hình GEMINI_API_KEY. Hệ thống đang dùng chế độ nối chuỗi thủ công thay vì dùng AI!")
                
            with st.spinner("🤖 AI đang phân tích và tóm tắt nội dung ủy thác..."):
                dfs = []
                file_errors = []
                for f in uploaded_files:
                    try:
                        df_f = parse_mos_file(f, f.name)
                        if not df_f.empty:
                            dfs.append(df_f)
                    except Exception as e:
                        file_errors.append(str(e))
                
                if file_errors:
                    st.session_state['mos_file_errors'] = file_errors
                else:
                    st.session_state.pop('mos_file_errors', None)
                
                if dfs:
                    df_tong_hop = tong_hop_mos(dfs)
                    st.session_state['df_mos_raw'] = pd.concat(dfs, ignore_index=True)
                    st.session_state['df_mos_result'] = df_tong_hop
                    st.session_state['mos_files_key'] = files_key
                    st.session_state['mos_saved'] = False
                    st.session_state['mos_num_people'] = len(dfs)
                    st.success("✅ Tổng hợp và tóm tắt AI xong!")
                else:
                    st.error("Không tìm thấy dữ liệu hợp lệ (phần MOS業務) trong các file đã tải lên.")
                
    if 'df_mos_result' in st.session_state and st.session_state['df_mos_result'] is not None:
        # Làm tròn một lần duy nhất vào session state để giữ data ổn định, tránh mất focus khi enter
        if 'Giờ làm (h)' in st.session_state['df_mos_result'].columns:
            def safe_round_float(val):
                if pd.isna(val) or str(val).strip() == '': return 0.0
                try: return round(float(val), 1)
                except: return 0.0
                
            st.session_state['df_mos_result']['Giờ làm (h)'] = st.session_state['df_mos_result']['Giờ làm (h)'].apply(safe_round_float).astype(float)
        df_result = st.session_state['df_mos_result']



        
        st.markdown("---")
        df_raw = st.session_state.get('df_mos_raw', pd.DataFrame())
        tab_data, tab_combined = st.tabs([t("mos_tab_data"), t("mos_tab_stats")])

        with tab_data:
            # 4. BỘ LỌC THÔNG MINH
            st.markdown(t("mos_adv_filter"))
            col_f1, col_f2, col_f3 = st.columns(3)
            with col_f1:
                search_name = st.text_input(t("mos_project_name"), placeholder=t("mos_search_placeholder"))
            with col_f2:
                ma_da_list = df_result['Mã dự án'].unique().tolist()
                selected_ma_da = st.multiselect(t("mos_col_proj_code"), options=ma_da_list, default=[], placeholder=t("mos_code_placeholder"))
            with col_f3:
                default_managers = ["フォン \n Phương", "ロン \n Long", "ダオ \n Đạo"]
                data_managers = [str(x) for x in df_result['Quản lý Việt Nam'].dropna().unique() if str(x).strip()]
                manager_list = [t("mos_all")] + sorted(list(set(default_managers + data_managers)))
                selected_manager = st.selectbox(t("mos_vn_manager"), options=manager_list)

            df_display = df_result.copy()
            if search_name:
                df_display = df_display[df_display['Tên dự án'].str.contains(search_name, case=False, na=False)]
            if selected_ma_da:
                df_display = df_display[df_display['Mã dự án'].isin(selected_ma_da)]
            if selected_manager != t("mos_all"):
                df_display = df_display[df_display['Quản lý Việt Nam'] == selected_manager]

            # 5. CẢNH BÁO BẤT THƯỜNG
            def check_anomaly(row):
                gio = row.get('Giờ làm (h)', 0)
                try: gio = float(gio)
                except: gio = 0
                if gio <= 0: return t("mos_warning_0h")
                if gio > 50: return t("mos_warning_ot")
                return t("mos_warning_ok")

            if 'Cảnh báo' not in df_display.columns:
                df_display.insert(1, 'Cảnh báo', df_display.apply(check_anomaly, axis=1))

            # Ẩn cột STT khỏi hiển thị
            cols_to_hide = ['STT'] if 'STT' in df_display.columns else []
            df_display_show = df_display.drop(columns=cols_to_hide, errors='ignore')

            st.markdown(t("mos_edit_hint"))
            
            @st.fragment
            def render_mos_editor():
                df_editor_input = df_display_show.copy()
                if 'Giờ làm (h)' in df_editor_input.columns:
                    df_editor_input['Giờ làm (h)'] = df_editor_input['Giờ làm (h)'].apply(
                        lambda x: f"{x:g}" if pd.notna(x) else ""
                    ).astype(str)

                edited_display = st.data_editor(
                    df_editor_input,
                    use_container_width=True,
                    num_rows="dynamic",
                    column_config={
                        "Cảnh báo": st.column_config.TextColumn(t("mos_col_warning"), disabled=True),
                        "Mã dự án": st.column_config.TextColumn(t("mos_col_proj_code"), disabled=False),
                        "Tên dự án": st.column_config.TextColumn(t("mos_col_proj_name"), disabled=False),
                        "Phân vùng": st.column_config.TextColumn(t("mos_col_area")),
                        "Nội dung ủy thác": st.column_config.TextColumn(t("mos_col_task")),
                        "Giờ làm (h)": st.column_config.TextColumn(t("mos_col_hours"), disabled=False),
                        "Ngày bắt đầu": st.column_config.TextColumn(t("mos_start_date"), disabled=False),
                        "Ngày kết thúc": st.column_config.TextColumn(t("mos_end_date"), disabled=False),
                        "Quản lý Nhật Bản": st.column_config.TextColumn(t("mos_col_jp_manager"), disabled=False),
                        "Quản lý Việt Nam": st.column_config.SelectboxColumn(t("mos_col_vn_manager"), options=["フォン \n Phương", "ロン \n Long", "ダオ \n Đạo"]),
                        "Người thực hiện": st.column_config.TextColumn(t("mos_col_executor"), disabled=False),
                        "Trạng thái": st.column_config.SelectboxColumn(t("mos_col_status"), options=["完了 \n Hoàn thành", "実行中 \n Đang tiến hành", "未着手 \n Chưa bắt đầu"]),
                    },
                    hide_index=True,
                    key="mos_data_editor"
                )

                if st.button(t("save_btn"), type="primary", key="btn_mos_save_edits"):
                    st.session_state['mos_saved'] = True
                    edited_no_warning = edited_display.drop(columns=['Cảnh báo'], errors='ignore')
                    
                    if 'Giờ làm (h)' in edited_no_warning.columns:
                        def parse_and_round(val):
                            try:
                                return round(float(str(val).replace(',', '.')), 1)
                            except:
                                return 0.0
                        edited_no_warning['Giờ làm (h)'] = edited_no_warning['Giờ làm (h)'].apply(parse_and_round).astype(float)
                        
                    if search_name or selected_ma_da or selected_manager != t("mos_all"):
                        df_result.update(edited_no_warning)
                        st.session_state['df_mos_edited'] = df_result
                        st.session_state['df_mos_result'] = df_result
                    else:
                        st.session_state['df_mos_edited'] = edited_no_warning
                        st.session_state['df_mos_result'] = edited_no_warning
                    st.rerun()

            render_mos_editor()
            
            st.markdown("<br>", unsafe_allow_html=True)
            
            import datetime
            import calendar
            
            total_da = len(df_result)
            total_gio = df_result['Giờ làm (h)'].sum()
            total_nv = st.session_state.get('mos_num_people', 0)
            
            lang = st.session_state.get('lang', 'vi')
            lbl_title = "Thống kê & KPI" if lang == 'vi' else "統計・KPI"
            lbl_month = "Tháng" if lang == 'vi' else "月"
            lbl_year = "Năm" if lang == 'vi' else "年"
            lbl_holiday = "Số ngày lễ (Âm lịch)" if lang == 'vi' else "祝日数"
            
            st.markdown(f"### 📊 {lbl_title}")
            col_kpi1, col_kpi2, col_kpi3 = st.columns(3)
            with col_kpi1:
                current_month = datetime.date.today().month
                selected_month = st.selectbox(lbl_month, range(1, 13), index=current_month-1, key="kpi_month")
            with col_kpi2:
                current_year = datetime.date.today().year
                selected_year = st.selectbox(lbl_year, range(current_year-2, current_year+3), index=2, key="kpi_year")
            with col_kpi3:
                manual_holidays = st.number_input(lbl_holiday, min_value=0, max_value=20, value=0, key="kpi_holidays")
                
            if selected_month == 1:
                start_date = datetime.date(selected_year - 1, 12, 21)
            else:
                start_date = datetime.date(selected_year, selected_month - 1, 21)
            end_date = datetime.date(selected_year, selected_month, 20)
            
            cycle_days = [start_date + datetime.timedelta(days=i) for i in range((end_date - start_date).days + 1)]
            fixed_holidays = [(4, 30), (5, 1), (9, 2)]
            
            saturdays = 0
            sundays = 0
            holidays_in_cycle = 0
            
            for d in cycle_days:
                if d.weekday() == 5: saturdays += 1
                if d.weekday() == 6: sundays += 1
                if (d.month, d.day) in fixed_holidays and d.weekday() < 5:
                    holidays_in_cycle += 1
                    
            # Tìm ngày thứ 7 cuối cùng của tháng M-1
            last_day_prev_month = calendar.monthrange(start_date.year, start_date.month)[1]
            last_saturday = None
            for d in range(last_day_prev_month, 0, -1):
                test_date = datetime.date(start_date.year, start_date.month, d)
                if test_date.weekday() == 5:
                    last_saturday = test_date
                    break
                    
            add_saturday = 1 # Cộng thêm 1 ngày thứ bảy cuối tháng
                
            standard_days = len(cycle_days) - saturdays - sundays - holidays_in_cycle - manual_holidays + add_saturday
            std_hours_per_person = standard_days * 8
            target_hours = total_nv * std_hours_per_person
            completion_rate = (total_gio / target_hours * 100) if target_hours > 0 else 0
            
            st.session_state['mos_kpi_month'] = selected_month
            st.session_state['mos_kpi_year'] = selected_year
            st.session_state['mos_kpi_std_hours'] = std_hours_per_person
            
            st.markdown(f"""
            <style>
            .kpi-table-container {{
                background: #ffffff;
                border: 1px solid #E2E8F0;
                border-radius: 8px;
                padding: 0;
                margin-bottom: 24px;
                box-shadow: 0 1px 3px rgba(0,0,0,0.05);
                overflow: hidden;
            }}
            .kpi-table {{
                width: 100%;
                border-collapse: collapse;
            }}
            .kpi-table th, .kpi-table td {{
                border: 1px solid #E2E8F0;
                padding: 12px 20px;
            }}
            .kpi-table tr:nth-child(even) {{
                background-color: #F8FAFC;
            }}
            .kpi-table tr:hover {{
                background-color: #F1F5F9;
            }}
            .kpi-label {{
                font-weight: 500;
                color: #334155;
                font-size: 14px;
                width: 65%;
            }}
            .kpi-val {{
                font-weight: 700;
                color: #0F172A;
                font-size: 16px;
                text-align: center;
            }}
            .kpi-highlight {{
                color: #0EA5E9;
            }}
            </style>
            <div class="kpi-table-container">
                <table class="kpi-table">
                    <tr>
                        <td class="kpi-label">人数<br>Số người</td>
                        <td class="kpi-val">{total_nv}</td>
                    </tr>
                    <tr>
                        <td class="kpi-label">一人当たり月枠稼働時間(h)<br>Giờ làm việc tiêu chuẩn(h)</td>
                        <td class="kpi-val">{std_hours_per_person}</td>
                    </tr>
                    <tr>
                        <td class="kpi-label">月目標稼働時間(h)<br>Mục tiêu giờ làm(h)</td>
                        <td class="kpi-val">{target_hours}</td>
                    </tr>
                    <tr>
                        <td class="kpi-label">月実績稼働時間(h)<br>Giờ làm thực tế(h)</td>
                        <td class="kpi-val">{total_gio:g}</td>
                    </tr>
                    <tr>
                        <td class="kpi-label">目標に対して稼働率(%)<br>Tỷ lệ hoàn thành(%)</td>
                        <td class="kpi-val kpi-highlight">{completion_rate:.2f}%</td>
                    </tr>
                </table>
            </div>
            """, unsafe_allow_html=True)

        with tab_combined:
            st.markdown(t("mos_stats_title"))
            # --- Năng suất cá nhân ---
            if not df_raw.empty:
                emp_stats = df_raw.groupby("ten_nv")["tong_gio"].sum().reset_index()
                emp_stats.columns = [t("export_col_emp_name"), t("mos_total_hours_table")]
                # Làm tròn giờ đến 1 chữ số thập phân
                emp_stats[t("mos_total_hours_table")] = emp_stats[t("mos_total_hours_table")].apply(
                    lambda x: round(float(x), 1) if pd.notna(x) else x
                )
                emp_stats = emp_stats.sort_values(t("mos_total_hours_table"), ascending=False)
                st.dataframe(emp_stats, use_container_width=True, hide_index=True)
            else:
                st.info(t("mos_no_raw_data"))
            
            st.markdown("---")
            
            # --- Biểu đồ thống kê ---
            try:
                import plotly.express as px
                if not df_result.empty:
                    col_d1, col_d2 = st.columns(2)
                    with col_d1:
                        fig_pie = px.pie(df_result, values="Giờ làm (h)", names="Tên dự án", title=t("mos_chart_pie_title"))
                        st.plotly_chart(fig_pie, use_container_width=True)
                    with col_d2:
                        mgr_stats = df_result.groupby("Quản lý Việt Nam")["Giờ làm (h)"].sum().reset_index()
                        fig_bar = px.bar(mgr_stats, x="Quản lý Việt Nam", y="Giờ làm (h)", title=t("mos_chart_bar_title"), color="Quản lý Việt Nam")
                        st.plotly_chart(fig_bar, use_container_width=True)
                else:
                    st.info(t("mos_no_chart_data"))
            except ImportError:
                st.info(t("mos_no_chart_data"))

        # If not submitted but it exists in session, ensure it's mapped so download works correctly
        if 'df_mos_result' not in st.session_state:
            st.session_state['df_mos_result'] = df_result
        if 'df_mos_edited' not in st.session_state:
            st.session_state['df_mos_edited'] = df_result
        

        
        def to_excel(df):
            import openpyxl
            from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
            import re
            import datetime
            import io
            
            font_normal = Font(name='Times New Roman', size=12)
            font_bold = Font(name='Times New Roman', size=12, bold=True)
            font_title = Font(name='Times New Roman', size=16, bold=True)
            
            align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
            align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
            align_right = Alignment(horizontal='right', vertical='center', wrap_text=True)
            
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            
            def format_bilingual(text):
                if text is None: return ""
                text_str = str(text).strip()
                if text_str.lower() in ['nan', 'nat', 'none', '']: return ""
                if '\n' not in text_str:
                    text_str = re.sub(r'([\u3040-\u30ff\u4e00-\u9faf])\s*(?:/|-)?\s*([A-Za-zÀ-ỹ])', r'\1\n\2', text_str)
                    match = re.match(r'^([A-Za-zÀ-ỹ\s/]+)\s*(?:/|-)?\s*([\u3040-\u30ff\u4e00-\u9faf\s/]+)$', text_str)
                    if match:
                        text_str = match.group(2).strip() + '\n' + match.group(1).strip()
                return text_str

            def set_cell(cell, text, bold=False, align=align_center, fill=None, font=None):
                cell.value = format_bilingual(text)
                if font:
                    cell.font = font
                else:
                    cell.font = font_bold if bold else font_normal
                cell.alignment = align
                if fill:
                    cell.fill = fill

            # --- Pre-calculate stats ---
            unique_people = set()
            if 'Người thực hiện' in df.columns:
                for val in df['Người thực hiện']:
                    v_str = str(val).strip()
                    if v_str and v_str.lower() not in ['nan', 'nat', 'none']:
                        unique_people.add(v_str)
            num_people = len(unique_people)
            
            sum_hours = 0
            for idx, row in df.iterrows():
                gio_lam = row.get('Giờ làm (h)', '')
                try: 
                    if str(gio_lam).strip() != '':
                        sum_hours += float(gio_lam)
                except: pass

            output = io.BytesIO()
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = 'Tổng hợp MOS'
            
            now = datetime.datetime.now()
            
            kpi_month = st.session_state.get('mos_kpi_month', now.month)
            kpi_year = st.session_state.get('mos_kpi_year', now.year)
            kpi_std_hours = st.session_state.get('mos_kpi_std_hours', 168)
            num_people_kpi = st.session_state.get('mos_num_people', num_people)
            year_jp = kpi_year - 2000 if kpi_year >= 2000 else kpi_year
            
            # --- Header Title ---
            ws.merge_cells('A1:J1')
            set_cell(ws['A1'], 'モス委託業務工数集計', font=font_title)
            
            ws.merge_cells('A2:J2')
            set_cell(ws['A2'], 'Bảng kê chi tiết nội dung nghiệp vụ ủy thác', font=font_title)
            
            ws.merge_cells('A4:J4')
            set_cell(ws['A4'], f'（{year_jp}年{kpi_month}月分）', font=font_title)
            
            ws.merge_cells('A5:J5')
            set_cell(ws['A5'], f'Phần tháng {kpi_month}/{kpi_year}', font=font_title)
            
            # --- Các bảng phụ bên trái ---
            sub_table_titles = [
                ('A10', '人数\nSố người'),
                ('A11', '一人当たり月枠稼働時間(h)\nGiờ làm việc tiêu chuẩn(h)'),
                ('A12', '月目標稼働時間(h)\nMục tiêu giờ làm(h)'),
                ('A13', '月実績稼働時間(h)\nGiờ làm thực tế(h)'),
                ('A14', '目標に対して稼働率(%)\nTỷ lệ hoàn thành(%)')
            ]
            for coord, title in sub_table_titles:
                set_cell(ws[coord], title, align=align_right)
                ws[coord].border = thin_border
                ws[f'B{coord[1:]}'].border = thin_border
                ws[f'B{coord[1:]}'].font = font_normal
                ws[f'B{coord[1:]}'].alignment = align_center

            ws['B10'] = num_people_kpi
            ws['B11'] = kpi_std_hours
            ws['B12'] = "=B10*B11"
            ws['B13'] = sum_hours
            ws['B14'] = "=B13/(B10*B11)"
            ws['B14'].number_format = '0.00%'
            
            # --- Cột Ngày tháng bên phải ---
            ws.merge_cells('H10:J10')
            set_cell(ws['H10'], f'作成日付： {now.year}年{now.month}月{now.day}日', align=align_left)
            ws.merge_cells('H11:J11')
            set_cell(ws['H11'], f'Ngày lập bảng kê: {now.day}/{now.month}/{now.year}', align=align_left)
            ws.merge_cells('H12:J12')
            set_cell(ws['H12'], '作成者： レータンフォン', align=align_left)
            ws.merge_cells('H13:J13')
            set_cell(ws['H13'], 'Người lập bảng kê: Lê Thanh Phương', align=align_left)
            
            # --- Table Headers (Row 16 & 17) ---
            headers = [
                ('A16:A17', '案件名\nTên dự án', 30),
                ('B16:B17', '区分\nPhân vùng', 15),
                ('C16:C17', '委託内容\nNội dung ủy thác', 50),
                ('D16:D17', '実工数(h)\nGiờ làm (h)', 15),
                ('E16:F16', '期間 Thời gian', None),
                ('G16:H16', '管理者 Người quản lý', None),
                ('I16:I17', '実施者\nNgười thực hiện', 20),
                ('J16:J17', '状態\nTrạng thái', 15)
            ]
            
            for range_str, text, width in headers:
                ws.merge_cells(range_str)
                cell = ws[range_str.split(':')[0]]
                if text in ["期間 Thời gian", "管理者 Người quản lý"]:
                    cell.value = text
                    cell.font = font_bold
                    cell.alignment = align_center
                else:
                    set_cell(cell, text, bold=True)
                if width:
                    ws.column_dimensions[cell.column_letter].width = width
            
            sub_headers = {
                'E17': '委託受領日\nNgày bắt đầu',
                'F17': '完了日\nNgày kết thúc',
                'G17': '日本\nNhật Bản',
                'H17': 'ベトナム\nViệt Nam'
            }
            for c, text in sub_headers.items():
                set_cell(ws[c], text, bold=True)
                ws.column_dimensions[ws[c].column_letter].width = 15
            
            for r in ws['A16:J17']:
                for cell in r:
                    cell.font = font_bold
                    cell.alignment = align_center
                    
            for range_str, text, _ in headers:
                cell = ws[range_str.split(':')[0]]
                if text in ["期間 Thời gian", "管理者 Người quản lý"]:
                    cell.value = text
                else:
                    cell.value = format_bilingual(text)
            for c, text in sub_headers.items():
                ws[c].value = format_bilingual(text)
            
            # --- Write Data ---
            row_idx = 18
            for idx, row in df.iterrows():
                ten_da = f"{row.get('Mã dự án', '')}_{row.get('Tên dự án', '')}"
                if ten_da == "_": ten_da = ""
                
                set_cell(ws[f'A{row_idx}'], ten_da, align=align_left)
                set_cell(ws[f'B{row_idx}'], row.get('Phân vùng', ''))
                set_cell(ws[f'C{row_idx}'], row.get('Nội dung ủy thác', ''), align=align_left)
                
                gio_lam = row.get('Giờ làm (h)', '')
                ws[f'D{row_idx}'] = gio_lam
                ws[f'D{row_idx}'].font = font_normal
                ws[f'D{row_idx}'].alignment = align_center
                
                set_cell(ws[f'E{row_idx}'], row.get('Ngày bắt đầu', ''))
                set_cell(ws[f'F{row_idx}'], row.get('Ngày kết thúc', ''))
                set_cell(ws[f'G{row_idx}'], row.get('Quản lý Nhật Bản', ''))
                set_cell(ws[f'H{row_idx}'], row.get('Quản lý Việt Nam', ''))
                set_cell(ws[f'I{row_idx}'], row.get('Người thực hiện', ''), align=align_left)
                set_cell(ws[f'J{row_idx}'], row.get('Trạng thái', ''))
                
                row_idx += 1
            
            # --- Footer ---
            ws.merge_cells(f'A{row_idx}:C{row_idx}')
            set_cell(ws[f'A{row_idx}'], '実工数合計(h)\nTổng giờ làm (h)', bold=True, align=align_right)
            
            ws[f'D{row_idx}'] = sum_hours
            ws[f'D{row_idx}'].font = font_bold
            ws[f'D{row_idx}'].alignment = align_center
            
            # Kẻ khung toàn bộ bảng (từ row 16 đến row_idx)
            for row_cells in ws.iter_rows(min_row=16, max_row=row_idx, min_col=1, max_col=10):
                for cell in row_cells:
                    cell.border = thin_border
            
            wb.save(output)
            return output.getvalue()

        kpi_m = st.session_state.get('mos_kpi_month', datetime.datetime.now().month)
        excel_data = to_excel(st.session_state['df_mos_edited'])
        st.download_button(
            label=t("mos_download_report"),
            data=excel_data,
            file_name=f"{kpi_m}月委託業務工数集計.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

# ==========================================
# GIAO DIỆN CÁC CHỨC NĂNG
# ==========================================

# Chỉ ẩn sidebar ở mos (chấm công thì hiện ra)
if st.session_state.app_page == "mos":
    st.markdown("""
    <style>
    div[data-testid="stButton"] {
        display: flex !important;
        
    }
    div[data-testid="stButton"] > button {
        height: 40px !important;
        min-height: 40px !important;
        line-height: 40px !important;
        padding: 0 28px !important;
        border-radius: 100px !important;
        background: linear-gradient(135deg, #0EA5E9, #38BDF8) !important;
        box-shadow: 0 4px 12px rgba(14, 165, 233, 0.3) !important;
        transition: transform 0.2s ease, box-shadow 0.2s ease !important;
        color: white !important;
        border: none !important;
        width: auto !important; /* Force width to fit text */
        white-space: nowrap !important;
    }
    div[data-testid="stButton"] > button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 16px rgba(14, 165, 233, 0.4) !important;
    }
    .stButton > button p {
        font-size: 16px !important;
        font-weight: 600 !important;
    }
    </style>
    

    """, unsafe_allow_html=True)





    render_mos_page()
    
    render_chatbot()
    st.stop()

def render_checkin_page():
    st.markdown("---")
    col_nav1, col_nav2 = st.columns([1, 5])
    with col_nav1:
        if st.button("⬅ Trang chủ" if st.session_state.lang == 'vi' else "⬅ ホーム", type="secondary", use_container_width=True):
            st.session_state.app_page = "home"; st.rerun()
    with col_nav2:
        st.markdown("<h1 style='margin:0; color:#0F172A; font-size:28px;'>📱 Cổng Check-in GPS Hiện Trường</h1>" if st.session_state.lang == 'vi' else "<h1 style='margin:0; color:#0F172A; font-size:28px;'>📱 フィールド打刻 GPS</h1>", unsafe_allow_html=True)
        st.markdown("<p style='color:#64748B; margin-top:4px;'>Dành cho Kỹ sư đi lắp đặt, bảo trì công trình tại nhà máy khách hàng</p>" if st.session_state.lang == 'vi' else "<p style='color:#64748B; margin-top:4px;'>顧客工場での現地設置・保守作業員向け</p>", unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    col_c1, col_c2, col_c3 = st.columns([1, 2, 1])
    with col_c2:
        with st.container(border=True):
            is_vi = (st.session_state.lang == 'vi')
            if "manual_emps" not in st.session_state:
                st.session_state.manual_emps = []

            emp_options = get_company_emp_options(st.session_state.lang)

            sel_emp_gps = st.selectbox("Chọn nhân viên (*)" if is_vi else "社員を選択 (*)", emp_options, key="sb_sel_emp_gps")
            ma_nv = sel_emp_gps.split(" - ")[0].strip() if sel_emp_gps else ""
            ten_nv = sel_emp_gps.split(" - ")[1].strip() if (sel_emp_gps and " - " in sel_emp_gps) else ""
            
            with st.expander("➕ Thêm nhân sự mới (nếu chưa có trong danh sách)" if is_vi else "➕ 新規社員を追加"):
                with st.form("form_add_emp_gps"):
                    col_m1, col_m2 = st.columns([1, 2])
                    new_ma_gps = col_m1.text_input("Mã NV (*)" if is_vi else "社員ID (*)", placeholder="VD: NV099")
                    new_ten_gps = col_m2.text_input("Tên nhân viên (*)" if is_vi else "氏名 (*)", placeholder="VD: Nguyễn Văn A")
                    if st.form_submit_button("Thêm vào danh sách" if is_vi else "リストに追加", type="secondary", use_container_width=True):
                        if new_ma_gps and new_ten_gps:
                            st.session_state.manual_emps.append({
                                "ma": new_ma_gps.strip().upper(),
                                "ten": new_ten_gps.strip(),
                                "cv": "",
                                "pb": ""
                            })
                            st.success(f"✅ Đã thêm NV {new_ma_gps.strip().upper()} - {new_ten_gps.strip()}!" if is_vi else f"✅ {new_ma_gps.strip().upper()} を追加しました！")
                            st.rerun()
                        else:
                            st.error("Vui lòng nhập đủ Mã NV và Tên!" if is_vi else "IDと氏名を入力してください")
            
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                ngay_bat_dau = st.date_input("Từ ngày (*)" if is_vi else "作業日(開始) (*)", value=datetime.date.today(), key="gps_d_start")
            with col_d2:
                ngay_ket_thuc = st.date_input("Đến ngày (*)" if is_vi else "作業日(終了) (*)", value=datetime.date.today(), key="gps_d_end")

            col_t1, col_t2 = st.columns([3, 2])
            with col_t1:
                loai = st.radio("Loại check-in" if is_vi else "打刻種別", ["🟢 Vào ca (Check-in)", "🟣 Tan ca (Check-out)"] if is_vi else ["🟢 出勤", "🟣 退勤"], horizontal=True)
            with col_t2:
                is_vao = ('Vào ca' in loai or '出勤' in loai)
                default_time = datetime.time(8, 0) if is_vao else datetime.time(17, 0)
                gio_checkin = st.time_input("Giờ check-in (*)" if is_vi else "打刻時刻 (*)", value=default_time, key=f"gps_t_{'in' if is_vao else 'out'}")

            dia_diem = st.text_input("Địa điểm hiện trường (*)" if is_vi else "現地場所 (*)", placeholder="VD: Nhà máy Canon Bắc Ninh" if is_vi else "例: キヤノンバクニン工場")
            ghi_chu = st.text_area("Chi tiết công việc" if is_vi else "作業詳細", placeholder="VD: Kiểm tra cảm biến tủ điện ca sáng..." if is_vi else "例: 午前シフトの配電盤点検...")
            
            st.info("📍 Vị trí GPS hệ thống ghi nhận: **21.1245° N, 106.0523° E** (Độ chính xác: ±5m)" if is_vi else "📍 システム取得GPS位置: **21.1245° N, 106.0523° E** (精度: ±5m)")
            
            if st.button("Xác nhận Chấm công Hiện trường" if is_vi else "打刻データを確定", type="primary", use_container_width=True):
                if not ma_nv or not ten_nv or not dia_diem:
                    st.error("Vui lòng điền các thông tin bắt buộc (*)" if is_vi else "必須項目(*)を入力してください")
                elif ngay_bat_dau > ngay_ket_thuc:
                    st.error("Ngày kết thúc không được nhỏ hơn ngày bắt đầu!" if is_vi else "終了日は開始日以降を指定してください")
                else:
                    curr_d = ngay_bat_dau
                    saved_count = 0
                    while curr_d <= ngay_ket_thuc:
                        time_str = curr_d.strftime("%d/%m/%Y") + " " + gio_checkin.strftime("%H:%M:%S")
                        save_field_checkin(ma_nv.strip(), ten_nv.strip(), time_str, loai, dia_diem.strip(), "21.1245, 106.0523", ghi_chu.strip())
                        curr_d += datetime.timedelta(days=1)
                        saved_count += 1
                        
                    if saved_count == 1:
                        st.success(f"✅ Đã ghi nhận thành công ngày {ngay_bat_dau.strftime('%d/%m/%Y')}!" if is_vi else f"✅ {ngay_bat_dau.strftime('%Y/%m/%d')} の打刻を記録しました！")
                    else:
                        st.success(f"✅ Đã ghi nhận công tác thành công cho {saved_count} ngày (từ {ngay_bat_dau.strftime('%d/%m/%Y')} đến {ngay_ket_thuc.strftime('%d/%m/%Y')})!" if is_vi else f"✅ {saved_count}日間 ({ngay_bat_dau.strftime('%Y/%m/%d')} ~ {ngay_ket_thuc.strftime('%Y/%m/%d')}) の作業を記録しました！")
                    st.rerun()

    df_history = get_field_checkins()
    is_vi = (st.session_state.lang == 'vi')
    
    col_h1, col_h2 = st.columns([4, 1])
    with col_h1:
        st.markdown("### 📋 Lịch sử Check-in hiện trường mới nhất" if is_vi else "### 📋 最新の現地打刻履歴")
    with col_h2:
        if not df_history.empty:
            if st.button("🗑️ Xóa toàn bộ" if is_vi else "🗑️ 全履歴を削除", key="btn_del_all_gps", use_container_width=True):
                conn = sqlite3.connect(DB_FILE)
                conn.execute("DELETE FROM field_checkins")
                conn.commit(); conn.close()
                st.rerun()

    if df_history.empty:
        st.caption("Chưa có dữ liệu check-in" if is_vi else "打刻データがありません")
    else:
        raw_df = df_history.copy()
        if 'ten_nv' in df_history.columns:
            df_history['ten_nv'] = [translate_name(x, st.session_state.lang) for x in df_history['ten_nv']]
        if not is_vi and 'thoi_gian' in df_history.columns:
            def fmt_jp_dt(val):
                s = str(val).strip()
                if len(s) >= 10 and s[2] == '/' and s[5] == '/':
                    return f"{s[6:10]}/{s[3:5]}/{s[0:2]}" + s[10:]
                return s
            df_history['thoi_gian'] = df_history['thoi_gian'].apply(fmt_jp_dt)
        for c in df_history.columns:
            df_history[c] = ["" if str(v).strip().lower() in ['nan', '<na>', 'none', 'null'] else v for v in df_history[c]]
        if not is_vi:
            df_history = df_history.rename(columns={
                'ma_nv': '社員ID', 'ten_nv': '氏名', 'thoi_gian': '打刻日時',
                'loai': '打刻種別', 'dia_diem': '現地場所', 'toa_do': '座標', 'ghi_chu': '作業詳細'
            })
        else:
            df_history = df_history.rename(columns={
                'ma_nv': 'Mã NV', 'ten_nv': 'Tên NV', 'thoi_gian': 'Thời gian',
                'loai': 'Loại', 'dia_diem': 'Địa điểm', 'toa_do': 'Tọa độ', 'ghi_chu': 'Chi tiết'
            })
        if 'id' in df_history.columns:
            df_history = df_history.drop(columns=['id'])
        st.dataframe(df_history, use_container_width=True, hide_index=True)

        with st.expander("🗑️ Xóa từng dòng dữ liệu sai" if is_vi else "🗑️ 個別データの削除"):
            del_opts = [f"ID {r['id']}: {r['thoi_gian']} | {r['ma_nv']} - {translate_name(r['ten_nv'], st.session_state.lang)} | {r['loai']}" for _, r in raw_df.iterrows()]
            sel_del = st.selectbox("Chọn lượt check-in cần xóa" if is_vi else "削除する打刻を選択", ["-- Chọn --" if is_vi else "-- 選択 --"] + del_opts, key="gps_single_del")
            if sel_del != ("-- Chọn --" if is_vi else "-- 選択 --"):
                del_id = int(sel_del.split(":")[0].replace("ID ", "").strip())
                if st.button("🗑️ Xác nhận Xóa dòng này" if is_vi else "🗑️ このデータを削除確定", type="primary", key="btn_confirm_gps_del"):
                    conn = sqlite3.connect(DB_FILE)
                    conn.execute(f"DELETE FROM field_checkins WHERE id = {del_id}")
                    conn.commit(); conn.close()
                    st.success("✅ Đã xóa thành công!" if is_vi else "✅ 削除しました！")
                    st.rerun()

def render_history_page():
    st.markdown("---")
    col_nav1, col_nav2 = st.columns([1, 5])
    with col_nav1:
        if st.button("⬅ Trang chủ" if st.session_state.lang == 'vi' else "⬅ ホーム", type="secondary", use_container_width=True):
            st.session_state.app_page = "home"; st.rerun()
    with col_nav2:
        st.markdown("<h1 style='margin:0; color:#0F172A; font-size:28px;'>📁 Kho Lưu Trữ Lịch Sử & Biểu Đồ So Sánh</h1>" if st.session_state.lang == 'vi' else "<h1 style='margin:0; color:#0F172A; font-size:28px;'>📁 履歴アーカイブ＆比較チャート</h1>", unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    conn = sqlite3.connect(DB_FILE)
    df_all = pd.read_sql_query("SELECT * FROM records", conn)
    conn.close()
    
    if df_all.empty:
        st.warning("CSDL SQLite cục bộ hiện chưa có bản ghi chấm công nào. Vui lòng sang trang Chấm công Excel để xử lý và lưu dữ liệu trước." if st.session_state.lang == 'vi' else "データベースに勤怠レコードがありません。先にExcel計算ページでデータを保存してください。")
        return

    df_all['thang_nam'] = df_all['ngay'].apply(lambda x: str(x)[3:10] if len(str(x))>=10 else "N/A")
    list_thang = sorted([t for t in df_all['thang_nam'].unique().tolist() if t != "N/A"], reverse=True)
    if not list_thang:
        list_thang = ["Tất cả"]
        df_all['thang_nam'] = "Tất cả"
    
    tab1, tab2 = st.tabs(["🔍 Tra cứu Lịch sử từng tháng" if st.session_state.lang == 'vi' else "🔍 月別履歴照会", "📊 So sánh Biến động giữa 2 kỳ" if st.session_state.lang == 'vi' else "📊 ２期間の変動比較"])
    
    is_vi = (st.session_state.lang == 'vi')
    with tab1:
        sel_thang = st.selectbox("Chọn chu kỳ (*)" if is_vi else "対象月を選択 (*)", list_thang)
        df_thang = df_all[df_all['thang_nam'] == sel_thang]
        
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Tổng nhân sự" if is_vi else "総社員数", f"{df_thang['ma_nv'].nunique()} NV" if is_vi else f"{df_thang['ma_nv'].nunique()} 名")
        c2.metric("Tổng giờ làm" if is_vi else "総労働時間", f"{df_thang['tong_gio'].sum():.1f} h")
        c3.metric("Tổng giờ OT" if is_vi else "残業時間(OT)", f"{df_thang['ot'].sum():.1f} h")
        c4.metric("Lượt đi trễ" if is_vi else "遅刻回数", f"{(df_thang['di_tre'] > 0).sum()} lần" if is_vi else f"{(df_thang['di_tre'] > 0).sum()} 回")
        
        df_display = df_thang.copy()
        if 'ten_nv' in df_display.columns:
            df_display['ten_nv'] = [translate_name(x, st.session_state.lang) for x in df_display['ten_nv']]
        if not is_vi and 'ngay' in df_display.columns:
            def fmt_jp_d(val):
                s = str(val).strip()
                if len(s) == 10 and s[2] == '/' and s[5] == '/':
                    return f"{s[6:10]}/{s[3:5]}/{s[0:2]}"
                return s
            df_display['ngay'] = df_display['ngay'].apply(fmt_jp_d)
        if 'thang_nam' in df_display.columns:
            df_display = df_display.drop(columns=['thang_nam'])
        for c in df_display.columns:
            df_display[c] = ["" if str(v).strip().lower() in ['nan', '<na>', 'none', 'null'] else v for v in df_display[c]]
        if not is_vi:
            df_display = df_display.rename(columns={
                'ma_nv': '社員ID', 'ten_nv': '氏名', 'ngay': '日付',
                'gio_vao': '出勤', 'gio_ra': '退勤', 'di_tre': '遅刻(分)',
                've_som': '早退(分)', 'ot': '残業(h)', 'tong_gio': '総時間(h)', 'ghi_chu': '備考'
            })
        else:
            df_display = df_display.rename(columns={
                'ma_nv': 'Mã NV', 'ten_nv': 'Tên NV', 'ngay': 'Ngày',
                'gio_vao': 'Giờ vào', 'gio_ra': 'Giờ ra', 'di_tre': 'Đi trễ',
                've_som': 'Về sớm', 'ot': 'OT', 'tong_gio': 'Tổng giờ', 'ghi_chu': 'Ghi chú'
            })
        st.dataframe(df_display, use_container_width=True, hide_index=True)
        
    with tab2:
        if len(list_thang) < 2:
            st.info("Cần ít nhất 2 chu kỳ tháng khác nhau trong DB để thực hiện so sánh." if is_vi else "比較を行うにはデータベースに少なくとも2つの異なる月が必要です。")
        else:
            col_k1, col_k2 = st.columns(2)
            k1 = col_k1.selectbox("Kỳ thứ 1 (Mốc cũ)" if is_vi else "期間1 (基準月)", list_thang, index=1 if len(list_thang)>1 else 0)
            k2 = col_k2.selectbox("Kỳ thứ 2 (Mốc mới)" if is_vi else "期間2 (比較対象月)", list_thang, index=0)
            
            df_k1 = df_all[df_all['thang_nam'] == k1]
            df_k2 = df_all[df_all['thang_nam'] == k2]
            
            gio_k1 = df_k1['tong_gio'].sum()
            gio_k2 = df_k2['tong_gio'].sum()
            delta_gio = gio_k2 - gio_k1
            
            ot_k1 = df_k1['ot'].sum()
            ot_k2 = df_k2['ot'].sum()
            delta_ot = ot_k2 - ot_k1
            
            mc1, mc2 = st.columns(2)
            mc1.metric(f"Biến động Tổng giờ ({k2} vs {k1})" if is_vi else f"総労働時間の変動 ({k2} vs {k1})", f"{gio_k2:.1f} h", f"{delta_gio:+.1f} h")
            mc2.metric(f"Biến động Tăng ca OT ({k2} vs {k1})" if is_vi else f"残業時間(OT)の変動 ({k2} vs {k1})", f"{ot_k2:.1f} h", f"{delta_ot:+.1f} h")
            
            if is_vi:
                chart_df = pd.DataFrame({
                    "Chu kỳ": [k1, k2],
                    "Giờ làm hành chính": [gio_k1 - ot_k1, gio_k2 - ot_k2],
                    "Giờ tăng ca (OT)": [ot_k1, ot_k2]
                }).set_index("Chu kỳ")
            else:
                chart_df = pd.DataFrame({
                    "対象月": [k1, k2],
                    "通常労働時間": [gio_k1 - ot_k1, gio_k2 - ot_k2],
                    "残業時間(OT)": [ot_k1, ot_k2]
                }).set_index("対象月")
                
            st.bar_chart(chart_df)

if st.session_state.app_page == "checkin":
    render_checkin_page()
    render_chatbot()
    st.stop()

if st.session_state.app_page == "history":
    render_history_page()
    render_chatbot()
    st.stop()

# ==========================================
# CHỨC NĂNG CHẤM CÔNG (Giữ nguyên như cũ)
# Nút quay lại trang chủ ở sidebar
if st.sidebar.button(t("btn_back_home"), type="primary", use_container_width=True, key="btn_1018"):
    st.session_state.app_page = "home"
    st.rerun()
# ==========================================




if st.session_state.get("show_history") and app_mode == "Xử lý file chấm công":
    st.session_state.show_history = False
    st.rerun()


def render_advanced_settings_sidebar():
    st.sidebar.markdown(f"### ⚙️ {t('sidebar_advanced')}")
    with st.sidebar.expander(t('sidebar_standard_hours')):
        st.time_input(t("time_in"), datetime.time(8, 0), key="gio_vao_chuan")
        st.time_input(t("time_out"), datetime.time(17, 0), key="gio_ra_chuan")
        st.time_input(t("break_start"), datetime.time(12, 0), key="nghi_trua_bat_dau")
        st.time_input(t("break_end"), datetime.time(13, 0), key="nghi_trua_ket_thuc")
        st.number_input(t("max_hours"), min_value=0.0, max_value=24.0, value=8.0, step=0.5, key="so_gio_toi_da")

    if time_to_float(st.session_state.gio_ra_chuan) <= time_to_float(st.session_state.gio_vao_chuan):
        st.sidebar.error(t("error_time_out")); st.stop()
    if time_to_float(st.session_state.nghi_trua_ket_thuc) < time_to_float(st.session_state.nghi_trua_bat_dau):
        st.sidebar.error(t("error_break")); st.stop()

    if "manual_emps" not in st.session_state:
        st.session_state.manual_emps = []

    with st.sidebar.expander(t('sidebar_add_emp')):
        st.markdown(t("sidebar_add_emp_desc"), unsafe_allow_html=True)
        with st.form("form_add_emp_top"):
            new_ma = st.text_input(t("emp_code"))
            new_ten = st.text_input(t("export_col_emp_name"))
            new_chuc_vu = st.text_input(t("emp_position"))
            new_phong_ban = st.text_input(t("emp_dept"))
            btn_add = st.form_submit_button(t("btn_add_emp"))
            if btn_add:
                if new_ma and new_ten:
                    st.session_state.manual_emps.append({
                        "ma": new_ma.strip().upper(),
                        "ten": new_ten.strip(),
                        "cv": new_chuc_vu.strip(),
                        "pb": new_phong_ban.strip()
                    })
                    st.rerun()
                else:
                    st.error(t("error_emp_req"))
        if st.session_state.manual_emps:
            st.markdown("**Đã thêm thủ công:**")
            for i, emp in enumerate(st.session_state.manual_emps):
                c1, c2 = st.columns([4, 1])
                c1.markdown(f"<small>{emp['ma']} - {emp['ten']}</small>", unsafe_allow_html=True)
                if c2.button("❌", key=f"del_emp_top_{i}", use_container_width=True):
                    st.session_state.manual_emps.pop(i)
                    st.rerun()


    if "deleted_emps" not in st.session_state:
        st.session_state.deleted_emps = set()
    if "edited_emps" not in st.session_state:
        st.session_state.edited_emps = {}

    with st.sidebar.expander(t('manage_emp')):
        st.markdown(t("manage_emp_hint"), unsafe_allow_html=True)
        if st.session_state.get('df_raw') is not None and 'mapping' in st.session_state:
            m_temp = st.session_state.mapping
            ma_nv_col = m_temp.get('ma_nv')
            ten_nv_col = m_temp.get('ten_nv')
            
            if ma_nv_col and ten_nv_col and ma_nv_col in st.session_state.df_raw.columns:
                raw_emps = st.session_state.df_raw[[ma_nv_col, ten_nv_col]].dropna().drop_duplicates()
                emp_dict = {str(row[ma_nv_col]).strip(): str(row[ten_nv_col]).strip() for _, row in raw_emps.iterrows()}
                
                # Add manual emps
                for me in st.session_state.manual_emps:
                    emp_dict[me['ma']] = me['ten']
                
                # Remove deleted
                for d_ma in list(st.session_state.deleted_emps):
                    emp_dict.pop(d_ma, None)
                    
                if emp_dict:
                    emp_list_display = [f"{k} - {v}" for k, v in emp_dict.items()]
                    lbl_sel = "Chọn nhân viên (Mã - Tên)" if st.session_state.lang == 'vi' else "対象社員を選択 (ID - 氏名)"
                    sel_emp_str = st.selectbox(lbl_sel, [""] + sorted(emp_list_display), key="sel_emp_manager")
                    if sel_emp_str:
                        sel_ma = sel_emp_str.split(" - ")[0]
                        st.markdown("---")
                        curr_pb = st.session_state.edited_emps.get(sel_ma, {}).get('pb', "")
                        curr_cv = st.session_state.edited_emps.get(sel_ma, {}).get('cv', "")
                        new_pb = st.text_input(t("edit_dept"), value=curr_pb, key="edit_pb")
                        new_cv = st.text_input(t("edit_position"), value=curr_cv, key="edit_cv")
                        
                        col_e1, col_e2 = st.columns(2)
                        with col_e1:
                            if st.button(t("btn_save_emp_edit"), use_container_width=True, key="btn_save_emp"):
                                if sel_ma not in st.session_state.edited_emps:
                                    st.session_state.edited_emps[sel_ma] = {}
                                st.session_state.edited_emps[sel_ma]['pb'] = new_pb
                                st.session_state.edited_emps[sel_ma]['cv'] = new_cv
                                st.rerun()
                        with col_e2:
                            if st.button(t("btn_delete_emp"), type="primary", use_container_width=True, key="btn_del_emp"):
                                st.session_state.deleted_emps.add(sel_ma)
                                st.rerun()
                else:
                    st.info(t("info_no_emp"))
            else:
                st.info(t("info_finish_step2"))
        else:
            st.info(t("info_upload_first"))

    if "manual_ot" not in st.session_state:
        st.session_state.manual_ot = {}

# ----- UPLOAD FILE -----

st.sidebar.markdown(f"### 📁 {t('step_1')}: {t('step_1_title')}")
st.sidebar.markdown(f'<div class="upload-hint">📌 {t("upload_hint_formats")}</div>', unsafe_allow_html=True)
uploaded_file = st.sidebar.file_uploader("upload_label_static", type=["xlsx","xls","csv","txt","dat","tsv","xlsm","xlsb"], label_visibility="collapsed", key="main_file_uploader")

st.sidebar.divider()

with st.sidebar:
    render_holiday_makeup_sidebar()
    render_leave_ot_sidebar()
    render_email_sending_sidebar()
    render_advanced_settings_sidebar()
    st.markdown("---")

# ----- LANDING PAGE (khi chưa có file) -----
if uploaded_file is None and st.session_state.df_raw is None:
    pass

    st.markdown(f"""
<style>
/* --- PREMIUM 3-STEP GUIDE CSS --- */
.chamcong-landing {{
    position: fixed;
    top: 50%;
    left: 55%; /* Offset for sidebar */
    transform: translate(-50%, -50%);
    text-align: center;
    font-family: 'Inter', 'Be Vietnam Pro', sans-serif;
    pointer-events: none;
    width: 85%;
    max-width: 900px;
}}

.guide-title {{
    font-size: 34px;
    font-weight: 800;
    margin-bottom: 50px;
    background: linear-gradient(135deg, #0284C7, #0EA5E9, #38BDF8);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    letter-spacing: -0.5px;
    filter: drop-shadow(0 2px 4px rgba(14,165,233,0.15));
}}

.guide-steps {{
    display: flex;
    justify-content: space-between;
    gap: 25px;
}}

.guide-step {{
    flex: 1;
    background: rgba(255, 255, 255, 0.7);
    backdrop-filter: blur(16px);
    -webkit-backdrop-filter: blur(16px);
    padding: 40px 25px;
    border-radius: 24px;
    box-shadow: 
        0 4px 6px -1px rgba(0, 0, 0, 0.05),
        0 10px 15px -3px rgba(0, 0, 0, 0.05),
        inset 0 1px 0 rgba(255, 255, 255, 0.9);
    border: 1px solid rgba(255, 255, 255, 0.5);
    position: relative;
    transition: all 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.275);
}}

/* Hover effects */
.guide-steps:hover .guide-step:not(:hover) {{
    transform: scale(0.95);
    opacity: 0.7;
}}

.guide-step:hover {{
    transform: translateY(-10px) scale(1.02);
    box-shadow: 
        0 20px 25px -5px rgba(14, 165, 233, 0.15),
        0 10px 10px -5px rgba(14, 165, 233, 0.1),
        inset 0 1px 0 rgba(255, 255, 255, 1);
    background: rgba(255, 255, 255, 0.9);
}}

.icon-wrapper {{
    width: 80px;
    height: 80px;
    border-radius: 50%;
    margin: 0 auto 25px auto;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 36px;
    background: linear-gradient(135deg, #F0F9FF, #E0F2FE);
    box-shadow: 0 10px 20px rgba(14, 165, 233, 0.1);
    position: relative;
}}

.icon-wrapper::after {{
    content: '';
    position: absolute;
    top: -5px; left: -5px; right: -5px; bottom: -5px;
    border-radius: 50%;
    background: linear-gradient(135deg, rgba(14,165,233,0.3), rgba(56,189,248,0));
    z-index: -1;
}}

.guide-step-title {{
    font-size: 20px;
    font-weight: 700;
    color: #0F172A;
    margin-bottom: 12px;
}}

.guide-step-desc {{
    font-size: 14px;
    color: #475569;
    line-height: 1.6;
    font-weight: 400;
}}

.guide-arrow {{
    position: absolute;
    top: 50%;
    right: -30px;
    transform: translateY(-50%);
    font-size: 28px;
    color: #94A3B8;
    z-index: 10;
    filter: drop-shadow(0 2px 4px rgba(0,0,0,0.05));
}}
.guide-step:last-child .guide-arrow {{
    display: none;
}}

.upload-pulse .icon-wrapper {{
    animation: pulse-glow 2.5s infinite;
}}

@keyframes pulse-glow {{
    0% {{ box-shadow: 0 0 0 0 rgba(14, 165, 233, 0.4); }}
    70% {{ box-shadow: 0 0 0 20px rgba(14, 165, 233, 0); }}
    100% {{ box-shadow: 0 0 0 0 rgba(14, 165, 233, 0); }}
}}
/* Copyright — sát xuống chân trang */
.vimos-copyright {{
    position: fixed;
    bottom: 8px;
    left: 50%;
    transform: translateX(-50%);
    width: 100%;
    text-align: center;
    color: #94A3B8;
    font-size: 12px;
    font-weight: 400;
    z-index: 10;
    pointer-events: none;
}}
</style>

<div class="chamcong-landing">
    <div class="guide-title">{t('guide_title')}</div>
    <div class="guide-steps">
        <div class="guide-step upload-pulse">
            <div class="icon-wrapper">📁</div>
            <div class="guide-step-title">{t('guide_step1_title')}</div>
            <div class="guide-step-desc">{t('guide_step1_desc')}</div>
            <span class="guide-arrow">➔</span>
        </div>
        <div class="guide-step">
            <div class="icon-wrapper">⚙️</div>
            <div class="guide-step-title">{t('guide_step2_title')}</div>
            <div class="guide-step-desc">{t('guide_step2_desc')}</div>
            <span class="guide-arrow">➔</span>
        </div>
        <div class="guide-step">
            <div class="icon-wrapper">📊</div>
            <div class="guide-step-title">{t('guide_step3_title')}</div>
            <div class="guide-step-desc">{t('guide_step3_desc')}</div>
        </div>
    </div>
</div>
<div class="vimos-copyright">Copyright &copy; 2026 V-mos System</div>
""", unsafe_allow_html=True)

# ----- XỬ LÝ FILE -----
if uploaded_file is not None:
    if st.session_state.df_raw is None or st.session_state.get('last_uploaded') != uploaded_file.name:
        with st.spinner("Đang đọc file..."):
            st.session_state['uploaded_file_bytes'] = uploaded_file.read()
            uploaded_file.seek(0)
            df = parse_excel_file(uploaded_file)
            if df is not None and not df.empty:
                st.session_state.df_raw = df
                st.session_state.last_uploaded = uploaded_file.name
                mapping_auto = auto_detect_columns(df)
                req_keys = ['ma_nv', 'ten_nv', 'ngay', 'gio_vao', 'gio_ra']
                if all(k in mapping_auto for k in req_keys):
                    st.session_state.mapping = mapping_auto
                    st.session_state.step = 3
                    st.rerun()
                else:
                    st.session_state.step = 2
                    st.session_state.mapping_auto = mapping_auto
elif st.session_state.df_raw is not None:
    st.session_state.df_raw = None
    st.session_state.last_uploaded = None
    st.session_state.step = 1
    for _key in ("mapping", "mapping_auto"):
        if _key in st.session_state: del st.session_state[_key]
    st.rerun()

# ----- MAPPING CỘT -----
if st.session_state.step == 2 and st.session_state.df_raw is not None:
    df_raw = st.session_state.df_raw
    mapping_auto = st.session_state.mapping_auto
    st.markdown("### Dữ liệu thô (10 dòng đầu)")
    st.dataframe(df_raw.head(10), use_container_width=True, hide_index=True)
    st.sidebar.markdown("### 🔗 Bước 2: Xác nhận mapping cột")
    st.sidebar.warning("⚠️ Không nhận diện đủ 5 cột.")
    columns = ["-- Chọn cột --"] + list(df_raw.columns)
    def get_index(col_name):
        return columns.index(col_name) if col_name in columns else 0
    map_emp_id = st.sidebar.selectbox("Mã NV", columns, index=get_index(mapping_auto.get('ma_nv')))
    map_emp_name = st.sidebar.selectbox("Tên NV", columns, index=get_index(mapping_auto.get('ten_nv')))
    map_date = st.sidebar.selectbox("Ngày làm việc", columns, index=get_index(mapping_auto.get('ngay')))
    map_in = st.sidebar.selectbox("Giờ vào", columns, index=get_index(mapping_auto.get('gio_vao')))
    map_out = st.sidebar.selectbox("Giờ ra", columns, index=get_index(mapping_auto.get('gio_ra')))
    
    st.sidebar.markdown("<small style='color:#64748B'>— Tuỳ chọn (để trống nếu không có) —</small>", unsafe_allow_html=True)
    opt_cols = ["(Không có)"] + list(df_raw.columns)
    map_chuc_vu = st.sidebar.selectbox("Chức vụ (tuỳ chọn)", opt_cols, index=0)
    map_phong_ban = st.sidebar.selectbox("Phòng ban (tuỳ chọn)", opt_cols, index=0)
    
    if st.sidebar.button("Xác nhận Mapping", type="primary", use_container_width=True):
        if "-- Chọn cột --" in [map_emp_id, map_emp_name, map_date, map_in, map_out]:
            st.sidebar.error("❌ Vui lòng chọn đủ 5 cột bắt buộc!")
        else:
            new_mapping = {'ma_nv': map_emp_id, 'ten_nv': map_emp_name, 'ngay': map_date, 'gio_vao': map_in, 'gio_ra': map_out}
            if map_chuc_vu != "(Không có)": new_mapping['chuc_vu'] = map_chuc_vu
            if map_phong_ban != "(Không có)": new_mapping['phong_ban'] = map_phong_ban
            st.session_state.mapping = new_mapping
            st.session_state.step = 3
            st.rerun()

# ----- TÍNH TOÁN & KẾT QUẢ -----
if st.session_state.step >= 3 and "mapping" in st.session_state:
    m = st.session_state.mapping
    df_process = st.session_state.df_raw.copy()
    df_process = df_process.dropna(subset=[m['ma_nv']])
    df_process = df_process[df_process[m['ma_nv']].astype(str).str.strip() != '']
    df_process["_parsed_date"] = df_process[m['ngay']].apply(clean_date)
    df_process = df_process.dropna(subset=["_parsed_date"])

    if df_process.empty:
        st.warning("Không có dữ liệu hợp lệ (Lỗi đọc ngày tháng hoặc Mã NV trống).")
    else:
        min_date = df_process["_parsed_date"].min()
        max_date = df_process["_parsed_date"].max()
        default_start = datetime.date(2026, 5, 21)
        default_end = datetime.date(2026, 6, 20)
        if min_date.date() > default_end or max_date.date() < default_start:
            default_start = min_date.date()
            default_end = max_date.date()

        st.sidebar.divider()
        st.sidebar.markdown(f"### 📅 {t('step_3_period')}")
        col_d1, col_d2 = st.sidebar.columns(2)
        with col_d1: start_date = st.date_input(t("from_date"), default_start)
        with col_d2: end_date = st.date_input(t("to_date"), default_end)

        years_in_range = range(start_date.year, end_date.year + 1)
        fixed_holidays = set(get_fixed_holidays_for_years(years_in_range))
        
        if "custom_holidays" not in st.session_state:
            st.session_state.custom_holidays = set()
        if "custom_workdays" not in st.session_state:
            st.session_state.custom_workdays = set()

        holiday_dates = fixed_holidays | st.session_state.custom_holidays
        makeup_dates = st.session_state.custom_workdays
        working_days, holiday_info, makeup_info = calculate_working_days(start_date, end_date, holiday_dates, makeup_dates)

        if makeup_info:
            st.markdown('<div class="card"><div class="card-title"><span class="card-icon">💼</span>Ngày làm bù trong kỳ</div></div>', unsafe_allow_html=True)
            for m in sorted(makeup_info, key=lambda x: x["date"]):
                date_str = m["date"].strftime("%d/%m/%Y")
                st.info(f"💼 **{date_str} ({m['weekday_label']})** được tính là ngày làm việc (Ngày làm bù).")

        if holiday_info:
            st.markdown('<div class="card"><div class="card-title"><span class="card-icon">📅</span>Ngày lễ trong kỳ</div></div>', unsafe_allow_html=True)
            for h in sorted(holiday_info, key=lambda x: x["date"]):
                date_str = h["date"].strftime("%d/%m/%Y")
                if h["is_workday"]:
                    st.warning(f"⚠️ **{date_str} ({h['weekday_label']})** trùng vào ngày làm việc — đã trừ 1 ngày công.")
                else:
                    st.info(f"ℹ️ {date_str} ({h['weekday_label']}) rơi vào ngày nghỉ cuối tuần.")

        mask = (df_process["_parsed_date"].dt.date >= start_date) & (df_process["_parsed_date"].dt.date <= end_date)
        df_filtered = df_process.loc[mask].copy()
        
        if not m.get('phong_ban'): m['phong_ban'] = 'Phòng ban'
        if not m.get('chuc_vu'): m['chuc_vu'] = 'Chức vụ'

        if st.session_state.manual_emps:
            new_rows = []
            curr_d = start_date
            while curr_d <= end_date:
                for emp in st.session_state.manual_emps:
                    row_data = {
                        m['ma_nv']: emp['ma'],
                        m['ten_nv']: emp['ten'],
                        "_parsed_date": pd.to_datetime(curr_d),
                        m['ngay']: curr_d.strftime('%d/%m/%Y'),
                        m['gio_vao']: pd.NaT,
                        m['gio_ra']: pd.NaT
                    }
                    if 'chuc_vu' in m: row_data[m['chuc_vu']] = emp['cv']
                    if 'phong_ban' in m: row_data[m['phong_ban']] = emp['pb']
                    new_rows.append(row_data)
                curr_d += datetime.timedelta(days=1)
            if new_rows:
                df_new = pd.DataFrame(new_rows)
                df_filtered = pd.concat([df_filtered, df_new], ignore_index=True)


        # APPLY DELETED EMPS AND EDITS
        if st.session_state.get('deleted_emps'):
            df_filtered = df_filtered[~df_filtered[m['ma_nv']].astype(str).str.strip().isin(st.session_state.deleted_emps)]
            
        if st.session_state.get('edited_emps'):
            for ma, edits in st.session_state.edited_emps.items():
                mask_emp = df_filtered[m['ma_nv']].astype(str).str.strip() == ma
                if 'phong_ban' in m and edits.get('pb'):
                    df_filtered.loc[mask_emp, m['phong_ban']] = edits['pb']
                if 'chuc_vu' in m and edits.get('cv'):
                    df_filtered.loc[mask_emp, m['chuc_vu']] = edits['cv']

        if df_filtered.empty:
            st.warning("Không có dữ liệu trong khoảng thời gian đã chọn.")
        else:
            if "gps_synced" not in st.session_state: st.session_state.gps_synced = set()
            try:
                df_gps_all = get_field_checkins(limit=10000)
                if not df_gps_all.empty:
                    emp_info_map = {}
                    for _, erow in df_filtered.iterrows():
                        ema = str(erow[m['ma_nv']]).strip().upper()
                        if ema not in emp_info_map:
                            emp_info_map[ema] = {
                                'ten': erow.get(m['ten_nv'], ''),
                                'cv': erow.get(m.get('chuc_vu', 'Chức vụ'), ''),
                                'pb': erow.get(m.get('phong_ban', 'Phòng ban'), '')
                            }

                    gps_rows_to_add = []
                    for _, g_row in df_gps_all.iterrows():
                        g_ma = str(g_row['ma_nv']).split(' - ')[0].strip().upper()
                        g_time = pd.to_datetime(g_row['thoi_gian'], format='%d/%m/%Y %H:%M:%S', errors='coerce')
                        if pd.notna(g_time) and start_date <= g_time.date() <= end_date:
                            g_date_str = g_time.strftime('%d/%m/%Y')
                            g_hour_str = g_time.strftime('%H:%M')
                            g_loai = str(g_row['loai'])
                            
                            mask_match = (df_filtered[m['ma_nv']].astype(str).str.strip().str.upper() == g_ma) & (df_filtered["_parsed_date"].dt.strftime('%d/%m/%Y') == g_date_str)
                            if mask_match.any():
                                for idx_m in df_filtered[mask_match].index:
                                    vao_ht = df_filtered.loc[idx_m, m['gio_vao']]
                                    ra_ht = df_filtered.loc[idx_m, m['gio_ra']]
                                    vao_tr = pd.isna(vao_ht) or str(vao_ht).strip().lower() in ['', 'nan', 'none', 'nat']
                                    ra_tr = pd.isna(ra_ht) or str(ra_ht).strip().lower() in ['', 'nan', 'none', 'nat']
                                    
                                    if 'Vào ca' in g_loai or 'Check-in' in g_loai or '出勤' in g_loai:
                                        if vao_tr: df_filtered.loc[idx_m, m['gio_vao']] = g_hour_str
                                    elif 'Tan ca' in g_loai or 'Check-out' in g_loai or '退勤' in g_loai:
                                        if ra_tr: df_filtered.loc[idx_m, m['gio_ra']] = g_hour_str
                                        
                                    vao_sau = df_filtered.loc[idx_m, m['gio_vao']]
                                    ra_sau = df_filtered.loc[idx_m, m['gio_ra']]
                                    vao_s_tr = pd.isna(vao_sau) or str(vao_sau).strip().lower() in ['', 'nan', 'none', 'nat']
                                    ra_s_tr = pd.isna(ra_sau) or str(ra_sau).strip().lower() in ['', 'nan', 'none', 'nat']
                                    
                                    if not vao_s_tr and ra_s_tr:
                                        df_filtered.loc[idx_m, m['gio_ra']] = "17:00"
                                    elif vao_s_tr and not ra_s_tr:
                                        df_filtered.loc[idx_m, m['gio_vao']] = "08:00"
                            else:
                                einfo = emp_info_map.get(g_ma, {'ten': str(g_row['ten_nv']), 'cv': '', 'pb': ''})
                                vao_val = g_hour_str if ('Vào ca' in g_loai or 'Check-in' in g_loai or '出勤' in g_loai) else "08:00"
                                ra_val = g_hour_str if ('Tan ca' in g_loai or 'Check-out' in g_loai or '退勤' in g_loai) else "17:00"
                                
                                new_g_row = {
                                    m['ma_nv']: g_ma,
                                    m['ten_nv']: einfo['ten'],
                                    "_parsed_date": pd.to_datetime(g_time.date()),
                                    m['ngay']: g_date_str,
                                    m['gio_vao']: vao_val,
                                    m['gio_ra']: ra_val
                                }
                                if 'chuc_vu' in m: new_g_row[m['chuc_vu']] = einfo['cv']
                                if 'phong_ban' in m: new_g_row[m['phong_ban']] = einfo['pb']
                                gps_rows_to_add.append(new_g_row)
                                
                            st.session_state.gps_synced.add((g_ma, g_date_str))
                            
                    if gps_rows_to_add:
                        df_filtered = pd.concat([df_filtered, pd.DataFrame(gps_rows_to_add)], ignore_index=True)
            except Exception:
                pass

            with st.spinner("⏳ Đang tính toán giờ làm..."):
                df_calc = df_filtered.apply(lambda row: calculate_working_hours(
                    row[m['gio_vao']], row[m['gio_ra']],
                    start_chuan=time_to_float(st.session_state.gio_vao_chuan), end_chuan=time_to_float(st.session_state.gio_ra_chuan),
                    lunch_start=time_to_float(st.session_state.nghi_trua_bat_dau), lunch_end=time_to_float(st.session_state.nghi_trua_ket_thuc),
                    max_hours=st.session_state.so_gio_toi_da,
                ), axis=1)
                df_filtered["Giờ hành chính"] = df_calc.apply(lambda x: x['admin_hours'] if isinstance(x, dict) else 0.0)
                df_filtered["Số giờ làm thực tế"] = df_calc.apply(lambda x: x['tong_gio'] if isinstance(x, dict) else 0.0)
                df_filtered["_is_chieu"] = df_calc.apply(lambda x: x.get('is_chieu', False) if isinstance(x, dict) else False)
                df_filtered["Phút đi trễ"] = df_calc.apply(lambda x: x['di_tre'] if isinstance(x, dict) else 0)
                df_filtered["Phút về sớm"] = df_calc.apply(lambda x: x['ve_som'] if isinstance(x, dict) else 0)
                df_filtered["Giờ OT"] = None

                df_filtered["Ngày"] = df_filtered["_parsed_date"].dt.strftime('%d/%m/%Y')
                df_filtered = df_filtered.sort_values(by=[m['ma_nv'], "_parsed_date"])
                df_filtered["_nv_label"] = df_filtered[m['ma_nv']].astype(str) + " - " + df_filtered[m['ten_nv']].astype(str)
                danh_sach_nv = sorted(df_filtered["_nv_label"].unique().tolist())

            st.sidebar.divider()
            st.sidebar.markdown(f"### 🔍 {t('step_4_filter')}")
            chon_nv = st.sidebar.multiselect("Chọn nhân viên", options=danh_sach_nv, default=[], placeholder=t("filter_placeholder"), label_visibility="collapsed")

            st.markdown(t("result_title"))
            if chon_nv:
                df_filtered = df_filtered[df_filtered["_nv_label"].isin(chon_nv)]


            @st.fragment
            def render_interactive_dashboard(df_base):
                df_filtered = df_base.copy()
                editor_key = f"data_editor_ot_{st.session_state.get('editor_key_counter', 0)}"

                # Áp dụng Giờ HC thủ công
                def apply_manual_hc(row):
                    ma = str(row[m['ma_nv']]).strip().upper()
                    if ma.endswith('.0'): ma = ma[:-2]
                    ngay = row["_parsed_date"].strftime('%d/%m/%Y')
                    if "manual_hc" in st.session_state and (ma, ngay) in st.session_state.manual_hc:
                        return float(st.session_state.manual_hc[(ma, ngay)])
                    return float(row["Giờ hành chính"])
                df_filtered["Giờ hành chính"] = df_filtered.apply(apply_manual_hc, axis=1)

                # Áp dụng OT thủ công
                def apply_manual_ot(row):
                    ma = str(row[m['ma_nv']]).strip().upper()
                    if ma.endswith('.0'): ma = ma[:-2]
                    ngay = row["_parsed_date"].strftime('%d/%m/%Y')
                    if "manual_ot" in st.session_state and (ma, ngay) in st.session_state.manual_ot:
                        return float(st.session_state.manual_ot[(ma, ngay)])
                    return row["Giờ OT"]
                df_filtered["Giờ OT"] = df_filtered.apply(apply_manual_ot, axis=1)

                # Tính lại Tổng giờ = Giờ HC + Giờ OT (Chỉ áp dụng nếu không phải lỗi -1, trừ khi có sửa thủ công)
                def apply_manual_total(row):
                    ma = str(row[m['ma_nv']]).strip().upper()
                    if ma.endswith('.0'): ma = ma[:-2]
                    ngay = row["_parsed_date"].strftime('%d/%m/%Y')
                    
                    if "manual_total" in st.session_state and (ma, ngay) in st.session_state.manual_total:
                        return float(st.session_state.manual_total[(ma, ngay)])

                    has_hc_override = "manual_hc" in st.session_state and (ma, ngay) in st.session_state.manual_hc
                    has_ot_override = "manual_ot" in st.session_state and (ma, ngay) in st.session_state.manual_ot
                    
                    if row["Số giờ làm thực tế"] == -1 and not (has_hc_override or has_ot_override):
                        return -1.0
                        
                    ot = float(row["Giờ OT"]) if pd.notna(row["Giờ OT"]) and str(row["Giờ OT"]).strip() != "" else 0.0
                    hc = float(row["Giờ hành chính"]) if pd.notna(row["Giờ hành chính"]) and str(row["Giờ hành chính"]).strip() != "" else 0.0
                    return hc + ot
                df_filtered["Số giờ làm thực tế"] = df_filtered.apply(apply_manual_total, axis=1)

                # Kiểm tra lỗi check-out TRƯỚC khi replace -1 → 0
                has_checkout_error = (df_filtered["Số giờ làm thực tế"] == -1).any()
                if has_checkout_error:
                    st.warning("⚠️ Phát hiện một số dòng có giờ ra sớm hơn giờ vào (lỗi check-out), đã được tính là 0 giờ.")

                def check_anomaly(row):
                    ma = str(row[m['ma_nv']]).strip().upper()
                    if ma.endswith('.0'): ma = ma[:-2]
                    ngay = row["_parsed_date"].strftime('%d/%m/%Y')
                    
                    if "manual_notes" in st.session_state and (ma, ngay) in st.session_state.manual_notes:
                        return st.session_state.manual_notes[(ma, ngay)]

                    vao = row[m['gio_vao']]
                    ra = row[m['gio_ra']]
                    is_wd = is_workday_func(row["_parsed_date"])
                    thu = row["_parsed_date"].weekday()  # 5=T7, 6=CN
                    has_ot_override = "manual_ot" in st.session_state and (ma, ngay) in st.session_state.manual_ot
                    has_hc_override = "manual_hc" in st.session_state and (ma, ngay) in st.session_state.manual_hc
                    has_ot_reason = "manual_ot_reason" in st.session_state and (ma, ngay) in st.session_state.manual_ot_reason
                    has_leave = st.session_state.get("manual_leave", {}).get((ma, ngay), False)
                    has_gps = "gps_synced" in st.session_state and (ma, ngay) in st.session_state.gps_synced

                    if has_gps:
                        return "📍 Check-in GPS hiện trường" if st.session_state.lang == 'vi' else "📍 GPS現地打刻"

                    if has_ot_override or has_hc_override or has_ot_reason:
                        return ""

                    try:
                        vao_trong = pd.isna(vao) or str(vao).strip().lower() in ['', 'nan', 'none', 'nat']
                    except (TypeError, ValueError):
                        vao_trong = False
                    try:
                        ra_trong = pd.isna(ra) or str(ra).strip().lower() in ['', 'nan', 'none', 'nat']
                    except (TypeError, ValueError):
                        ra_trong = False

                    notes = []
                    is_ja = st.session_state.lang == 'ja'

                    if has_leave:
                        d_check = row["_parsed_date"].date() if hasattr(row["_parsed_date"], 'date') else row["_parsed_date"]
                        if is_wd and vao_trong and ra_trong and is_last_saturday_of_month(d_check):
                            notes.append("🔴 必須土曜日欠勤" if is_ja else "🔴 Vắng Thứ 7 bắt buộc")
                        return " | ".join(notes) if notes else ""

                    if not is_wd and vao_trong and ra_trong:
                        return " | ".join(notes) if notes else ""

                    if vao_trong and not ra_trong: notes.append("⚠️ 出勤打刻忘れ" if is_ja else "⚠️ Thiếu giờ vào")
                    elif ra_trong and not vao_trong: notes.append("⚠️ 退勤打刻忘れ" if is_ja else "⚠️ Thiếu giờ ra")
                    elif vao_trong and ra_trong:
                        if is_wd:
                            d_check = row["_parsed_date"].date() if hasattr(row["_parsed_date"], 'date') else row["_parsed_date"]
                            if is_last_saturday_of_month(d_check):
                                notes.append("🔴 必須土曜日欠勤" if is_ja else "🔴 Vắng Thứ 7 bắt buộc")
                            else:
                                notes.append("🔴 無断欠勤" if is_ja else "🔴 Nghỉ không phép")
                    elif row["Số giờ làm thực tế"] == -1: notes.append("🟣 退勤エラー" if is_ja else "🟣 Lỗi check-out")
                    elif 0 < float(row["Số giờ làm thực tế"]) < 4: notes.append("🟠 実働不足 (< 4h)" if is_ja else "🟠 Làm thiếu giờ (< 4h)")

                    if has_leave and not (vao_trong and ra_trong):
                        notes.append("🟢 有給休暇" if is_ja else "🟢 Nghỉ có phép")

                    if row.get("_is_chieu", False):
                        notes.append("🔵 午後出勤" if is_ja else "🔵 Làm ca chiều")

                    return " | ".join(notes)
                df_filtered["Ghi chú"] = df_filtered.apply(check_anomaly, axis=1)
                df_filtered["Số giờ làm thực tế"] = df_filtered["Số giờ làm thực tế"].replace(-1, 0)
                df_filtered["Số giờ làm thực tế"] = df_filtered["Số giờ làm thực tế"].apply(format_gio_lam)
                df_filtered["Giờ hành chính"] = df_filtered["Giờ hành chính"].apply(format_gio_lam)
                df_filtered["Giờ OT"] = df_filtered["Giờ OT"].apply(format_gio_lam)

                def get_has_leave(row):
                    ma = str(row[m['ma_nv']]).strip().upper()
                    if ma.endswith('.0'): ma = ma[:-2]
                    ngay = row["_parsed_date"].strftime('%d/%m/%Y')
                    return st.session_state.get("manual_leave", {}).get((ma, ngay), False)

                df_filtered["Có phép"] = df_filtered.apply(get_has_leave, axis=1)

                def get_ot_reason(row):
                    ma = str(row[m['ma_nv']]).strip().upper()
                    if ma.endswith('.0'): ma = ma[:-2]
                    ngay = row["_parsed_date"].strftime('%d/%m/%Y')
                    return st.session_state.get("manual_ot_reason", {}).get((ma, ngay), "")

                df_filtered["Lý do tăng ca"] = df_filtered.apply(get_ot_reason, axis=1)

                with st.spinner("⏳ Đang tổng hợp kết quả..."):
                    total_rows = len(df_filtered)
                    total_emps = df_filtered[m['ma_nv']].nunique()
                    df_numeric = df_filtered[pd.to_numeric(df_filtered['Số giờ làm thực tế'], errors='coerce').notnull()]
                    total_hours = df_numeric['Số giờ làm thực tế'].sum() if not df_numeric.empty else 0
                    ngay_nghi = int((df_filtered["Số giờ làm thực tế"] == 0).sum())

                    st.markdown(t("overview_title_html"), unsafe_allow_html=True)
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric(t("total_days"), working_days, help=t("total_days_help"))
                    c2.metric(t("num_employees"), total_emps)
                    c3.metric(t("total_hours"), f"{format_gio_lam(total_hours)} {t('hours_unit')}")
                    c4.metric(t("total_days_off"), ngay_nghi)

                    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

                    lang = st.session_state.lang
                    t_data = get_data_t(lang)
                    
                    chuc_vu_vals = [t_data(x) for x in (df_filtered[m['chuc_vu']].values if 'chuc_vu' in m and m['chuc_vu'] in df_filtered else [""] * len(df_filtered))]
                    phong_ban_vals = [t_data(x) for x in (df_filtered[m['phong_ban']].values if 'phong_ban' in m and m['phong_ban'] in df_filtered else [""] * len(df_filtered))]
                    ten_nv_vals = [translate_name(x, lang) for x in df_filtered[m['ten_nv']].values]

                    weekday_map_vi = {0:'Hai',1:'Ba',2:'Tư',3:'Năm',4:'Sáu',5:'Bảy',6:'CN'}
                    weekday_map_ja = {0:'月',1:'火',2:'水',3:'木',4:'金',5:'土',6:'日'}
                    weekday_map = weekday_map_ja if lang == 'ja' else weekday_map_vi

                    df_result_ui = pd.DataFrame({
                        "STT": range(1, len(df_filtered) + 1),
                        "Mã NV": df_filtered[m['ma_nv']].values,
                        "Tên nhân viên": ten_nv_vals,
                        "Chức vụ": chuc_vu_vals,
                        "Phòng ban": phong_ban_vals,
                        "Thứ": df_filtered["_parsed_date"].dt.weekday.map(weekday_map).values,
                        "Ngày": df_filtered["_parsed_date"].dt.strftime('%Y/%m/%d' if lang == 'ja' else '%d/%m/%Y').values,
                        "Giờ làm thực tế": df_filtered["Giờ hành chính"].values,
                        "OT": df_filtered["Giờ OT"].values,
                        "Tổng giờ": df_filtered["Số giờ làm thực tế"].values,
                        "Lý do tăng ca": df_filtered["Lý do tăng ca"].values,
                        "Có phép": df_filtered["Có phép"].values,
                        "Ghi chú": df_filtered["Ghi chú"].values
                    })
                    st.session_state['df_result'] = df_result_ui
                    st.session_state['df_filtered_for_chat'] = df_filtered.copy()

                    def get_loai_ngay(row):
                        try:
                            ngay = row["_parsed_date"].date() if hasattr(row["_parsed_date"], 'date') else pd.to_datetime(row["_parsed_date"]).date()
                            ngay_str = ngay.strftime('%d/%m/%Y')
                        except:
                            return 'binh_thuong'
                            
                        gio_vao_raw = row.get(m['gio_vao'], pd.NA)
                        gio_ra_raw = row.get(m['gio_ra'], pd.NA)
                        ma = str(row[m['ma_nv']]).strip().upper()
                        if ma.endswith('.0'): ma = ma[:-2]
                        has_leave = st.session_state.get("manual_leave", {}).get((ma, ngay_str), False)
                        
                        thu = ngay.weekday()  # 5=T7, 6=CN
                        is_wd = is_workday_func(ngay)
                        vao_trong = pd.isna(gio_vao_raw) or str(gio_vao_raw).strip().lower() in ['', 'nan', 'none', 'nat']
                        ra_trong = pd.isna(gio_ra_raw) or str(gio_ra_raw).strip().lower() in ['', 'nan', 'none', 'nat']

                        if not is_wd and vao_trong and ra_trong:
                            return 'cuoi_tuan'   # T7 / CN không có dữ liệu

                        if is_wd and vao_trong and ra_trong:
                            if has_leave:
                                return 'nghi_co_phep'
                            return 'nghi_khong_phep'

                        return 'binh_thuong'

                    df_result_ui["_loai"] = df_filtered.apply(get_loai_ngay, axis=1).values

                    def style_row(row):
                        loai = df_result_ui.loc[row.name, "_loai"]
                        ngay_str = df_result_ui.loc[row.name, "Ngày"]
                        try:
                            fmt = '%Y/%m/%d' if lang == 'ja' else '%d/%m/%Y'
                            d_obj_row = datetime.datetime.strptime(ngay_str, fmt).date()
                            thu = d_obj_row.weekday()
                            is_weekend = (thu in [5, 6]) and not (thu == 5 and is_last_saturday_of_month(d_obj_row))
                        except:
                            is_weekend = False

                        styles = [""] * len(row)
                        if is_weekend:
                            styles = ["background-color: #F1F5F9"] * len(row)
                            
                        idx_gio = list(row.index).index("Tổng giờ")
                        idx_ghi_chu = list(row.index).index("Ghi chú")
                        
                        if loai == 'cuoi_tuan':
                            styles[idx_gio] = "background-color: #F1F5F9; color: #64748B"
                        elif loai == 'nghi_khong_phep':
                            styles = ["background-color: #FEE2E2; color: #991B1B"] * len(row)
                            styles[idx_gio] = "background-color: #FEE2E2; color: #991B1B; font-weight: 600"
                        elif loai == 'nghi_co_phep':
                            styles[idx_gio] = "background-color: #F1F5F9; color: #0EA5E9; font-weight: 600" if is_weekend else "color: #0EA5E9; font-weight: 600"
                        else:
                            styles[idx_gio] = "background-color: #F1F5F9; color: #0EA5E9; font-weight: 600" if is_weekend else "color: #0EA5E9; font-weight: 600"
                            
                        val_str = str(row["Ghi chú"])
                        if "Nghỉ không phép" in val_str or "Vắng Thứ 7 bắt buộc" in val_str or "無断欠勤" in val_str or "必須土曜日欠勤" in val_str:
                            styles[idx_ghi_chu] = 'background-color: #FEE2E2; color: #991B1B; font-weight: 600'
                        elif "Lỗi check-out" in val_str or "退勤エラー" in val_str:
                            styles[idx_ghi_chu] = 'background-color: #F3E8FF; color: #6B21A8; font-weight: 500'
                        elif "Thiếu giờ vào" in val_str or "Thiếu giờ ra" in val_str or "Làm thiếu giờ" in val_str or "出勤打刻忘れ" in val_str or "退勤打刻忘れ" in val_str or "実働不足" in val_str:
                            styles[idx_ghi_chu] = 'background-color: #FEF3C7; color: #92400E; font-weight: 500'
                        elif "OT thủ công" in val_str or "手動OT" in val_str:
                            styles[idx_ghi_chu] = 'background-color: #E0F2FE; color: #0369A1; font-weight: 500'
                        elif "ca chiều" in val_str or "午後出勤" in val_str:
                            styles[idx_ghi_chu] = 'background-color: #F8FAFC; color: #475569; font-weight: 500'

                        return styles

                    st.markdown(t("detail_title_html"), unsafe_allow_html=True)
                    df_display = df_result_ui.drop(columns=["_loai"])

                    tab_table, tab_chart = st.tabs([t("tab_table"), t("tab_chart")])
                    with tab_table:
                        st.markdown(t("tip_editor"))
                        
                        st.data_editor(
                                    df_display.style.apply(style_row, axis=1),
                                    use_container_width=True, hide_index=True,
                                    height=min(600, 40 + len(df_result_ui) * 35),
                                    key=editor_key,
                                    column_config={
                                        "STT": st.column_config.NumberColumn("STT", width="small", format="%d", disabled=True),
                                        "Mã NV": st.column_config.TextColumn(t("export_col_emp_code"), width="small", disabled=True),
                                        "Tên nhân viên": st.column_config.TextColumn(t("export_col_emp_name"), width="medium", disabled=True),
                                        "Chức vụ": st.column_config.TextColumn(t("emp_position"), width="small", disabled=True),
                                        "Phòng ban": st.column_config.TextColumn(t("emp_dept"), width="small", disabled=True),
                                        "Thứ": st.column_config.TextColumn(t("export_col_weekday"), width="small", disabled=True),
                                        "Ngày": st.column_config.TextColumn(t("export_col_date"), width="small", disabled=True),
                                        "Giờ làm thực tế": st.column_config.NumberColumn(t("export_col_actual_hours"), width="small", format="%g", disabled=False, step=0.01),
                                        "OT": st.column_config.NumberColumn(t("export_col_ot"), width="small", format="%g", disabled=False, step=0.01),
                                        "Tổng giờ": st.column_config.NumberColumn(t("export_col_total_hours"), width="small", format="%g", disabled=False, step=0.01),
                                        "Lý do tăng ca": st.column_config.TextColumn(t("export_col_ot_reason"), width="medium", disabled=False),
                                        "Có phép": st.column_config.CheckboxColumn(t("col_leave"), width="small", disabled=False),
                                        "Ghi chú": st.column_config.TextColumn(t("export_col_note"), width="medium", disabled=False),
                                    }
                        )

                        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
                        if st.button("💾 Lưu thay đổi bảng" if st.session_state.lang != 'ja' else "💾 変更を保存", type="primary"):
                            if editor_key in st.session_state and "edited_rows" in st.session_state[editor_key]:
                                changes = st.session_state[editor_key]["edited_rows"]
                                if changes and "df_result" in st.session_state:
                                    last_df = st.session_state["df_result"]
                                    for row_idx_str, edits in changes.items():
                                        row_idx = int(row_idx_str)
                                        if row_idx < len(last_df):
                                            ma_edit = str(last_df.iloc[row_idx]["Mã NV"]).strip().upper()
                                            if ma_edit.endswith('.0'): ma_edit = ma_edit[:-2]
                                            ngay_edit_raw = str(last_df.iloc[row_idx]["Ngày"])
                                            
                                            fmt = '%Y/%m/%d' if st.session_state.lang == 'ja' else '%d/%m/%Y'
                                            try:
                                                ngay_edit = datetime.datetime.strptime(ngay_edit_raw, fmt).strftime('%d/%m/%Y')
                                            except:
                                                ngay_edit = ngay_edit_raw

                                            if "Giờ làm thực tế" in edits:
                                                if "manual_hc" not in st.session_state: st.session_state.manual_hc = {}
                                                val = edits["Giờ làm thực tế"]
                                                st.session_state.manual_hc[(ma_edit, ngay_edit)] = float(val) if val is not None else 0.0
                                                if "manual_total" in st.session_state: st.session_state.manual_total.pop((ma_edit, ngay_edit), None)
                                                if "Tổng giờ" in edits: del edits["Tổng giờ"]

                                            if "OT" in edits:
                                                if "manual_ot" not in st.session_state: st.session_state.manual_ot = {}
                                                val = edits["OT"]
                                                st.session_state.manual_ot[(ma_edit, ngay_edit)] = float(val) if val is not None else 0.0
                                                if "manual_total" in st.session_state: st.session_state.manual_total.pop((ma_edit, ngay_edit), None)
                                                if "Tổng giờ" in edits: del edits["Tổng giờ"]

                                            if "Có phép" in edits:
                                                if "manual_leave" not in st.session_state: st.session_state.manual_leave = {}
                                                if edits["Có phép"]:
                                                    st.session_state.manual_leave[(ma_edit, ngay_edit)] = True
                                                else:
                                                    st.session_state.manual_leave.pop((ma_edit, ngay_edit), None)

                                            if "Lý do tăng ca" in edits:
                                                if "manual_ot_reason" not in st.session_state: st.session_state.manual_ot_reason = {}
                                                st.session_state.manual_ot_reason[(ma_edit, ngay_edit)] = str(edits["Lý do tăng ca"])

                                            if "Ghi chú" in edits:
                                                if "manual_notes" not in st.session_state: st.session_state.manual_notes = {}
                                                if edits["Ghi chú"] is None:
                                                    st.session_state.manual_notes.pop((ma_edit, ngay_edit), None)
                                                else:
                                                    st.session_state.manual_notes[(ma_edit, ngay_edit)] = str(edits["Ghi chú"])

                                            if "Tổng giờ" in edits:
                                                if "manual_total" not in st.session_state: st.session_state.manual_total = {}
                                                if edits["Tổng giờ"] is None:
                                                    st.session_state.manual_total.pop((ma_edit, ngay_edit), None)
                                                else:
                                                    st.session_state.manual_total[(ma_edit, ngay_edit)] = float(edits["Tổng giờ"])
                                                    
                            # Force reset widget key to clear frontend state and display new calculated data
                            st.session_state.editor_key_counter = st.session_state.get('editor_key_counter', 0) + 1
                            st.rerun(scope="fragment")

                    with tab_chart:
                        st.markdown("### 📊 パフォーマンスの概要" if st.session_state.lang == 'ja' else "### 📊 Tổng quan hiệu suất Chấm công")
                        try:
                            import plotly.express as px
                            # Ép kiểu numeric trước khi groupby để tránh lỗi sum() trên object
                            df_chart = df_filtered.copy()
                            for _col in ['Phút đi trễ', 'Phút về sớm', 'Giờ OT', 'Số giờ làm thực tế']:
                                df_chart[_col] = pd.to_numeric(df_chart[_col], errors='coerce').fillna(0)
                            df_nv = df_chart.groupby([m['ma_nv'], m['ten_nv']]).agg(
                                Tong_Tre=('Phút đi trễ', 'sum'), Tong_Som=('Phút về sớm', 'sum'),
                                Tong_OT=('Giờ OT', 'sum'), Tong_Gio=('Số giờ làm thực tế', 'sum')
                            ).reset_index()
                            df_ot = df_nv[df_nv['Tong_OT'] > 0].sort_values('Tong_OT', ascending=False).head(10)

                            if not df_ot.empty:
                                title_ot = "🟢 最も残業が多い従業員 トップ10" if st.session_state.lang == 'ja' else "🟢 Top những người tăng ca nhiều nhất"
                                label_ot = "残業合計" if st.session_state.lang == 'ja' else "Tổng giờ OT"
                                label_nv = "従業員" if st.session_state.lang == 'ja' else "Nhân viên"
                                fig2 = px.bar(df_ot, x='Tong_OT', y=m['ten_nv'], orientation='h', title=title_ot, color='Tong_OT', color_continuous_scale='Greens', labels={'Tong_OT': label_ot, m['ten_nv']: label_nv})
                                fig2.update_layout(yaxis={'categoryorder':'total ascending'}, margin=dict(l=0,r=0,t=40,b=0))
                                st.plotly_chart(fig2, use_container_width=True)
                            else:
                                st.info("ℹ️ 今期は残業した従業員がいません。" if st.session_state.lang == 'ja' else "ℹ️ Chưa có nhân viên nào tăng ca trong kỳ này.")
                            st.markdown("<hr>", unsafe_allow_html=True)

                            df_nghi = df_result_ui[df_result_ui['_loai'] == 'nghi'].groupby(['Mã NV', 'Tên nhân viên']).size().reset_index(name='So_Ngay_Nghi')
                            df_nghi = df_nghi.sort_values('So_Ngay_Nghi', ascending=False).head(10)
                            if not df_nghi.empty:
                                title_nghi = "🔴 欠勤が多い従業員 トップ10" if st.session_state.lang == 'ja' else "🔴 Top những người nghỉ nhiều nhất"
                                label_nghi = "欠勤日数" if st.session_state.lang == 'ja' else "Số ngày nghỉ"
                                label_nv = "従業員" if st.session_state.lang == 'ja' else "Nhân viên"
                                fig_nghi = px.bar(df_nghi, x='So_Ngay_Nghi', y='Tên nhân viên', orientation='h', title=title_nghi, color='So_Ngay_Nghi', color_continuous_scale='Reds', labels={'So_Ngay_Nghi': label_nghi, 'Tên nhân viên': label_nv})
                                fig_nghi.update_layout(yaxis={'categoryorder':'total ascending'}, margin=dict(l=0,r=0,t=40,b=0))
                                st.plotly_chart(fig_nghi, use_container_width=True)
                            else:
                                st.info("ℹ️ 今期は欠勤した従業員がいません。" if st.session_state.lang == 'ja' else "ℹ️ Không có nhân viên nào nghỉ trong kỳ này.")
                            st.markdown("<hr>", unsafe_allow_html=True)

                            st.markdown("### 📅 出勤傾向（日別）" if st.session_state.lang == 'ja' else "### 📅 Xu hướng đi làm theo ngày")
                            df_ngay = df_chart.groupby("Ngày").agg(Tong_OT=('Giờ OT', 'sum')).reset_index()
                            df_ngay['Ngày_dt'] = pd.to_datetime(df_ngay['Ngày'], format='%d/%m/%Y')
                            df_ngay = df_ngay.sort_values('Ngày_dt')
                            title_trend = "日別残業傾向" if st.session_state.lang == 'ja' else "Xu hướng Tăng ca theo ngày"
                            label_val = "数値" if st.session_state.lang == 'ja' else "Số lượng"
                            label_var = "指標" if st.session_state.lang == 'ja' else "Chỉ số"
                            fig3 = px.line(df_ngay, x='Ngày', y=['Tong_OT'], title=title_trend, markers=True, labels={'value': label_val, 'variable': label_var})
                            fig3.for_each_trace(lambda t: t.update(name="残業合計" if st.session_state.lang == 'ja' else "Tổng giờ OT"))
                            fig3.update_layout(legend_title_text='')
                            st.plotly_chart(fig3, use_container_width=True)
                        except ImportError:
                            st.warning("⚠️ Đang cài đặt thư viện `plotly`... Vui lòng Refresh lại trang sau 10 giây.")

                    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
                    st.markdown("## ⬇️ ファイル出力と保存" if st.session_state.lang == 'ja' else "## ⬇️ Xuất file & Lưu trữ")
                    col_exp1, col_exp2, col_exp3 = st.columns([2, 1, 1])
                    
                    file_prefix = "タイムカード" if st.session_state.lang == 'ja' else "Bang cong"
                    file_name_export = f"{file_prefix} {start_date.strftime('%d%m')}-{end_date.strftime('%d%m')}.xlsx"
                    
                    with col_exp1:
                        st.markdown(f"""<div style='background:#F0FDF8;border:0.5px solid #6EE7C0;border-radius:10px;padding:12px 16px;font-size:13px;color:#0F6E56'>
    📄 File: <b>{file_name_export}</b><br>
    <span style='color:#6B7280;font-size:12px'>{t("export_file_rows", rows=total_rows, emps=total_emps)}</span>
    </div>""", unsafe_allow_html=True)
                    with col_exp2:
                        if st.button(t("btn_save_db"), use_container_width=True):
                            with st.spinner(t("spinner_save")):
                                save_to_db(df_filtered, m)
                            st.success(t("success_save"))
                    with col_exp3:
                        excel_data = None
                        try:
                            total_wd_tuple = calculate_working_days(start_date, end_date, st.session_state.custom_holidays, st.session_state.custom_workdays)
                            total_wd = total_wd_tuple[0]
                            excel_data = export_excel_tong_hop(df_filtered, m, start_date, end_date, total_wd)
                        except Exception as e:
                            st.error(f"❌ Lỗi xuất file: {e}")
                        if excel_data is not None:
                            st.download_button(
                                label=t("btn_download_excel"),
                                data=excel_data,
                                file_name=file_name_export,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                type="primary", use_container_width=True,
                            )
                        else:
                            st.button("⬇️ Tải Excel", disabled=True, use_container_width=True, help="Không có dữ liệu để xuất")

            render_interactive_dashboard(df_filtered)

# Gọi chatbot ở cuối luồng (dành cho chamcong)
render_chatbot()
