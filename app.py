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

DB_FILE = 'attendance.db'

def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS records (
            ma_nv TEXT,
            ten_nv TEXT,
            ngay TEXT,
            gio_vao TEXT,
            gio_ra TEXT,
            di_tre INTEGER,
            ve_som INTEGER,
            ot REAL,
            tong_gio REAL,
            ghi_chu TEXT,
            UNIQUE(ma_nv, ngay)
        )
    ''')
    conn.commit()
    conn.close()

init_db()

def save_to_db(df_filtered, mapping):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    records = []
    for _, row in df_filtered.iterrows():
        records.append((
            str(row[mapping['ma_nv']]),
            str(row[mapping['ten_nv']]),
            str(row["Ngày"]),
            str(row[mapping['gio_vao']]),
            str(row[mapping['gio_ra']]),
            int(row.get("Phút đi trễ", 0)),
            int(row.get("Phút về sớm", 0)),
            float(row.get("Giờ OT", 0)),
            float(row.get("Số giờ làm thực tế", 0)),
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
# 1. CẤU HÌNH VÀ HÀM TIỆN ÍCH
# ==========================================
APP_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(APP_DIR, "assets")
FAVICON_PATH = os.path.join(ASSETS_DIR, "favicon.png")
LOGO_HEADER_PATH = os.path.join(ASSETS_DIR, "logo_header.png")

def load_favicon():
    """Trả về ảnh favicon (logo công ty); nếu thiếu file thì dùng emoji mặc định."""
    try:
        return PILImage.open(FAVICON_PATH)
    except Exception:
        return "📊"

def load_logo_base64():
    """Đọc logo header và encode base64 để nhúng trực tiếp vào HTML; trả về None nếu thiếu file."""
    try:
        with open(LOGO_HEADER_PATH, "rb") as f:
            return base64.b64encode(f.read()).decode("utf-8")
    except Exception:
        return None

LOGO_HEADER_B64 = load_logo_base64()

st.set_page_config(page_title="Bảng Chấm Công", page_icon=load_favicon(), layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Be+Vietnam+Pro:wght@400;500;600;700;800&family=Inter:wght@400;500;600;700&display=swap');

:root {
    --brand-50:  #ECFDF5;
    --brand-100: #D1FAE5;
    --brand-400: #34D399;
    --brand-500: #10B981;
    --brand-600: #059669;
    --brand-700: #047857;
    --blue-500:  #3B82F6;
    --ink-900:   #0F172A;
    --ink-700:   #334155;
    --ink-500:   #64748B;
    --line:      #E2E8F0;
    --bg-color:  #F8FAFC;
    --card-shadow: 0 4px 20px rgba(0, 0, 0, 0.05);
}

html, body, .stApp {
    font-family: 'Inter', 'Be Vietnam Pro', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    background-color: var(--bg-color);
}

/* Ẩn footer mặc định của Streamlit */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

/* Nền tổng thể */
.stApp {
    background: var(--bg-color);
}

/* Padding tổng thể */
.block-container {
    padding-top: 1.5rem !important;
    padding-left: 2rem !important;
    padding-right: 2rem !important;
    max-width: 1180px !important;
}

/* ===== Header app ===== */
.app-header {
    background: linear-gradient(135deg, var(--brand-500) 0%, var(--blue-500) 100%);
    border-radius: 20px;
    padding: 24px 32px;
    display: flex;
    align-items: center;
    gap: 20px;
    margin-bottom: 24px;
    box-shadow: 0 10px 30px rgba(16, 185, 129, 0.25);
    color: white;
}
.app-header-icon {
    font-size: 32px;
    display: flex;
    align-items: center;
    justify-content: center;
    background: rgba(255, 255, 255, 0.2);
    backdrop-filter: blur(10px);
    border-radius: 16px;
    padding: 12px 16px;
    border: 1px solid rgba(255, 255, 255, 0.3);
}
.app-header-icon img {
    height: 38px;
    width: auto;
    display: block;
}
.app-header-icon-emoji {
    font-size: 30px;
}
.app-header-title {
    font-size: 22px;
    font-weight: 700;
    margin: 0;
    letter-spacing: -0.01em;
}
.app-header-sub {
    font-size: 14px;
    color: rgba(255, 255, 255, 0.85);
    margin: 4px 0 0 0;
}
.app-header-badge {
    margin-left: auto;
    background: rgba(255, 255, 255, 0.2);
    border: 1px solid rgba(255, 255, 255, 0.4);
    color: white;
    font-size: 13px;
    font-weight: 600;
    padding: 8px 16px;
    border-radius: 100px;
}

/* ===== Card container ===== */
.card {
    background: #FFFFFF;
    border: 1px solid var(--line);
    border-radius: 16px;
    padding: 22px 26px;
    margin-bottom: 18px;
    box-shadow: var(--card-shadow);
}
.card-title {
    font-size: 15px;
    font-weight: 700;
    color: var(--ink-900);
    margin: 0 0 16px 0;
    display: flex;
    align-items: center;
    gap: 12px;
    letter-spacing: -0.01em;
}
.card-icon {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 32px;
    height: 32px;
    font-size: 16px;
    border-radius: 10px;
    background: var(--brand-50);
    color: var(--brand-600);
}

/* Upload zone */
.upload-hint {
    background-color: #ECFDF5;
    border-radius: 12px;
    padding: 12px 18px;
    font-size: 14px;
    color: #065F46;
    font-weight: 500;
    margin-bottom: 16px;
    border: 1px solid #A7F3D0;
}

/* Nút bấm tổng quát */
.stButton > button,
.stDownloadButton > button {
    background: linear-gradient(135deg, var(--brand-500), var(--blue-500)) !important;
    color: white !important;
    border: none !important;
    border-radius: 12px !important;
    padding: 12px 24px !important;
    font-size: 14.5px !important;
    font-weight: 600 !important;
    box-shadow: 0 4px 14px rgba(16, 185, 129, 0.3) !important;
    transition: all 0.2s ease !important;
}
.stButton > button:hover,
.stDownloadButton > button:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 6px 18px rgba(16, 185, 129, 0.4) !important;
    filter: brightness(1.05) !important;
}
.stButton > button:active,
.stDownloadButton > button:active {
    transform: translateY(0) !important;
}

/* Hiển thị rõ focus để hỗ trợ điều hướng bàn phím */
.stButton > button:focus-visible,
.stDownloadButton > button:focus-visible {
    outline: 2px solid var(--blue-500) !important;
    outline-offset: 2px !important;
}

/* Streamlit file uploader */
[data-testid="stFileUploader"] {
    background: #FFFFFF;
    border-radius: 16px;
    padding: 12px;
    border: 1px dashed var(--brand-400);
    box-shadow: var(--card-shadow);
}
[data-testid="stFileUploaderDropzone"] {
    padding: 20px !important;
    border-radius: 12px !important;
}

/* Dataframe */
[data-testid="stDataFrame"] {
    border-radius: 16px;
    overflow: hidden;
    border: 1px solid var(--line);
    box-shadow: var(--card-shadow);
}

/* Date & Time input */
[data-testid="stDateInput"] input,
[data-testid="stTimeInput"] input,
[data-testid="stNumberInput"] input {
    border-radius: 10px !important;
    border: 1px solid var(--line) !important;
}

/* Expander (tuỳ chỉnh giờ làm chuẩn) */
[data-testid="stExpander"] {
    border: 1px solid var(--line) !important;
    border-radius: 16px !important;
    background: #FFFFFF !important;
    box-shadow: var(--card-shadow) !important;
}

/* Metric widget của Streamlit */
[data-testid="stMetric"] {
    background: #FFFFFF;
    border: 1px solid var(--line);
    border-radius: 16px;
    padding: 18px 20px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.02);
    border-left: 4px solid var(--brand-500);
}
[data-testid="stMetricLabel"] {
    font-size: 13px !important;
    color: var(--ink-500) !important;
    font-weight: 500 !important;
}
[data-testid="stMetricValue"] {
    font-family: 'Inter', 'Be Vietnam Pro', sans-serif !important;
    font-size: 26px !important;
    font-weight: 700 !important;
    color: var(--ink-900) !important;
}

/* Spinner */
.stSpinner > div {
    border-top-color: var(--brand-500) !important;
}

/* Alert / Warning */
.stAlert {
    border-radius: 12px !important;
    border: none !important;
    box-shadow: 0 2px 8px rgba(0,0,0,0.03) !important;
}

/* Selectbox */
[data-testid="stSelectbox"] > div {
    border-radius: 10px !important;
}

/* Multiselect (lọc nhân viên) */
[data-testid="stMultiSelect"] > div > div {
    border-radius: 10px !important;
}
[data-baseweb="tag"] {
    background-color: var(--brand-50) !important;
    color: var(--brand-700) !important;
    border-radius: 6px !important;
}

/* Subheader / Heading thuần markdown */
.stSubheader, h3 {
    font-size: 16px !important;
    font-weight: 700 !important;
    color: var(--ink-900) !important;
    letter-spacing: -0.01em !important;
}
</style>
""", unsafe_allow_html=True)

def to_time_obj(val):
    if pd.isna(val) or val == "":
        return None
    if isinstance(val, datetime.time):
        return val
    if isinstance(val, datetime.datetime):
        return val.time()
    if isinstance(val, (int, float)):
        # Xử lý format serial time của Excel
        if 0 <= val < 1:
            total_seconds = int(val * 86400)
            hours = total_seconds // 3600
            minutes = (total_seconds % 3600) // 60
            seconds = total_seconds % 60
            return datetime.time(hour=hours, minute=minutes, second=seconds)
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
    """Trả về số float được làm tròn đúng 1 chữ số thập phân."""
    if val is None or (isinstance(val, float) and math.isnan(val)) or str(val).lower() in ['nan', 'nat', 'none', '']:
        return 0.0
    return round(float(val), 1)

def format_gio_lam_str(val):
    """Trả về string để hiển thị luôn có 1 số thập phân (VD: '8.0', '7.5')."""
    return f"{format_gio_lam(val):.1f}"

# ==========================================
# 2. LOGIC TÍNH TOÁN (CALCULATOR)
# ==========================================
def calc_overlap(s1, e1, s2, e2):
    return max(0.0, min(e1, e2) - max(s1, s2))

def calculate_working_hours(time_in, time_out, start_chuan=8.0, end_chuan=17.0, lunch_start=12.0, lunch_end=13.0, max_hours=8.0):
    vao_trong = time_in is None or str(time_in).strip().lower() in ['', 'nan', 'none', 'nat']
    ra_trong = time_out is None or str(time_out).strip().lower() in ['', 'nan', 'none', 'nat']
    
    result = {'admin_hours': 0.0, 'tong_gio': 0.0, 'di_tre': 0, 've_som': 0, 'ot': 0.0}
    
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

    # 1. Giờ làm hành chính (chỉ tính trong khoảng giờ chuẩn)
    eff_in = max(in_f, start_chuan)
    eff_out = min(out_f, end_chuan)
    
    if eff_out > eff_in:
        total_admin = eff_out - eff_in
        lunch_overlap = calc_overlap(eff_in, eff_out, lunch_start, lunch_end)
        admin_hours = total_admin - lunch_overlap
        admin_hours = min(max(admin_hours, 0.0), max_hours)
    else:
        admin_hours = 0.0
    result['admin_hours'] = admin_hours

    # 2. Phút đi trễ / về sớm (chỉ xét trong khoảng giờ chuẩn)
    if in_f > start_chuan and in_f < end_chuan:
        tre = (in_f - start_chuan) - calc_overlap(start_chuan, in_f, lunch_start, lunch_end)
        result['di_tre'] = int(tre * 60)
        
    if out_f < end_chuan and out_f > start_chuan:
        som = (end_chuan - out_f) - calc_overlap(out_f, end_chuan, lunch_start, lunch_end)
        result['ve_som'] = int(som * 60)
        
    # 3. Tính OT: Số giờ ngoài khung
    ot_hours = 0.0
    if out_f > end_chuan:
        ot_hours += (out_f - end_chuan)
    if in_f < start_chuan:
        ot_hours += (start_chuan - in_f)
    result['ot'] = max(0.0, ot_hours)
    
    # 4. Tổng giờ = Chỉ tính giờ hành chính (theo yêu cầu)
    result['tong_gio'] = admin_hours
    
    return result

def clean_date(val):
    if pd.isna(val): return None
    val_str = str(val).strip()
    try:
        if isinstance(val, datetime.datetime):
            return val
        # Parse ngày (ưu tiên format ngày/tháng/năm của VN)
        return pd.to_datetime(val_str, dayfirst=True)
    except:
        return None

def parse_holiday_dates(text):
    """
    Phân tích chuỗi ngày nghỉ lễ do người dùng nhập (phân tách bởi dấu phẩy
    hoặc xuống dòng, định dạng dd/mm/yyyy) thành 1 set datetime.date.
    Trả về (set_ngay_le, danh_sach_chuoi_loi).
    """
    if not text or not text.strip():
        return set(), []
    holidays = set()
    invalid = []
    raw_items = [x.strip() for x in text.replace("\n", ",").split(",") if x.strip()]
    for item in raw_items:
        try:
            holidays.add(pd.to_datetime(item, dayfirst=True).date())
        except Exception:
            invalid.append(item)
    return holidays, invalid

# Các ngày lễ dương lịch cố định theo Luật Lao động VN (tháng, ngày)
# Lưu ý: Tết Nguyên đán, Giỗ Tổ Hùng Vương (10/3 âm lịch) đổi theo năm
# nên KHÔNG đưa vào đây — người dùng tự bổ sung thêm vào ô nhập.
FIXED_VN_HOLIDAYS = [
    (1, 1),   # Tết Dương lịch
    (4, 30),  # Ngày Giải phóng miền Nam
    (5, 1),   # Quốc tế Lao động
    (9, 2),   # Quốc khánh
]

def get_fixed_holidays_for_years(years):
    """Trả về list datetime.date các ngày lễ cố định dương lịch cho các năm trong `years`."""
    dates = []
    for y in sorted(years):
        for (mo, day) in FIXED_VN_HOLIDAYS:
            dates.append(datetime.date(y, mo, day))
    return sorted(dates)

def calculate_working_days(start_date, end_date, holidays=None):
    """
    Đếm số ngày công trong khoảng [start_date, end_date] theo quy tắc:
    - Thứ 2 - Thứ 6: ngày làm việc, trừ khi rơi vào ngày nghỉ lễ.
    - Thứ 7 cuối cùng của tháng: ngày làm việc BẮT BUỘC, trừ khi rơi vào ngày nghỉ lễ.
    - Thứ 7 (không phải cuối tháng) và Chủ nhật: luôn là ngày nghỉ, không tính công.

    Trả về (so_ngay_cong, danh_sach_ngay_le_trong_ky)
    - so_ngay_cong: tổng số ngày công sau khi đã cộng Thứ 7 cuối tháng và trừ ngày lễ.
    - danh_sach_ngay_le_trong_ky: list dict {date, weekday_label, is_workday}
      với is_workday = True nếu ngày lễ đó trùng vào ngày làm việc (đã bị trừ khỏi ngày công).
    """
    if holidays is None:
        holidays = set()

    weekday_names = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6", "Thứ 7", "Chủ nhật"]
    total = 0
    holiday_info = []
    current = start_date
    one_day = datetime.timedelta(days=1)
    while current <= end_date:
        wd = current.weekday()  # 0 = Thứ 2 ... 5 = Thứ 7, 6 = Chủ nhật
        # Thứ 7 là "cuối tháng" nếu 7 ngày sau đó đã sang tháng khác
        is_last_saturday = (wd == 5) and (current + datetime.timedelta(days=7)).month != current.month
        is_workday_by_default = (wd < 5) or is_last_saturday  # Thứ 2-6 hoặc Thứ 7 cuối tháng

        if current in holidays:
            holiday_info.append({
                "date": current,
                "weekday_label": weekday_names[wd],
                "is_workday": is_workday_by_default,
            })
            # Ngày lễ trùng ngày làm việc -> không cộng công (coi như đã trừ)
        elif is_workday_by_default:
            total += 1

        current += one_day

    return total, holiday_info

# ==========================================
# 3. LOGIC PARSER ĐỌC FILE
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
        # Đếm số keyword xuất hiện trong dòng
        if sum(1 for kw in keywords if any(kw in cell for cell in row_str)) >= 3:
            return i
    return 0  # fallback

def parse_excel_file(uploaded_file):
    file_name = uploaded_file.name.lower()
    
    try:
        header_row = find_header_row(uploaded_file, file_name)
    except Exception as e:
        st.error(f"❌ Lỗi khi đọc file bằng thư viện Pandas: {e}")
        st.stop()
        
    uploaded_file.seek(0)

    try:
        if file_name.endswith(('.csv', '.txt', '.dat', '.tsv')):
            df = pd.read_csv(uploaded_file, header=header_row, sep=None, engine='python')
        else:
            df = pd.read_excel(uploaded_file, header=header_row)
    except Exception as e:
        st.error(f"❌ Lỗi khi đọc bảng dữ liệu ở dòng {header_row}: {e}")
        st.stop()
        
    # Bỏ các dòng hoàn toàn trống
    df = df.dropna(how='all')
    df.columns = df.columns.astype(str).str.strip()
    return df

def auto_detect_columns(df):
    mapping = {}
    for col in df.columns:
        col_lower = str(col).strip().lower()
        if 'mã' in col_lower and 'nv' not in col_lower.replace('mã',''):
            mapping['ma_nv'] = col
        elif col_lower == 'tên nhân viên' or col_lower == 'tên nv' \
             or col_lower == 'họ tên' or col_lower == 'tên':
            mapping['ten_nv'] = col
        elif 'ngày' in col_lower and 'ca' not in col_lower:
            mapping['ngay'] = col
        elif col_lower == 'vào' or col_lower == 'vao':
            mapping['gio_vao'] = col
        elif col_lower == 'ra' and 'sớm' not in col_lower \
             and 'som' not in col_lower:
            mapping['gio_ra'] = col
    return mapping

# ==========================================
# 4. LOGIC EXPORTER XUẤT FILE
# ==========================================
def export_excel_tong_hop(uploaded_file, df_filtered, mapping, max_hours=8.0):
    """
    Xuất file Báo cáo 2 Tab: Tab 1 là dữ liệu gốc chèn thêm cột, Tab 2 là báo cáo phân tích mới.
    """
    from openpyxl.utils.dataframe import dataframe_to_rows
    
    uploaded_file.seek(0)
    try:
        wb = openpyxl.load_workbook(uploaded_file)
    except Exception:
        wb = openpyxl.Workbook()
        wb.active.title = "Dữ liệu gốc"

    ws = wb.active
    
    # === TAB 1: GỘP CỘT VÀO DỮ LIỆU GỐC ===
    target_ma = str(mapping['ma_nv']).strip().lower()
    target_ngay = str(mapping['ngay']).strip().lower()
    target_ra = str(mapping['gio_ra']).strip().lower()

    header_row_idx, col_ra_idx, col_ngay_idx, col_ma_nv_idx, col_gio_idx = None, None, None, None, None
    for row_idx, row in enumerate(ws.iter_rows(max_row=100), start=1):
        row_vals = [str(c.value).strip().lower() if c.value else '' for c in row]
        if target_ra in row_vals:
            header_row_idx = row_idx
            for ci, val in enumerate(row_vals, start=1):
                if val == target_ma: col_ma_nv_idx = ci
                elif val == target_ngay: col_ngay_idx = ci
                elif val == target_ra: col_ra_idx = ci
                elif val in ['giờ', 'gio']: col_gio_idx = ci
            break
            
    if header_row_idx and col_ra_idx:
        insert_col = col_gio_idx if col_gio_idx else col_ra_idx + 1
        ws.insert_cols(insert_col)
        
        header_cell_ra = ws.cell(row=header_row_idx, column=col_ra_idx)
        header_cell_new = ws.cell(row=header_row_idx, column=insert_col)
        header_cell_new.value = "Giờ làm thực tế"
        header_cell_new.font = copy(header_cell_ra.font)
        header_cell_new.fill = copy(header_cell_ra.fill)
        header_cell_new.alignment = copy(header_cell_ra.alignment)
        header_cell_new.border = copy(header_cell_ra.border)
        
        lookup = {}
        for _, row in df_filtered.iterrows():
            ma = str(row[mapping['ma_nv']]).strip().upper()
            if ma.endswith('.0'): ma = ma[:-2]
            try:
                if hasattr(row['_parsed_date'], 'date'): ngay_key = row['_parsed_date'].strftime('%d/%m/%Y')
                else: ngay_key = str(row[mapping['ngay']]).strip()
            except: ngay_key = str(row[mapping['ngay']]).strip()
            gio = row['Số giờ làm thực tế']
            lookup[(ma, ngay_key)] = gio
            
        for row_idx in range(header_row_idx + 1, ws.max_row + 1):
            cell_ngay = ws.cell(row=row_idx, column=col_ngay_idx)
            cell_ma   = ws.cell(row=row_idx, column=col_ma_nv_idx)
            cell_new  = ws.cell(row=row_idx, column=insert_col)
            
            cell_ra_ref = ws.cell(row=row_idx, column=col_ra_idx)
            cell_new.font = copy(cell_ra_ref.font)
            cell_new.alignment = copy(cell_ra_ref.alignment)
            cell_new.border = copy(cell_ra_ref.border)
            
            ma_val = str(cell_ma.value).strip().upper() if cell_ma.value else ''
            if ma_val.endswith('.0'): ma_val = ma_val[:-2]
            ngay_val = cell_ngay.value
            
            if ngay_val and ma_val:
                try:
                    if hasattr(ngay_val, 'strftime'): ngay_str = ngay_val.strftime('%d/%m/%Y')
                    else:
                        ngay_str = str(ngay_val).strip()
                        for fmt in ['%d/%m/%Y', '%Y-%m-%d', '%m/%d/%Y']:
                            try:
                                ngay_str = datetime.datetime.strptime(ngay_str[:10], fmt).strftime('%d/%m/%Y')
                                break
                            except: continue
                except: ngay_str = ''
                
                gio = lookup.get((ma_val, ngay_str), '')
                cell_new.value = gio if gio != '' else 0.0
                cell_new.number_format = '0.0'
            else:
                cell_new.value = ''

    # === TAB 2: BÁO CÁO PHÂN TÍCH ===
    def get_col(df, possible_names, default=''):
        for name in possible_names:
            for c in df.columns:
                if str(c).strip().lower() == name.lower(): return df[c].values
        return [default] * len(df)
        
    thu_map = {0:'Hai', 1:'Ba', 2:'Tư', 3:'Năm', 4:'Sáu', 5:'Bảy', 6:'CN'}
    thu_series = df_filtered["_parsed_date"].dt.weekday.map(thu_map)
    gio_hc = pd.to_numeric(df_filtered["Giờ hành chính"], errors='coerce').fillna(0)
    cong = (gio_hc / max_hours).round(2)
    
    df_export = pd.DataFrame({
        "Mã N.Viên": df_filtered[mapping['ma_nv']],
        "Tên nhân viên": df_filtered[mapping['ten_nv']],
        "Phòng ban": get_col(df_filtered, ['phòng ban', 'phong ban']),
        "Chức vụ": get_col(df_filtered, ['chức vụ', 'chuc vu']),
        "Ngày": df_filtered["Ngày"],
        "Thứ": thu_series,
        "Vào": df_filtered[mapping['gio_vao']],
        "Ra": df_filtered[mapping['gio_ra']],
        "Công": cong,
        "Giờ": df_filtered["Giờ hành chính"],
        "Công+": 0, "Giờ+": 0,
        "Vào Trễ": df_filtered["Phút đi trễ"],
        "Ra sớm": df_filtered["Phút về sớm"],
        "TC1": df_filtered["Giờ OT"],
        "TC2": 0, "TC3": 0,
        "Tên ca": get_col(df_filtered, ['tên ca', 'ten ca', 'ca làm việc'], 'HC'),
        "Kí hiệu": get_col(df_filtered, ['kí hiệu', 'ký hiệu', 'ki hieu'], 'X'),
        "Kí hiệu+": '',
        "Tổng giờ": df_filtered["Số giờ làm thực tế"],
        "Ghi chú (Hệ thống)": df_filtered["Ghi chú"]
    })
    
    ws2 = wb.create_sheet(title='Chi tiết phân tích')
    for r in dataframe_to_rows(df_export, index=False, header=True):
        ws2.append(r)
        
    header_font = Font(bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    for cell in ws2[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
        
    for row in ws2.iter_rows(min_row=2):
        # Format numeric columns to 1 decimal place
        # Columns: Công(9), Giờ(10), Công+(11), Giờ+(12), Vào Trễ(13), Ra Sớm(14), TC1(15), TC2(16), TC3(17), Tổng giờ(20)
        numeric_cols = [9, 10, 11, 12, 13, 14, 15, 16, 17, 20]
        for col_idx in numeric_cols:
            if row[col_idx - 1].value is not None:
                try:
                    row[col_idx - 1].value = float(row[col_idx - 1].value)
                    row[col_idx - 1].number_format = '0.0'
                except:
                    pass
                    
        ghi_chu_cell = row[-1]
        if ghi_chu_cell.value and "⚠️" in str(ghi_chu_cell.value):
            for cell in row:
                cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()

# ==========================================
# 5. GIAO DIỆN CHÍNH (UI)
# ==========================================
header_placeholder = st.empty()

def render_header(badge_text=None):
    badge_html = f'<div class="app-header-badge">{badge_text}</div>' if badge_text else ''
    if LOGO_HEADER_B64:
        icon_html = f'<img src="data:image/png;base64,{LOGO_HEADER_B64}" alt="VIET.MOS logo">'
    else:
        icon_html = '<span class="app-header-icon-emoji">📊</span>'
    header_placeholder.markdown(f"""
    <div class="app-header">
        <div class="app-header-icon">{icon_html}</div>
        <div>
            <p class="app-header-title">Ứng dụng Bảng Chấm Công</p>
            <p class="app-header-sub">Tính giờ làm việc thực tế từ file Excel máy chấm công</p>
        </div>
        {badge_html}
    </div>
    """, unsafe_allow_html=True)

render_header()

# Khởi tạo state
if "step" not in st.session_state:
    st.session_state.step = 1
if "df_raw" not in st.session_state:
    st.session_state.df_raw = None

# ----------------- PHẦN A: UPLOAD & ĐIỀU HƯỚNG -----------------
st.sidebar.markdown("## ⚙️ Hệ thống Quản trị")
app_mode = st.sidebar.radio("Chế độ hiển thị", ["Xử lý file chấm công", "Lịch sử Lưu trữ"], label_visibility="collapsed")

if app_mode == "Lịch sử Lưu trữ":
    st.markdown("## 🗄️ Lịch sử Chấm công")
    conn = sqlite3.connect(DB_FILE)
    df_db = pd.read_sql_query("SELECT * FROM records ORDER BY ma_nv, ngay DESC", conn)
    conn.close()
    
    if df_db.empty:
        st.info("Chưa có dữ liệu nào được lưu. Hãy xử lý file chấm công và bấm 'Lưu dữ liệu vào Hệ thống'.")
    else:
        st.dataframe(df_db, use_container_width=True)
    st.stop()

st.sidebar.markdown("### 📁 Bước 1: Upload file chấm công")
st.sidebar.markdown('<div class="upload-hint">✅ Hỗ trợ: <b>.xlsx, .xls, .csv, .txt, .dat, .tsv</b></div>', unsafe_allow_html=True)
uploaded_file = st.sidebar.file_uploader("Chọn file chấm công", type=["xlsx", "xls", "csv", "txt", "dat", "tsv", "xlsm", "xlsb"], label_visibility="collapsed")

if uploaded_file is not None:
    if st.session_state.df_raw is None or st.session_state.get('last_uploaded') != uploaded_file.name:
        with st.spinner("Đang đọc file..."):
            st.session_state['uploaded_file_bytes'] = uploaded_file.read()
            uploaded_file.seek(0)
            df = parse_excel_file(uploaded_file)
            if df is not None and not df.empty:
                st.session_state.df_raw = df
                st.session_state.last_uploaded = uploaded_file.name
                
                # Tự động map cột
                mapping_auto = auto_detect_columns(df)
                if len(mapping_auto) == 5:
                    st.session_state.mapping = mapping_auto
                    st.session_state.step = 3
                    st.rerun()
                else:
                    st.session_state.step = 2
                    st.session_state.mapping_auto = mapping_auto
elif st.session_state.df_raw is not None:
    # Người dùng đã xóa/gỡ file khỏi ô upload -> xóa sạch dữ liệu & kết quả đang hiển thị
    st.session_state.df_raw = None
    st.session_state.last_uploaded = None
    st.session_state.step = 1
    for _key in ("mapping", "mapping_auto"):
        if _key in st.session_state:
            del st.session_state[_key]
    st.rerun()

# ----------------- PHẦN B: MAPPING CỘT (Nếu thiếu) -----------------
if st.session_state.step == 2 and st.session_state.df_raw is not None:
    df_raw = st.session_state.df_raw
    mapping_auto = st.session_state.mapping_auto
    
    st.markdown("### Dữ liệu thô (10 dòng đầu)")
    st.dataframe(df_raw.head(10), use_container_width=True, hide_index=True)
    
    st.sidebar.markdown("### 🔗 Bước 2: Xác nhận mapping cột")
    st.sidebar.warning("⚠️ Không nhận diện đủ 5 cột.")
    
    map_emp_id = st.sidebar.selectbox("Mã NV", columns, index=get_index(mapping_auto.get('ma_nv')))
    map_emp_name = st.sidebar.selectbox("Tên NV", columns, index=get_index(mapping_auto.get('ten_nv')))
    map_date = st.sidebar.selectbox("Ngày làm việc", columns, index=get_index(mapping_auto.get('ngay')))
    map_in = st.sidebar.selectbox("Giờ vào", columns, index=get_index(mapping_auto.get('gio_vao')))
    map_out = st.sidebar.selectbox("Giờ ra", columns, index=get_index(mapping_auto.get('gio_ra')))
        
    if st.sidebar.button("Xác nhận Mapping", type="primary", use_container_width=True):
        if "-- Chọn cột --" in [map_emp_id, map_emp_name, map_date, map_in, map_out]:
            st.sidebar.error("❌ Vui lòng chọn đủ 5 cột!")
        else:
            st.session_state.mapping = {
                'ma_nv': map_emp_id, 'ten_nv': map_emp_name, 
                'ngay': map_date, 'gio_vao': map_in, 'gio_ra': map_out
            }
            st.session_state.step = 3
            st.rerun()

# ----------------- PHẦN C & D: LỌC & KẾT QUẢ -----------------
if st.session_state.step >= 3 and "mapping" in st.session_state:
    m = st.session_state.mapping
    df_process = st.session_state.df_raw.copy()
    
    # 1. Làm sạch sơ bộ
    df_process = df_process.dropna(subset=[m['ma_nv']])
    df_process = df_process[df_process[m['ma_nv']].astype(str).str.strip() != '']
    
    # 2. Xử lý cột ngày
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

        st.sidebar.markdown("### 📅 Bước 3: Chọn kỳ công")
        col_d1, col_d2 = st.sidebar.columns(2)
        with col_d1:
            start_date = st.date_input("Từ ngày", default_start)
        with col_d2:
            end_date = st.date_input("Đến ngày", default_end)
            
        # Cập nhật Badge trên Header linh động theo ngày
        render_header(f"Kỳ công {start_date.strftime('%d/%m')} – {end_date.strftime('%d/%m/%Y')}")

        # Ngày lễ cố định dương lịch (01/01, 30/04, 01/05, 02/09) được áp dụng NGẦM
        # cho đúng (các) năm nằm trong kỳ công đã chọn — không cần hiển thị/nhập tay.
        years_in_range = range(start_date.year, end_date.year + 1)
        fixed_holidays = set(get_fixed_holidays_for_years(years_in_range))

        with st.sidebar.expander("📅 Ngày lễ âm lịch & bù (nhập tay)"):
            st.markdown("""
            <small>
            Định dạng <b>dd/mm/yyyy</b>, mỗi ngày 1 dòng.
            </small>
            """, unsafe_allow_html=True)
            
            ngay_le_thu_cong_raw = st.text_area(
                "Danh sách ngày",
                placeholder="Ví dụ:\n06/04/2026\n28/01/2026\n29/01/2026\n02/05/2026",
                height=100,
                label_visibility="collapsed"
            )
            
            # Parse
            NGAY_LE_THU_CONG = set()
            if ngay_le_thu_cong_raw.strip():
                loi = []
                for line in ngay_le_thu_cong_raw.strip().splitlines():
                    line = line.strip()
                    if not line:
                        continue
                    try:
                        d = datetime.datetime.strptime(line, '%d/%m/%Y').date()
                        NGAY_LE_THU_CONG.add(d)
                    except:
                        loi.append(line)
                
                if loi:
                    st.warning(f"⚠️ Không đọc được: {', '.join(loi)} "
                               "(cần định dạng dd/mm/yyyy)")
                elif NGAY_LE_THU_CONG:
                    st.success(f"✅ Đã thêm {len(NGAY_LE_THU_CONG)} ngày "
                               "vào danh sách nghỉ.")
                               
        holiday_dates = fixed_holidays | NGAY_LE_THU_CONG

        working_days, holiday_info = calculate_working_days(start_date, end_date, holiday_dates)

        if holiday_info:
            st.markdown('<div class="card"><div class="card-title"><span class="card-icon">📅</span>Ngày lễ trong kỳ</div></div>', unsafe_allow_html=True)
            for h in sorted(holiday_info, key=lambda x: x["date"]):
                date_str = h["date"].strftime("%d/%m/%Y")
                if h["is_workday"]:
                    st.warning(f"⚠️ **{date_str} ({h['weekday_label']})** trùng vào ngày làm việc — đã trừ 1 ngày công khỏi kỳ này.")
                else:
                    st.info(f"ℹ️ {date_str} ({h['weekday_label']}) rơi vào ngày nghỉ cuối tuần — không ảnh hưởng đến số ngày công.")

        with st.sidebar.expander("⚙️ Tuỳ chỉnh giờ làm chuẩn"):
            col_h1, col_h2 = st.columns(2)
            with col_h1:
                gio_vao_chuan = st.time_input("Giờ vào", datetime.time(8, 0))
            with col_h2:
                gio_ra_chuan = st.time_input("Giờ ra", datetime.time(17, 0))
                
            col_h3, col_h4 = st.columns(2)
            with col_h3:
                nghi_trua_bat_dau = st.time_input("Nghỉ trưa từ", datetime.time(12, 0))
            with col_h4:
                nghi_trua_ket_thuc = st.time_input("Nghỉ trưa đến", datetime.time(13, 0))
                
            so_gio_toi_da = st.number_input("Tối đa giờ/ngày", min_value=0.0, max_value=24.0, value=8.0, step=0.5)

        if time_to_float(gio_ra_chuan) <= time_to_float(gio_vao_chuan):
            st.sidebar.error("❌ Giờ ra chuẩn phải sau giờ vào chuẩn.")
            st.stop()
        if time_to_float(nghi_trua_ket_thuc) < time_to_float(nghi_trua_bat_dau):
            st.sidebar.error("❌ Giờ nghỉ trưa không hợp lệ.")
            st.stop()

        mask = (df_process["_parsed_date"].dt.date >= start_date) & (df_process["_parsed_date"].dt.date <= end_date)
        df_filtered = df_process.loc[mask].copy()
        
        if df_filtered.empty:
            st.warning("Không có dữ liệu trong khoảng thời gian đã chọn.")
        else:
            with st.spinner("⏳ Đang tính toán giờ làm..."):
                # Tính giờ (theo giờ chuẩn đã tuỳ chỉnh ở trên)
                df_calc = df_filtered.apply(
                    lambda row: calculate_working_hours(
                        row[m['gio_vao']], row[m['gio_ra']],
                        start_chuan=time_to_float(gio_vao_chuan),
                        end_chuan=time_to_float(gio_ra_chuan),
                        lunch_start=time_to_float(nghi_trua_bat_dau),
                        lunch_end=time_to_float(nghi_trua_ket_thuc),
                        max_hours=so_gio_toi_da,
                    ), axis=1
                )
                
                df_filtered["Giờ hành chính"] = df_calc.apply(lambda x: x['admin_hours'] if isinstance(x, dict) else 0.0)
                df_filtered["Số giờ làm thực tế"] = df_calc.apply(lambda x: x['tong_gio'] if isinstance(x, dict) else 0.0)
                df_filtered["Phút đi trễ"] = df_calc.apply(lambda x: x['di_tre'] if isinstance(x, dict) else 0)
                df_filtered["Phút về sớm"] = df_calc.apply(lambda x: x['ve_som'] if isinstance(x, dict) else 0)
                df_filtered["Giờ OT"] = df_calc.apply(lambda x: x['ot'] if isinstance(x, dict) else 0.0)
                
                def check_anomaly(row):
                    vao = row[m['gio_vao']]
                    ra = row[m['gio_ra']]
                    vao_trong = pd.isna(vao) or str(vao).strip().lower() in ['', 'nan', 'none', 'nat']
                    ra_trong = pd.isna(ra) or str(ra).strip().lower() in ['', 'nan', 'none', 'nat']
                    
                    if vao_trong and not ra_trong:
                        return "⚠️ Thiếu giờ vào"
                    elif ra_trong and not vao_trong:
                        return "⚠️ Thiếu giờ ra"
                    elif row["Số giờ làm thực tế"] == -1:
                        return "⚠️ Lỗi check-out"
                    elif 0 < row["Số giờ làm thực tế"] < 4:
                        return "⚠️ Làm thiếu giờ (< 4h)"
                    else:
                        return ""
                        
                df_filtered["Ghi chú"] = df_filtered.apply(check_anomaly, axis=1)
                
                # Bắt lỗi cờ -1 và chuyển thành 0
                if (df_filtered["Số giờ làm thực tế"] == -1).any():
                    st.warning("⚠️ Phát hiện một số dòng có giờ ra sớm hơn giờ vào (lỗi check-out), đã được tính là 0 giờ.")
                
                df_filtered["Số giờ làm thực tế"] = df_filtered["Số giờ làm thực tế"].replace(-1, 0)
                
                # Format và Làm tròn giờ
                df_filtered["Số giờ làm thực tế"] = df_filtered["Số giờ làm thực tế"].apply(format_gio_lam)
                df_filtered["Giờ hành chính"] = df_filtered["Giờ hành chính"].apply(format_gio_lam)
                df_filtered["Giờ OT"] = df_filtered["Giờ OT"].apply(format_gio_lam)
                
                df_filtered["Ngày"] = df_filtered["_parsed_date"].dt.strftime('%d/%m/%Y')
                
                # Sắp xếp theo ID -> Ngày
                df_filtered = df_filtered.sort_values(by=[m['ma_nv'], "_parsed_date"])

                # Nhãn duy nhất "Mã NV - Tên NV" để lọc (tránh trùng khi 2 người cùng tên)
                df_filtered["_nv_label"] = df_filtered[m['ma_nv']].astype(str) + " - " + df_filtered[m['ten_nv']].astype(str)
                danh_sach_nv = sorted(df_filtered["_nv_label"].unique().tolist())

            st.sidebar.markdown("### 🔍 Bước 4: Lọc dữ liệu")
            chon_nv = st.sidebar.multiselect(
                "Chọn nhân viên",
                options=danh_sach_nv,
                default=[],
                placeholder="Chọn nhân viên muốn xem (bỏ trống để xem tất cả)",
                label_visibility="collapsed",
            )
            
            st.markdown("## 📊 Kết quả Tổng hợp")
            if chon_nv:
                df_filtered = df_filtered[df_filtered["_nv_label"].isin(chon_nv)]

            with st.spinner("⏳ Đang tổng hợp kết quả..."):
                # Thống kê — hiển thị trước bảng
                total_rows = len(df_filtered)
                total_emps = df_filtered[m['ma_nv']].nunique()
                df_numeric = df_filtered[pd.to_numeric(df_filtered['Số giờ làm thực tế'], errors='coerce').notnull()]
                total_hours = df_numeric['Số giờ làm thực tế'].sum() if not df_numeric.empty else 0
                ngay_nghi = int((df_filtered["Số giờ làm thực tế"] == 0).sum())

                st.markdown('<div class="card"><div class="card-title"><span class="card-icon">📈</span>Tổng quan kỳ công</div></div>', unsafe_allow_html=True)
                c1, c2, c3, c4 = st.columns(4)
                c1.metric(
                    "Số ngày làm việc",
                    working_days,
                    help="Tổng số ngày trong kỳ: trừ Chủ nhật và Thứ 7 (trừ Thứ 7 cuối tháng — ngày làm bắt buộc), trừ tiếp các ngày nghỉ lễ trùng ngày làm việc."
                )
                c2.metric("Số nhân viên", total_emps)
                c3.metric("Tổng giờ làm", f"{format_gio_lam(total_hours)} giờ")
                c4.metric("Số ngày nghỉ", ngay_nghi)

                st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

                # Tạo bảng UI (chỉ hiển thị 5 cột)
                gio_lam_numeric = df_filtered["Số giờ làm thực tế"].tolist()
                
                df_result_ui = pd.DataFrame({
                    "STT": range(1, len(df_filtered) + 1),
                    "Mã NV": df_filtered[m['ma_nv']].values,
                    "Tên nhân viên": df_filtered[m['ten_nv']].values,
                    "Ngày": df_filtered["Ngày"].values,
                    "Giờ HC": [format_gio_lam_str(v) for v in df_filtered["Giờ hành chính"]],
                    "Trễ (p)": df_filtered["Phút đi trễ"].values,
                    "Sớm (p)": df_filtered["Phút về sớm"].values,
                    "OT": [format_gio_lam_str(v) for v in df_filtered["Giờ OT"]],
                    "Tổng giờ": [format_gio_lam_str(v) for v in df_filtered["Số giờ làm thực tế"]],
                    "Ghi chú": df_filtered["Ghi chú"].values
                })
                def get_loai_ngay(ngay_str, gio_vao_raw, gio_ra_raw):
                    try:
                        ngay = datetime.datetime.strptime(ngay_str, '%d/%m/%Y').date()
                    except:
                        return 'binh_thuong'
            
                    # 1. Ngày lễ cố định hoặc nhập tay (bao gồm nghỉ bù)
                    if ngay in fixed_holidays or ngay in NGAY_LE_THU_CONG:
                        return 'le'
            
                    # 2. Cuối tuần
                    if ngay.weekday() in [5, 6]:
                        return 'nghi'
            
                    # 3. Ngày thường nhưng thiếu dữ liệu
                    vao_trong = pd.isna(gio_vao_raw) or str(gio_vao_raw).strip().lower() in ['', 'nan', 'none', 'nat']
                    ra_trong  = pd.isna(gio_ra_raw) or str(gio_ra_raw).strip().lower()  in ['', 'nan', 'none', 'nat']
                    if vao_trong or ra_trong:
                        return 'bat_thuong'
            
                    return 'binh_thuong'

                df_result_ui["_loai"] = [
                    get_loai_ngay(r["Ngày"], r[m['gio_vao']], r[m['gio_ra']]) 
                    for _, r in df_filtered.iterrows()
                ]
                
                def style_notes(val):
                    if "⚠️" in str(val):
                        return 'background-color: #FEE2E2; color: #991B1B; font-weight: 500'
                    return ''
                    
                def style_row(row):
                    loai = df_result_ui.loc[row.name, "_loai"]  # lấy từ df gốc có _loai
                    styles = [""] * len(row)
                    idx_gio = list(row.index).index("Tổng giờ")
                    
                    if loai == "le":
                        styles = ["background-color: #FEE2E2"] * len(row)
                        styles[idx_gio] = "background-color: #FEE2E2; color: #DC2626; font-weight: 600"
                    elif loai == "nghi":
                        styles = ["background-color: #F3F4F6"] * len(row)
                        styles[idx_gio] = "background-color: #F3F4F6; color: #9CA3AF"
                    else:
                        styles[idx_gio] = "color: #0F6E56; font-weight: 500"
                    
                    return styles
                    
                st.markdown('<div class="card"><div class="card-title"><span class="card-icon">📋</span>Bước 3 — Kết quả chi tiết</div></div>', unsafe_allow_html=True)
                
                df_display = df_result_ui.drop(columns=["_loai"])
                
                tab_table, tab_chart = st.tabs(["📑 Bảng Dữ liệu", "📈 Biểu đồ Thống kê (Sắp ra mắt)"])
                
                with tab_table:
                    st.dataframe(
                        df_display.style.apply(style_row, axis=1).map(style_notes, subset=["Ghi chú"]), 
                        use_container_width=True, 
                        hide_index=True,
                        height=min(600, 40 + len(df_result_ui) * 35),
                        column_config={
                            "STT": st.column_config.NumberColumn("STT", width="small", format="%d"),
                            "Mã NV": st.column_config.TextColumn("Mã NV", width="small"),
                            "Tên nhân viên": st.column_config.TextColumn("Tên nhân viên", width="medium"),
                            "Ngày": st.column_config.TextColumn("Ngày", width="small"),
                            "Giờ HC": st.column_config.TextColumn("Giờ HC", width="small"),
                            "Trễ (p)": st.column_config.NumberColumn("Trễ (p)", width="small"),
                            "Sớm (p)": st.column_config.NumberColumn("Sớm (p)", width="small"),
                            "OT": st.column_config.TextColumn("OT", width="small"),
                            "Tổng giờ": st.column_config.TextColumn("Tổng giờ", width="small"),
                            "Ghi chú": st.column_config.TextColumn("Ghi chú", width="medium"),
                        }
                    )
                
                with tab_chart:
                    st.info("Tính năng Biểu đồ thống kê đang được phát triển...")

                st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
                
                # ----------------- PHẦN E: XUẤT FILE & LƯU DB -----------------
                st.markdown("## ⬇️ Xuất file & Lưu trữ")
                
                col_exp1, col_exp2, col_exp3 = st.columns([2, 1, 1])
                with col_exp1:
                    st.markdown(f"""
                    <div style='background:#F0FDF8;border:0.5px solid #6EE7C0;border-radius:10px;padding:12px 16px;font-size:13px;color:#0F6E56'>
                        📄 File: <b>chitiet_chamcong_{start_date.strftime('%d%m%Y')}_{end_date.strftime('%d%m%Y')}.xlsx</b><br>
                        <span style='color:#6B7280;font-size:12px'>Gồm {total_rows} dòng · {total_emps} nhân viên</span>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col_exp2:
                    if st.button("💾 Lưu vào Hệ thống", use_container_width=True):
                        with st.spinner("Đang lưu dữ liệu..."):
                            save_to_db(df_filtered, m)
                        st.success("Đã lưu thành công!")
                
                with col_exp3:
                    import io
                    file_bytes = st.session_state.get('uploaded_file_bytes')
                    if file_bytes:
                        file_obj = io.BytesIO(file_bytes)
                        excel_data = export_excel_tong_hop(file_obj, df_filtered, m, so_gio_toi_da)
                        
                        file_name_export = (
                            f"chitiet_chamcong_"
                            f"{start_date.strftime('%d%m%Y')}_"
                            f"{end_date.strftime('%d%m%Y')}.xlsx"
                        )
                    
                    st.download_button(
                        label="⬇️ Tải Excel",
                        data=excel_data,
                        file_name=file_name_export,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True,
                    )
