import streamlit as st
import pandas as pd
import datetime
import io
import math
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

# ==========================================
# 1. CẤU HÌNH VÀ HÀM TIỆN ÍCH
# ==========================================
st.set_page_config(page_title="Bảng Chấm Công", page_icon="📊", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Be+Vietnam+Pro:wght@400;500;600;700;800&family=Space+Grotesk:wght@500;600;700&display=swap');

:root {
    --brand-50:  #ECFDF6;
    --brand-100: #D3F8E8;
    --brand-200: #A8EFD2;
    --brand-400: #34C794;
    --brand-500: #129D74;
    --brand-600: #0C7F5E;
    --brand-700: #096048;
    --brand-900: #07372B;
    --gold-400:  #D9B27C;
    --gold-500:  #C79A5C;
    --ink-900:   #0B1220;
    --ink-700:   #3B4456;
    --ink-500:   #6B7686;
    --line:      rgba(15, 32, 26, 0.09);
}

html, body, .stApp {
    font-family: 'Be Vietnam Pro', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
}

/* Ẩn header và footer mặc định của Streamlit */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}

/* Nền tổng thể — gradient mesh rất nhẹ, không lấn nội dung */
.stApp {
    background:
        radial-gradient(circle at 10% -8%, rgba(52, 199, 148, 0.12) 0%, transparent 40%),
        radial-gradient(circle at 100% 0%, rgba(217, 178, 124, 0.10) 0%, transparent 38%),
        #F5F8F7;
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
    position: relative;
    background: linear-gradient(135deg, var(--brand-900) 0%, var(--brand-700) 55%, var(--brand-600) 100%);
    border-radius: 18px;
    padding: 22px 28px;
    display: flex;
    align-items: center;
    gap: 16px;
    margin-bottom: 22px;
    box-shadow: 0 18px 36px -16px rgba(7, 55, 43, 0.45);
    overflow: hidden;
}
.app-header::before {
    content: "";
    position: absolute;
    left: 0; right: 0; bottom: 0;
    height: 3px;
    background: linear-gradient(90deg, var(--brand-400), var(--gold-400), var(--brand-400));
}
.app-header::after {
    content: "";
    position: absolute;
    inset: 0;
    background: linear-gradient(120deg, transparent 42%, rgba(217,178,124,0.16) 52%, transparent 62%);
    pointer-events: none;
}
.app-header-icon {
    position: relative;
    font-size: 24px;
    width: 50px;
    height: 50px;
    flex-shrink: 0;
    display: flex;
    align-items: center;
    justify-content: center;
    border-radius: 14px;
    background: rgba(255,255,255,0.14);
    border: 1px solid rgba(255,255,255,0.20);
}
.app-header-title {
    font-size: 19px;
    font-weight: 700;
    color: #ffffff;
    margin: 0;
    letter-spacing: -0.01em;
}
.app-header-sub {
    font-size: 12.5px;
    color: rgba(255,255,255,0.72);
    margin: 4px 0 0 0;
}
.app-header-badge {
    position: relative;
    margin-left: auto;
    background: rgba(255,255,255,0.12);
    border: 1px solid rgba(255,255,255,0.22);
    color: #ffffff;
    font-size: 12px;
    font-weight: 500;
    padding: 7px 16px;
    border-radius: 100px;
    white-space: nowrap;
}

/* ===== Card container ===== */
.card {
    background: #ffffff;
    border: 1px solid var(--line);
    border-radius: 16px;
    padding: 18px 24px;
    margin-bottom: 16px;
    box-shadow: 0 1px 2px rgba(16,24,40,0.03), 0 10px 24px -14px rgba(9,96,72,0.18);
}
.card-title {
    font-size: 14px;
    font-weight: 600;
    color: var(--ink-900);
    margin: 0;
    display: flex;
    align-items: center;
    gap: 10px;
    letter-spacing: -0.01em;
}
.card-icon {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 30px;
    height: 30px;
    flex-shrink: 0;
    font-size: 15px;
    border-radius: 9px;
    background: linear-gradient(135deg, var(--brand-50), var(--brand-100));
    border: 1px solid var(--brand-200);
}

/* Metric cards (tự định nghĩa) */
.metrics-row {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 12px;
    margin-bottom: 16px;
}
.metric-card {
    position: relative;
    background: linear-gradient(180deg, #FBFEFD 0%, #F3FAF8 100%);
    border: 1px solid var(--line);
    border-radius: 14px;
    padding: 16px 18px;
    text-align: center;
    overflow: hidden;
}
.metric-card::before {
    content: "";
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 3px;
    background: linear-gradient(90deg, var(--brand-500), var(--gold-400));
}
.metric-label {
    font-size: 12px;
    color: var(--ink-500);
    margin-bottom: 5px;
    font-weight: 500;
}
.metric-value {
    font-family: 'Space Grotesk', 'Be Vietnam Pro', sans-serif;
    font-size: 25px;
    font-weight: 600;
    color: var(--ink-900);
}
.metric-value span {
    font-size: 13px;
    font-weight: 400;
    color: var(--ink-500);
    font-family: 'Be Vietnam Pro', sans-serif;
}

/* Upload zone */
.upload-hint {
    background: linear-gradient(135deg, var(--brand-50), #ffffff 70%);
    border: 1.5px dashed var(--brand-400);
    border-radius: 12px;
    padding: 11px 16px;
    font-size: 13px;
    color: var(--brand-700);
    margin-bottom: 12px;
}

/* Section divider */
.section-divider {
    border: none;
    border-top: 1px solid var(--line);
    margin: 16px 0;
}

/* Nút xuất file */
.stDownloadButton > button {
    background: linear-gradient(135deg, var(--brand-500), var(--brand-700)) !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    padding: 11px 24px !important;
    font-size: 14px !important;
    font-weight: 600 !important;
    width: 100% !important;
    box-shadow: 0 10px 22px -10px rgba(9,96,72,0.55) !important;
    transition: transform .15s ease, box-shadow .15s ease, filter .15s ease !important;
}
.stDownloadButton > button:hover {
    filter: brightness(1.08) !important;
    transform: translateY(-1px) !important;
    box-shadow: 0 14px 26px -10px rgba(9,96,72,0.6) !important;
}
.stDownloadButton > button:active {
    transform: translateY(0) !important;
}

/* Nút primary (Xác nhận Mapping...) */
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, var(--brand-500), var(--brand-700)) !important;
    border: none !important;
    border-radius: 10px !important;
    font-weight: 600 !important;
    box-shadow: 0 8px 20px -10px rgba(9,96,72,0.5) !important;
    transition: transform .15s ease, filter .15s ease !important;
}
.stButton > button[kind="primary"]:hover {
    filter: brightness(1.08) !important;
    transform: translateY(-1px) !important;
}

/* Hiển thị rõ focus để hỗ trợ điều hướng bàn phím */
.stButton > button:focus-visible,
.stDownloadButton > button:focus-visible {
    outline: 2px solid var(--gold-500) !important;
    outline-offset: 2px !important;
}

/* Streamlit file uploader */
[data-testid="stFileUploader"] {
    background: #ffffff;
    border-radius: 12px;
    padding: 4px;
}
[data-testid="stFileUploaderDropzone"] {
    border-radius: 10px !important;
}

/* Dataframe */
[data-testid="stDataFrame"] {
    border-radius: 12px;
    overflow: hidden;
    border: 1px solid var(--line);
    box-shadow: 0 1px 2px rgba(16,24,40,0.03);
}

/* Date input */
[data-testid="stDateInput"] input {
    border-radius: 8px !important;
}

/* Expander (tuỳ chỉnh giờ làm chuẩn) */
[data-testid="stExpander"] {
    border: 1px solid var(--line) !important;
    border-radius: 12px !important;
    background: #FAFCFB !important;
}

/* Metric widget của Streamlit */
[data-testid="stMetric"] {
    position: relative;
    background: linear-gradient(180deg, #FBFEFD 0%, #F2FAF7 100%);
    border: 1px solid var(--line);
    border-radius: 14px;
    padding: 14px 16px 14px 19px;
    overflow: hidden;
    transition: box-shadow .2s ease;
}
[data-testid="stMetric"]::before {
    content: "";
    position: absolute;
    top: 0; left: 0; bottom: 0;
    width: 3px;
    background: linear-gradient(180deg, var(--brand-500), var(--gold-400));
}
[data-testid="stMetric"]:hover {
    box-shadow: 0 8px 18px -10px rgba(9,96,72,0.3);
}
[data-testid="stMetricLabel"] {
    font-size: 12px !important;
    color: var(--ink-500) !important;
    font-weight: 500 !important;
}
[data-testid="stMetricValue"] {
    font-family: 'Space Grotesk', 'Be Vietnam Pro', sans-serif !important;
    font-size: 23px !important;
    font-weight: 600 !important;
    color: var(--ink-900) !important;
}

/* Spinner */
.stSpinner > div {
    border-top-color: var(--brand-500) !important;
}

/* Alert / Warning */
.stAlert {
    border-radius: 12px !important;
    border: 1px solid var(--line) !important;
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
    background-color: var(--brand-100) !important;
    color: var(--brand-700) !important;
}

/* Subheader / Heading thuần markdown */
.stSubheader, h3 {
    font-size: 15px !important;
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
    """Trả về số (int hoặc float 1 chữ số thập phân) để tính toán."""
    if val is None or (isinstance(val, float) and math.isnan(val)) or str(val).lower() in ['nan', 'nat', 'none', '']:
        return 0
    val = round(float(val), 1)
    return int(val) if val == int(val) else val

def format_gio_lam_str(val):
    """Trả về string để hiển thị: 8 thay vì 8.0, 7.5 thay vì 7.500000."""
    num = format_gio_lam(val)
    return str(num)

# ==========================================
# 2. LOGIC TÍNH TOÁN (CALCULATOR)
# ==========================================
def calculate_working_hours(time_in, time_out, start_chuan=8.0, end_chuan=17.0, lunch_start=12.0, lunch_end=13.0, max_hours=8.0):
    """
    Tính số giờ làm việc thực tế:
    - Mặc định: Bắt đầu 08:00, Kết thúc 17:00, Nghỉ trưa 12:00-13:00, tối đa 8 giờ/ngày
    - Có thể tuỳ chỉnh qua các tham số start_chuan / end_chuan / lunch_start / lunch_end / max_hours
    - Trả về 0 nếu thiếu dữ liệu (ngày nghỉ)
    """
    vao_trong = time_in is None or str(time_in).strip().lower() in ['', 'nan', 'none', 'nat']
    ra_trong = time_out is None or str(time_out).strip().lower() in ['', 'nan', 'none', 'nat']
    
    if vao_trong or ra_trong:
        return 0
        
    t_in = to_time_obj(time_in)
    t_out = to_time_obj(time_out)

    if t_in is None or t_out is None:
        return 0

    in_f = time_to_float(t_in)
    out_f = time_to_float(t_out)

    if out_f < in_f:
        return -1  # Cờ để cảnh báo lỗi giờ ra < giờ vào

    # Ép giờ chuẩn
    eff_in = max(in_f, start_chuan)
    eff_out = min(out_f, end_chuan)

    if eff_out <= eff_in:
        return 0

    total_hours = eff_out - eff_in

    # Trừ giờ nghỉ trưa (Overlap với khoảng nghỉ trưa đã chọn)
    overlap_start = max(eff_in, lunch_start)
    overlap_end = min(eff_out, lunch_end)
    lunch_overlap = max(0.0, overlap_end - overlap_start)

    final_hours = total_hours - lunch_overlap
    
    # Ràng buộc max giờ/ngày và min 0 tiếng
    final_hours = min(max(final_hours, 0.0), max_hours)
    return final_hours

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

# ==========================================
# 3. LOGIC PARSER ĐỌC FILE
# ==========================================
def find_header_row(file, file_name):
    if file_name.endswith('.csv'):
        df_raw = pd.read_csv(file, header=None, nrows=30)
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
        if file_name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, header=header_row)
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
def export_to_excel(df, start_date, end_date):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Kết quả công')
        worksheet = writer.sheets['Kết quả công']
        
        # Định dạng Header
        header_font = Font(bold=True, color='FFFFFF', size=11)
        header_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
        center_alignment = Alignment(horizontal='center', vertical='center')
        thin_border = Border(
            left=Side(style='thin'), right=Side(style='thin'), 
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        
        # Format Header row
        for cell in worksheet[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center_alignment
            cell.border = thin_border
            
        # Freeze pane ở dòng 2 (đóng băng header)
        worksheet.freeze_panes = 'A2'
            
        # Format Data rows
        data_font = Font(name='Calibri', size=11)
        gray_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
        white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        
        for idx, col in enumerate(worksheet.columns, 1):
            max_length = 0
            column_letter = worksheet.cell(row=1, column=idx).column_letter
            
            for i, cell in enumerate(col):
                if i > 0: # Dữ liệu
                    cell.font = data_font
                    cell.alignment = center_alignment
                    
                    # Format Số thập phân cho cột "Giờ làm thực tế" (Cột E / 5)
                    if column_letter == 'E':
                        try:
                            num = format_gio_lam(cell.value)
                            cell.value = str(num)   # ghi string "8", "7.5", "0"
                            if num == 0:
                                cell.fill = gray_fill
                            else:
                                cell.fill = white_fill
                        except:
                            cell.fill = white_fill
                    else:
                        cell.fill = white_fill
                        
                cell.border = thin_border
                
                # Tính độ rộng cột lớn nhất
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
                    
            adjusted_width = (max_length + 2)
            worksheet.column_dimensions[column_letter].width = adjusted_width
            
    return output.getvalue()

# ==========================================
# 5. GIAO DIỆN CHÍNH (UI)
# ==========================================
header_placeholder = st.empty()

def render_header(badge_text=None):
    badge_html = f'<div class="app-header-badge">{badge_text}</div>' if badge_text else ''
    header_placeholder.markdown(f"""
    <div class="app-header">
        <div class="app-header-icon">📊</div>
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

# ----------------- PHẦN A: UPLOAD -----------------
st.markdown("""
<div class="card">
<div class="card-title"><span class="card-icon">📁</span>Bước 1 — Upload file chấm công</div>
</div>
""", unsafe_allow_html=True)
st.markdown('<div class="upload-hint">✅ Hỗ trợ file <b>.xlsx</b>, <b>.xls</b>, <b>.csv</b> xuất từ máy chấm công. Hệ thống tự động nhận diện cột Vào / Ra.</div>', unsafe_allow_html=True)
uploaded_file = st.file_uploader("Chọn file Excel/CSV chấm công", type=["xlsx", "xls", "csv"], label_visibility="collapsed")

if uploaded_file is not None:
    if st.session_state.df_raw is None or st.session_state.get('last_uploaded') != uploaded_file.name:
        with st.spinner("Đang đọc file..."):
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
    
    st.markdown('<div class="card"><div class="card-title"><span class="card-icon">🔗</span>Bước 2 — Xác nhận mapping cột</div></div>', unsafe_allow_html=True)
    st.warning("⚠️ Không thể nhận diện đủ 5 cột tự động. Vui lòng chọn thủ công bên dưới.")
    columns = ["-- Chọn cột --"] + df_raw.columns.tolist()
    
    def get_index(col_name):
        return columns.index(col_name) if col_name in columns else 0

    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        map_emp_id = st.selectbox("Mã NV", columns, index=get_index(mapping_auto.get('ma_nv')))
    with col2:
        map_emp_name = st.selectbox("Tên NV", columns, index=get_index(mapping_auto.get('ten_nv')))
    with col3:
        map_date = st.selectbox("Ngày làm việc", columns, index=get_index(mapping_auto.get('ngay')))
    with col4:
        map_in = st.selectbox("Giờ vào", columns, index=get_index(mapping_auto.get('gio_vao')))
    with col5:
        map_out = st.selectbox("Giờ ra", columns, index=get_index(mapping_auto.get('gio_ra')))
        
    if st.button("Xác nhận Mapping", type="primary"):
        if "-- Chọn cột --" in [map_emp_id, map_emp_name, map_date, map_in, map_out]:
            st.error("❌ Vui lòng chọn đầy đủ 5 cột cần thiết để tiếp tục!")
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

        st.markdown('<div class="card"><div class="card-title"><span class="card-icon">📅</span>Bước 2 — Chọn kỳ công</div></div>', unsafe_allow_html=True)
        col_d1, col_d2, _ = st.columns([1, 1, 2])
        with col_d1:
            start_date = st.date_input("Từ ngày", default_start)
        with col_d2:
            end_date = st.date_input("Đến ngày", default_end)
            
        # Cập nhật Badge trên Header linh động theo ngày
        render_header(f"Kỳ công {start_date.strftime('%d/%m')} – {end_date.strftime('%d/%m/%Y')}")

        with st.expander("⚙️ Tuỳ chỉnh giờ làm chuẩn (mặc định 08:00–17:00 · nghỉ trưa 12:00–13:00 · tối đa 8 giờ/ngày)"):
            col_h1, col_h2, col_h3, col_h4, col_h5 = st.columns(5)
            with col_h1:
                gio_vao_chuan = st.time_input("Giờ vào chuẩn", datetime.time(8, 0))
            with col_h2:
                gio_ra_chuan = st.time_input("Giờ ra chuẩn", datetime.time(17, 0))
            with col_h3:
                nghi_trua_bat_dau = st.time_input("Nghỉ trưa từ", datetime.time(12, 0))
            with col_h4:
                nghi_trua_ket_thuc = st.time_input("Nghỉ trưa đến", datetime.time(13, 0))
            with col_h5:
                so_gio_toi_da = st.number_input("Tối đa giờ/ngày", min_value=0.0, max_value=24.0, value=8.0, step=0.5)

        if time_to_float(gio_ra_chuan) <= time_to_float(gio_vao_chuan):
            st.error("❌ Giờ ra chuẩn phải sau giờ vào chuẩn. Vui lòng kiểm tra lại mục Tuỳ chỉnh giờ làm chuẩn.")
            st.stop()
        if time_to_float(nghi_trua_ket_thuc) < time_to_float(nghi_trua_bat_dau):
            st.error("❌ Giờ kết thúc nghỉ trưa phải sau giờ bắt đầu nghỉ trưa. Vui lòng kiểm tra lại mục Tuỳ chỉnh giờ làm chuẩn.")
            st.stop()

        mask = (df_process["_parsed_date"].dt.date >= start_date) & (df_process["_parsed_date"].dt.date <= end_date)
        df_filtered = df_process.loc[mask].copy()
        
        if df_filtered.empty:
            st.warning("Không có dữ liệu trong khoảng thời gian đã chọn.")
        else:
            with st.spinner("⏳ Đang tính toán giờ làm..."):
                # Tính giờ (theo giờ chuẩn đã tuỳ chỉnh ở trên)
                df_filtered["Số giờ làm thực tế"] = df_filtered.apply(
                    lambda row: calculate_working_hours(
                        row[m['gio_vao']], row[m['gio_ra']],
                        start_chuan=time_to_float(gio_vao_chuan),
                        end_chuan=time_to_float(gio_ra_chuan),
                        lunch_start=time_to_float(nghi_trua_bat_dau),
                        lunch_end=time_to_float(nghi_trua_ket_thuc),
                        max_hours=so_gio_toi_da,
                    ), axis=1
                )
                
                # Bắt lỗi cờ -1 và chuyển thành 0
                if (df_filtered["Số giờ làm thực tế"] == -1).any():
                    st.warning("⚠️ Phát hiện một số dòng có giờ ra sớm hơn giờ vào (lỗi check-out), đã được tính là 0 giờ.")
                
                df_filtered["Số giờ làm thực tế"] = df_filtered["Số giờ làm thực tế"].replace(-1, 0)
                
                # Format và Làm tròn giờ
                df_filtered["Số giờ làm thực tế"] = df_filtered["Số giờ làm thực tế"].apply(format_gio_lam)
                
                df_filtered["Ngày"] = df_filtered["_parsed_date"].dt.strftime('%d/%m/%Y')
                
                # Sắp xếp theo ID -> Ngày
                df_filtered = df_filtered.sort_values(by=[m['ma_nv'], "_parsed_date"])

                # Nhãn duy nhất "Mã NV - Tên NV" để lọc (tránh trùng khi 2 người cùng tên)
                df_filtered["_nv_label"] = df_filtered[m['ma_nv']].astype(str) + " - " + df_filtered[m['ten_nv']].astype(str)
                danh_sach_nv = sorted(df_filtered["_nv_label"].unique().tolist())

            st.markdown('<div class="card"><div class="card-title"><span class="card-icon">🔍</span>Lọc theo nhân viên</div></div>', unsafe_allow_html=True)
            chon_nv = st.multiselect(
                "Chọn nhân viên muốn xem (để trống = xem tất cả)",
                options=danh_sach_nv,
                default=[],
                placeholder="Tìm theo mã hoặc tên nhân viên...",
                label_visibility="collapsed",
            )
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
                c1.metric("Tổng số dòng", total_rows)
                c2.metric("Số nhân viên", total_emps)
                c3.metric("Tổng giờ làm", f"{format_gio_lam(total_hours)} giờ")
                c4.metric("Ngày nghỉ / thiếu", ngay_nghi)

                st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

                # Tạo bảng UI (chỉ hiển thị 5 cột)
                gio_lam_numeric = df_filtered["Số giờ làm thực tế"].tolist()
                
                df_result_ui = pd.DataFrame({
                    "STT": range(1, len(df_filtered) + 1),
                    "Mã nhân viên": df_filtered[m['ma_nv']].values,
                    "Tên nhân viên": df_filtered[m['ten_nv']].values,
                    "Ngày": df_filtered["Ngày"].values,
                    "Giờ làm thực tế": [format_gio_lam_str(v) for v in df_filtered["Số giờ làm thực tế"]]
                })
                
                def style_hours(val):
                    if str(val).strip() == "0":
                        return 'color: #B87333; background-color: #FEF3E2'
                    return 'color: #0F6E56; font-weight: 500'
                    
                st.markdown('<div class="card"><div class="card-title"><span class="card-icon">📋</span>Bước 3 — Kết quả chi tiết</div></div>', unsafe_allow_html=True)
                st.dataframe(
                    df_result_ui.style.map(style_hours, subset=["Giờ làm thực tế"]), 
                    use_container_width=True, 
                    hide_index=True,
                    height=min(600, 40 + len(df_result_ui) * 35),
                    column_config={
                        "STT": st.column_config.NumberColumn("STT", width="small", format="%d"),
                        "Mã nhân viên": st.column_config.TextColumn(width="small"),
                        "Tên nhân viên": st.column_config.TextColumn(width="medium"),
                        "Ngày": st.column_config.TextColumn(width="small"),
                        "Giờ làm thực tế": st.column_config.TextColumn("Giờ làm thực tế", width="small"),
                    }
                )

                st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
                
                # ----------------- PHẦN E: XUẤT FILE -----------------
                st.markdown('<div class="card"><div class="card-title"><span class="card-icon">⬇️</span>Bước 4 — Xuất file Excel</div></div>', unsafe_allow_html=True)
                
                col_exp1, col_exp2 = st.columns([3, 1])
                with col_exp1:
                    st.markdown(f"""
                    <div style='background:#F0FDF8;border:0.5px solid #6EE7C0;border-radius:10px;padding:12px 16px;font-size:13px;color:#0F6E56'>
                        📄 File sẽ được lưu với tên: 
                        <b>ket_qua_cong_{start_date.strftime('%d%m%Y')}_{end_date.strftime('%d%m%Y')}.xlsx</b><br>
                        <span style='color:#6B7280;font-size:12px'>Gồm {total_rows} dòng · {total_emps} nhân viên · Kỳ {start_date.strftime('%d/%m/%Y')} – {end_date.strftime('%d/%m/%Y')}</span>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col_exp2:
                    df_export = df_result_ui.copy()
                    df_export["Giờ làm thực tế"] = gio_lam_numeric
                    excel_data = export_to_excel(df_export, start_date, end_date)
                    file_name_export = f"ket_qua_cong_{start_date.strftime('%d%m%Y')}_{end_date.strftime('%d%m%Y')}.xlsx"
                    
                    st.download_button(
                        label="⬇️ Tải file Excel",
                        data=excel_data,
                        file_name=file_name_export,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True,
                    )
