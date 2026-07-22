import sqlite3
import pandas as pd
import streamlit as st
from log_config import logger
from translations import translate_name

DB_FILE = 'attendance.db'

@st.cache_resource
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
    c.execute('''
        CREATE TABLE IF NOT EXISTS employees (
            ma_nv TEXT PRIMARY KEY, ten_nv TEXT, chuc_vu TEXT, phong_ban TEXT
        )
    ''')
    try:
        c.execute('ALTER TABLE records ADD COLUMN version INTEGER DEFAULT 1')
    except sqlite3.OperationalError:
        pass # Column already exists
    
    # Social Network & Vinh danh tables
    c.execute('''
        CREATE TABLE IF NOT EXISTS newsfeed (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            author TEXT,
            content TEXT,
            timestamp DATETIME,
            likes INTEGER DEFAULT 0
        )
    ''')
    
    try:
        c.execute('ALTER TABLE newsfeed ADD COLUMN real_author TEXT')
    except sqlite3.OperationalError:
        pass
        
    try:
        c.execute('ALTER TABLE newsfeed ADD COLUMN is_anonymous INTEGER DEFAULT 0')
    except sqlite3.OperationalError:
        pass
        
    c.execute('''
        CREATE TABLE IF NOT EXISTS system_notifications (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date_str TEXT,
            title_vi TEXT,
            content_vi TEXT,
            title_ja TEXT,
            content_ja TEXT,
            type TEXT DEFAULT 'info'
        )
    ''')
    
    # Insert defaults if empty
    c.execute("SELECT COUNT(*) FROM system_notifications")
    if c.fetchone()[0] == 0:
        defaults = [
            ("01/07/2026", "Cập nhật phiên bản V.MOS Enterprise v3.0", "Nâng cấp toàn diện tính năng thống kê MOS: Tự động đếm số lượng nhân sự, tính toán giờ làm việc tiêu chuẩn và tự động đồng bộ kết quả chỉnh sửa xuống bảng KPI.", "V.MOS Enterprise v3.0 アップデート", "MOS統計機能の全面アップグレード：人数の自動集計、標準稼働時間の自動計算、および編集結果のKPI表への自動同期機能を追加しました。", "info"),
            ("01/07/2026", "Cập nhật quy tắc Kỳ công dự án MOS", "Từ nay, chu kỳ công từ ngày 21 tháng trước đến ngày 20 tháng này sẽ được tự động tính vào kỳ công của tháng này (ví dụ: 21/05 - 20/06 là Tháng 6).", "新しい給与計算期間のルール", "今後、前月21日から当月20日までの期間は当月の期間として自動計算されます（例：5月21日～6月20日は6月度となります）。", "info"),
            ("20/06/2026", "Lịch nghỉ Lễ & Kiểm tra công kỳ tháng 6", "Hạn chót duyệt đơn nghỉ phép và xác nhận công tăng ca (OT) là 17:00 ngày cuối tháng.", "休暇および6月度の工数確認について", "休暇申請および残業（OT）の承認期限は、月末日の17:00までとなります。", "secondary")
        ]
        c.executemany("INSERT INTO system_notifications (date_str, title_vi, content_vi, title_ja, content_ja, type) VALUES (?, ?, ?, ?, ?, ?)", defaults)
        
    c.execute('''
        CREATE TABLE IF NOT EXISTS vinh_danh (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sender TEXT,
            recipient TEXT,
            kudos_type TEXT,
            message TEXT,
            timestamp DATETIME
        )
    ''')
    
    c.execute('''
        CREATE TABLE IF NOT EXISTS courses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT,
            description TEXT,
            icon TEXT,
            file_name TEXT,
            file_path TEXT,
            timestamp DATETIME
        )
    ''')
    
    c.execute('''
        CREATE TABLE IF NOT EXISTS comments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            post_id INTEGER,
            author TEXT,
            content TEXT,
            timestamp DATETIME
        )
    ''')
    
    # Recreate exam tables for file assignment workflow
    c.execute("DROP TABLE IF EXISTS questions")
    c.execute("DROP TABLE IF EXISTS exam_results")
    c.execute("DROP TABLE IF EXISTS exams")
    
    c.execute('''
        CREATE TABLE IF NOT EXISTS exams (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT,
            deadline TEXT,
            file_name TEXT,
            file_path TEXT,
            timestamp DATETIME,
            status TEXT DEFAULT 'active'
        )
    ''')

    c.execute('''
        CREATE TABLE IF NOT EXISTS exam_submissions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            exam_id INTEGER,
            author TEXT,
            file_name TEXT,
            file_path TEXT,
            timestamp DATETIME,
            FOREIGN KEY (exam_id) REFERENCES exams (id)
        )
    ''')

    conn.commit()
    conn.close()

# init_db() called here
init_db()

@st.cache_data(ttl=120)
def _get_db_emps():
    db_emps = {}
    try:
        conn = sqlite3.connect(DB_FILE)
        try:
            df_emp = pd.read_sql_query("SELECT DISTINCT ma_nv, ten_nv FROM employees WHERE ma_nv IS NOT NULL AND ten_nv IS NOT NULL", conn)
            for _, r in df_emp.iterrows():
                ma = str(r['ma_nv']).strip()
                ten = str(r['ten_nv']).strip()
                if ma and ten and ma.lower() not in ['nan', 'none', ''] and ten.lower() not in ['nan', 'none', '']:
                    db_emps[ma] = ten
        except Exception as e: logger.warning(f"Lỗi: {e}", exc_info=True); pass
        try:
            df_rec = pd.read_sql_query("SELECT DISTINCT ma_nv, ten_nv FROM records WHERE ma_nv IS NOT NULL AND ten_nv IS NOT NULL", conn)
            for _, r in df_rec.iterrows():
                ma = str(r['ma_nv']).strip()
                ten = str(r['ten_nv']).strip()
                if ma and ten and ma.lower() not in ['nan', 'none', ''] and ten.lower() not in ['nan', 'none', '']:
                    db_emps[ma] = ten
        except Exception as e: logger.warning(f"Lỗi: {e}", exc_info=True); pass
        try:
            df_fc = pd.read_sql_query("SELECT DISTINCT ma_nv, ten_nv FROM field_checkins WHERE ma_nv IS NOT NULL AND ten_nv IS NOT NULL", conn)
            for _, r in df_fc.iterrows():
                ma = str(r['ma_nv']).strip()
                ten = str(r['ten_nv']).strip()
                if ma and ten and ma.lower() not in ['nan', 'none', ''] and ten.lower() not in ['nan', 'none', '']:
                    db_emps[ma] = ten
        except Exception as e: logger.warning(f"Lỗi: {e}", exc_info=True); pass
        conn.close()
    except Exception as e: logger.warning(f"Lỗi: {e}", exc_info=True); pass
    return db_emps

def get_company_emp_options(lang, auto_detect_columns_func=None):
    emps = {}

    if 'df_raw' in st.session_state and st.session_state.df_raw is not None and auto_detect_columns_func is not None:
        try:
            m_a = auto_detect_columns_func(st.session_state.df_raw)
            if 'ma_nv' in m_a and 'ten_nv' in m_a:
                uniq_e = st.session_state.df_raw[[m_a['ma_nv'], m_a['ten_nv']]].drop_duplicates()
                for _, r in uniq_e.iterrows():
                    ma = str(r[m_a['ma_nv']]).strip()
                    ten = str(r[m_a['ten_nv']]).strip()
                    if ma and ten and ma.lower() not in ['nan', 'none', ''] and ten.lower() not in ['nan', 'none', '']:
                        emps[ma] = ten
        except Exception as e: logger.warning(f"Lỗi: {e}", exc_info=True); pass

    emps.update(_get_db_emps())

    if 'manual_emps' in st.session_state:
        for me in st.session_state.manual_emps:
            ma = str(me.get('ma', '')).strip()
            ten = str(me.get('ten', '')).strip()
            if ma and ten: emps[ma] = ten

    def is_boss(ma, ten):
        m = str(ma).strip().upper()
        t = str(ten).strip().lower()
        return m in ['GD01', 'VM001'] or 'otaki' in t or 'masahide' in t or '大滝' in str(ten) or '正秀' in str(ten)

    raw_opts = [f"{ma} - {translate_name(ten, lang)}" for ma, ten in sorted(emps.items()) if not is_boss(ma, ten)]
    opts = [o for o in raw_opts if not any(b in o.upper() for b in ['VM001', 'GD01', 'OTAKI', 'MASAHIDE', '大滝', '正秀'])]
    return sorted(list(set(opts)))

def get_company_emp_dict(lang, auto_detect_columns_func=None):
    emps = {}
    opts = get_company_emp_options(lang, auto_detect_columns_func)
    for o in opts:
        parts = o.split(" - ")
        if len(parts) >= 2:
            emps[parts[0].strip()] = parts[1].strip()
    return emps

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
    df = pd.read_sql_query("SELECT * FROM field_checkins ORDER BY id DESC LIMIT ?", conn, params=(limit,))
    conn.close()
    return df

def save_to_db(df_filtered, mapping):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    conflicts = []
    
    for _, row in df_filtered.iterrows():
        ma_nv = str(row[mapping['ma_nv']])
        ten_nv = str(row[mapping['ten_nv']])
        ngay = str(row["Ngày"])
        gio_vao = str(row[mapping['gio_vao']])
        gio_ra = str(row[mapping['gio_ra']])
        di_tre = int(row.get("Phút đi trễ", 0))
        ve_som = int(row.get("Phút về sớm", 0))
        ot = float(row.get("Giờ OT", 0))
        tong_gio = float(row.get("Số giờ làm thực tế", 0))
        ghi_chu = str(row.get("Ghi chú", ""))
        
        expected_version = row.get("version", None)
        
        # Check current version in DB
        c.execute("SELECT version FROM records WHERE ma_nv = ? AND ngay = ?", (ma_nv, ngay))
        db_row = c.fetchone()
        
        if db_row is not None:
            db_version = db_row[0]
            if pd.notna(expected_version) and db_version != expected_version:
                # Conflict detected! Someone else updated it
                conflicts.append(f"{ma_nv} - {ten_nv} (Ngày {ngay})")
                continue
            
            # Update existing record and increment version
            c.execute('''
                UPDATE records
                SET ten_nv=?, gio_vao=?, gio_ra=?, di_tre=?, ve_som=?, ot=?, tong_gio=?, ghi_chu=?, version=?
                WHERE ma_nv=? AND ngay=?
            ''', (ten_nv, gio_vao, gio_ra, di_tre, ve_som, ot, tong_gio, ghi_chu, db_version + 1, ma_nv, ngay))
        else:
            # New record
            c.execute('''
                INSERT INTO records
                (ma_nv, ten_nv, ngay, gio_vao, gio_ra, di_tre, ve_som, ot, tong_gio, ghi_chu, version)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1)
            ''', (ma_nv, ten_nv, ngay, gio_vao, gio_ra, di_tre, ve_som, ot, tong_gio, ghi_chu))
            
    conn.commit()
    conn.close()
    return conflicts
