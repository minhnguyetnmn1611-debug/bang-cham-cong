import streamlit as st
import pandas as pd
import datetime
import sqlite3
from theme import get_theme
from translations import get_t, translate_name
from email_service import send_email_notification
from db import DB_FILE, get_company_emp_dict

def render_kpi_schedule_page():
    OTAKI_B64 = st.session_state.get('OTAKI_B64', '')
    if not OTAKI_B64:
        import os, base64
        base_dir = os.path.dirname(os.path.abspath(__file__))
        img_path = os.path.join(base_dir, "assets", "otaki.png")
        if os.path.exists(img_path):
            try:
                with open(img_path, "rb") as f:
                    OTAKI_B64 = base64.b64encode(f.read()).decode("utf-8")
            except Exception:
                OTAKI_B64 = ""

    t = get_t(st.session_state.get('lang', 'vi'))
    T = get_theme(st.session_state.get('theme_mode', 'light'))
    is_sepia = st.session_state.get('eye_care_sepia', False)

    is_vi = (st.session_state.get('lang', 'vi') == 'vi')
    
    banner_title = t("auto_text_page_kpi_schedule_1")
    banner_desc = t("auto_text_page_kpi_schedule_2")

    st.markdown(f"""
    <div style="background: linear-gradient(135deg, #EC4899 0%, #1E40AF 100%); padding: 24px; border-radius: 18px; color: white; margin-bottom: 24px; box-shadow: 0 10px 30px rgba(236,72,153,0.3);">
        <h2 style="margin: 0; font-size: 28px; font-weight: 800; color: white; display: flex; align-items: center; gap: 12px;">
            <span>&#9733;</span> {banner_title}
        </h2>
        <p style="margin: 8px 0 0 0; opacity: 0.95; font-size: 15px; line-height: 1.5;">
            {banner_desc}
        </p>
    </div>
    """, unsafe_allow_html=True)

    tab1_lbl = t("auto_text_page_kpi_schedule_3")
    tab2_lbl = t("auto_text_page_kpi_schedule_4")
    tab3_lbl = t("auto_text_page_kpi_schedule_5")

    tab1, tab2, tab3 = st.tabs([tab1_lbl, tab2_lbl, tab3_lbl])

    # Load employees
    sys_emps = get_company_emp_dict(st.session_state.get('lang', 'vi'))
    emp_list = []
    dept_map = {
        "CK01": "Thiết kế cơ khí", "VM011": "Thiết kế cơ khí", "VM014": "Thiết kế cơ khí", "VM039": "Thiết kế cơ khí",
        "VM018": "Thiết kế điện điều khiển", "VM020": "Thiết kế điện điều khiển", "VM025": "Thiết kế điện điều khiển", "VM028": "Thiết kế điện điều khiển",
        "VM022": "Mô phỏng 3D", "VM024": "Mô phỏng 3D", "VM013": "Mô phỏng 3D", "VM037": "Mô phỏng 3D", "VM038": "Mô phỏng 3D",
        "HC01": "Bộ phận hành chính - kế toán", "VM012": "Bộ phận hành chính - kế toán"
    }
    dept_ja_map = {
        "Thiết kế cơ khí": "機械設計部",
        "Mô phỏng 3D": "3Dシミュレーション部",
        "Thiết kế điện điều khiển": "制御・電気設計部",
        "Bộ phận hành chính - kế toán": "総務・経理部"
    }

    def resolve_emp_dept(ma, ten):
        ma_u = str(ma).strip().upper()
        if ma_u in dept_map:
            return dept_map[ma_u]

        # 1. Tra cứu trực tiếp trong DB SQLite bảng employees
        try:
            conn = sqlite3.connect(DB_FILE)
            cursor = conn.cursor()
            cursor.execute("SELECT phong_ban FROM employees WHERE UPPER(ma_nv) = ?", (ma_u,))
            row = cursor.fetchone()
            conn.close()
            if row and row[0] and str(row[0]).strip():
                return str(row[0]).strip()
        except Exception:
            pass

        # 2. So sánh chính xác theo từ nguyên chữ (Word token matching)
        import unicodedata
        s_norm = unicodedata.normalize('NFD', str(ten).lower())
        no_accents = ''.join(c for c in s_norm if unicodedata.category(c) != 'Mn').replace('đ', 'd')
        words = set(no_accents.split())

        if 'long' in words or 'dao' in words:
            return "Thiết kế cơ khí"
        elif 'phuong' in words:
            return "Bộ phận hành chính - kế toán"
        elif 'du' in words:
            return "Thiết kế điện điều khiển"
        elif any(w in words for w in ['hung', 'quan', 'hang', 'nguyet']):
            return "Mô phỏng 3D"

        return "Mô phỏng 3D"

    for ma, ten in sys_emps.items():
        if str(ma).strip().upper() in ['GD01', 'VM001'] or 'otaki' in str(ten).lower() or 'masahide' in str(ten).lower() or '大滝' in str(ten):
            continue
        pb_vn = resolve_emp_dept(ma, ten)
        pb_display = pb_vn if is_vi else dept_ja_map.get(pb_vn, pb_vn)
        emp_list.append({"ma": ma, "ten": ten, "pb": pb_display, "pb_vn": pb_vn})

    with tab1:
        c1, c2 = st.columns([1.5, 1])
        with c1:
            dept_lbl = t("auto_text_page_kpi_schedule_6")
            dept_opts = ["Tất cả bộ phận", "Thiết kế cơ khí", "Mô phỏng 3D", "Thiết kế điện điều khiển", "Bộ phận hành chính - kế toán"] if is_vi else ["全部署", "機械設計部", "3Dシミュレーション部", "制御・電気設計部", "総務・経理部"]
            dept_filter = st.selectbox(dept_lbl, dept_opts, key="gantt_dept")
        with c2:
            month_lbl = t("auto_text_page_kpi_schedule_7")
            month_opts = ["Tháng 07/2026", "Tháng 08/2026", "Tháng 09/2026"] if is_vi else ["2026年07月", "2026年08月", "2026年09月"]
            month_filter = st.selectbox(month_lbl, month_opts, key="gantt_month")

        st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
        
        approved_leaves = set()
        approved_leaves_reasons = {}
        for k, v in st.session_state.get('manual_leave', {}).items():
            if isinstance(k, tuple) and len(k) == 2:
                key = (str(k[0]).strip().upper(), str(k[1]).strip())
                approved_leaves.add(key)
                reason_str = v if isinstance(v, str) and v not in ["True", "False"] else (t("auto_text_page_kpi_schedule_8"))
                approved_leaves_reasons[key] = reason_str

        for req in st.session_state.get('pending_hr_approvals', []):
            if req.get('type') in ['Nghỉ phép', '有給休暇'] and ('Duyệt' in req.get('status', '') or '承認' in req.get('status', '')):
                key = (str(req.get('emp')).strip().upper(), str(req.get('date')).strip())
                approved_leaves.add(key)
                approved_leaves_reasons[key] = req.get('reason', t("auto_text_page_kpi_schedule_8"))

        approved_ots = set()
        approved_ots_reasons = {}
        for k, v in st.session_state.get('manual_ot_reason', {}).items():
            if isinstance(k, tuple) and len(k) == 2:
                key = (str(k[0]).strip().upper(), str(k[1]).strip())
                approved_ots.add(key)
                approved_ots_reasons[key] = str(v)

        total_staff = len(emp_list) if emp_list else 8
        tech_staff = total_staff
        
        month_num = "07"
        if "08" in month_filter: month_num = "08"
        elif "09" in month_filter: month_num = "09"

        import calendar
        num_days = calendar.monthrange(2026, int(month_num))[1]

        # Lọc nghỉ phép trong tháng được chọn
        approved_leaves_month = {k for k in approved_leaves if f"/{month_num}/2026" in k[1] or f"/{int(month_num)}/2026" in k[1] or f"2026-{month_num}-" in k[1]}
        al_staff = len(approved_leaves_month) if approved_leaves_month else len(approved_leaves)

        mos_staff = max(0, tech_staff - al_staff)

        m1, m2, m3 = st.columns(3)
        with m1:
            readiness = round((mos_staff / max(1, tech_staff)) * 100.0, 1)
            m1_title = t("auto_text_page_kpi_schedule_10")
            m1_sub = t("auto_text_page_kpi_schedule_11")
            st.markdown(f"""
            <div style="background: {T['bg_card']}; padding: 18px; border-radius: 14px; border: 1.5px solid {T['border']}; box-shadow: 0 4px 10px rgba(0,0,0,0.02);">
                <div style="color: {T['text_secondary']}; font-size: 13.5px; font-weight: 700;">{m1_title}</div>
                <div style="color: #EC4899; font-size: 28px; font-weight: 800; margin-top: 4px;">{readiness}% <span style="font-size: 14px; color: #10B981; font-weight: 700;">{m1_sub}</span></div>
            </div>
            """, unsafe_allow_html=True)
        with m2:
            m2_title = t("auto_text_page_kpi_schedule_12")
            m2_sub = t("auto_text_page_kpi_schedule_13")
            st.markdown(f"""
            <div style="background: {T['bg_card']}; padding: 18px; border-radius: 14px; border: 1.5px solid {T['border']}; box-shadow: 0 4px 10px rgba(0,0,0,0.02);">
                <div style="color: {T['text_secondary']}; font-size: 13.5px; font-weight: 700;">{m2_title}</div>
                <div style="color: #EC4899; font-size: 28px; font-weight: 800; margin-top: 4px;">{mos_staff} <span style="font-size: 14px; color: {T['text_tertiary']};">/ {tech_staff} {m2_sub}</span></div>
            </div>
            """, unsafe_allow_html=True)
        with m3:
            m3_title = t("auto_text_page_kpi_schedule_14")
            m3_sub = t("auto_text_page_kpi_schedule_15")
            st.markdown(f"""
            <div style="background: {T['bg_card']}; padding: 18px; border-radius: 14px; border: 1.5px solid {T['border']}; box-shadow: 0 4px 10px rgba(0,0,0,0.02);">
                <div style="color: {T['text_secondary']}; font-size: 13.5px; font-weight: 700;">{m3_title}</div>
                <div style="color: #F59E0B; font-size: 28px; font-weight: 800; margin-top: 4px;">{al_staff} <span style="font-size: 14px; color: {T['text_tertiary']};">{m3_sub}</span></div>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("<div style='height: 15px;'></div>", unsafe_allow_html=True)
        gantt_heading = t("auto_text_page_kpi_schedule_16")
        st.markdown(f"<h4 style='color: {T['text_primary']}; font-size: 17px; font-weight: 800; margin-bottom: 12px;'>{gantt_heading}</h4>", unsafe_allow_html=True)

        f_emps = [e for e in emp_list if dept_filter in ["Tất cả bộ phận", "全部署"] or e["pb"] == dept_filter or e["pb_vn"] == dept_filter]
        if not f_emps: f_emps = emp_list

        gantt_rows = ""
        day_lbl = t("auto_text_page_kpi_schedule_17")
        for idx, emp in enumerate(f_emps):
            char_avt = emp['ten'][0].upper()
            emp_code = str(emp['ma']).strip().upper()
            emp_ten = emp.get('ten', '')
            is_ceo_row = (emp_code == 'GD01' or 'otaki' in emp_ten.lower() or 'masahide' in emp_ten.lower() or '\u5927\u6edd' in emp_ten)
            if is_ceo_row:
                if OTAKI_B64:
                    row_avt = f'<img src="data:image/png;base64,{OTAKI_B64}" style="width: 38px; height: 38px; border-radius: 50%; object-fit: cover; border: 1.5px solid #F59E0B; flex-shrink: 0;">'
                else:
                    row_avt = f'<div style="width: 38px; height: 38px; border-radius: 50%; background: {T["primary_gradient"]}; color: white; display: flex; align-items: center; justify-content: center; font-weight: 800; font-size: 15px; flex-shrink: 0;">{char_avt}</div>'
            else:
                row_avt = f'<div style="width: 38px; height: 38px; border-radius: 50%; background: linear-gradient(135deg, #E0F2FE, #F0F9FF); color: #DB2777; border: 1.5px solid #BAE6FD; display: flex; align-items: center; justify-content: center; font-weight: 800; font-size: 15px; flex-shrink: 0; box-shadow: 0 2px 5px rgba(236,72,153,0.08);">{char_avt}</div>'

            days_html = ""
            for d in range(1, num_days + 1):
                date_iso = f"2026-{month_num}-{d:02d}"
                date_vn = f"{d:02d}/{month_num}/2026"
                date_vn_short = f"{d}/{int(month_num)}/2026"
                
                matched_leave_reason = None
                for dt in [date_iso, date_vn, date_vn_short]:
                    if (emp_code, dt) in approved_leaves_reasons:
                        matched_leave_reason = approved_leaves_reasons[(emp_code, dt)]
                        break
                        
                matched_ot_reason = None
                for dt in [date_iso, date_vn, date_vn_short]:
                    if (emp_code, dt) in approved_ots_reasons:
                        matched_ot_reason = approved_ots_reasons[(emp_code, dt)]
                        break
                        
                try:
                    is_weekend = (datetime.date(2026, int(month_num), d).weekday() in [5, 6])
                except Exception:
                    is_weekend = False
                
                if matched_leave_reason is not None:
                    status = "al"
                elif matched_ot_reason is not None:
                    status = "ot"
                elif is_weekend:
                    status = "off"
                else:
                    status = "work"
                
                if status == "work":
                    work_desc = f"Làm việc chuyên môn ({emp['pb']})" if ("hành chính" in emp['pb_vn'].lower() or "総務" in emp['pb']) else "Thực hiện dự án (MOS / Kỹ thuật)"
                    if not is_vi:
                        work_desc = f"専門業務 ({emp['pb']})" if ("hành chính" in emp['pb_vn'].lower() or "総務" in emp['pb']) else "プロジェクト実施 (MOS / 技術)"
                    title = work_desc
                    cell_style = f"background: {T['bg_card_hover']}; border: 1px solid {T['border']}; color: {T['text_primary']};"
                    inner_content = ""
                elif status == "al":
                    title = f"Nghỉ phép ({matched_leave_reason})" if is_vi else f"有給休暇 ({matched_leave_reason})"
                    cell_style = f"background: {T['bg_card']}; border: 1.5px solid {T['warning']}; color: {T['warning']}; font-weight: 800; font-size: 11px; box-shadow: 0 2px 6px rgba(245,158,11,0.15);"
                    inner_content = "AL"
                elif status == "ot":
                    title = f"Tăng ca OT ({matched_ot_reason})" if is_vi else f"残業OT ({matched_ot_reason})"
                    cell_style = f"background: {T['bg_card_hover']}; border: 1.5px solid {T['accent']}; color: {T['accent']}; font-weight: 800; font-size: 11px; box-shadow: 0 2px 6px rgba(139,92,246,0.15);"
                    inner_content = "OT"
                else:
                    title = t("auto_text_page_kpi_schedule_20")
                    cell_style = f"background: transparent; border: 1px dashed {T['border']}; opacity: 0.65;"
                    inner_content = ""
                
                day_title_lbl = f"Ngày {d}" if is_vi else f"{d}日"
                days_html += f'<div title="{day_title_lbl}: {title}" style="flex: 1; min-width: 28px; height: 32px; border-radius: 6px; display: flex; align-items: center; justify-content: center; transition: all 0.2s; {cell_style}">{inner_content}</div>'

            row_html = f'<div style="display: flex; align-items: center; border-bottom: 1.5px solid {T["border"]}; background: {T["bg_card"]}; min-width: 1100px; transition: background 0.2s;"><div style="width: 240px; display: flex; align-items: center; gap: 12px; flex-shrink: 0; position: sticky; left: 0; z-index: 2; background: {T["bg_card"]}; border-right: 1.5px solid {T["border"]}; padding: 12px 16px; box-shadow: 4px 0 10px rgba(0,0,0,0.02);">{row_avt}<div style="overflow: hidden;"><div style="font-weight: 800; color: {T["text_primary"]}; font-size: 14.5px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; font-family: \'Plus Jakarta Sans\', sans-serif;">{emp["ten"]}</div><div style="font-size: 12.5px; color: {T["text_secondary"]}; font-weight: 600; margin-top: 2px;">{emp["pb"]}</div></div></div><div style="flex: 1; display: flex; align-items: center; gap: 4px; padding: 10px 14px;">{days_html}</div></div>'
            gantt_rows += row_html

        hdr_col = (t("auto_text_page_kpi_schedule_22"))
        header_days_html = ""
        for d in range(1, num_days + 1):
            try:
                is_wknd = (datetime.date(2026, int(month_num), d).weekday() in [5, 6])
            except Exception:
                is_wknd = False
            txt_col = T["text_tertiary"] if is_wknd else T["text_secondary"]
            bg_col = T["bg_card_hover"] if is_wknd else "transparent"
            header_days_html += f'<div style="flex: 1; min-width: 28px; height: 32px; display: flex; align-items: center; justify-content: center; background: {bg_col}; border-radius: 6px; font-weight: 800; font-size: 12px; color: {txt_col};">{d:02d}</div>'

        container_html = f'<div style="background: {T["bg_card"]}; border: 1.5px solid {T["border"]}; border-radius: 20px; overflow-x: auto; max-width: 100%; box-shadow: 0 4px 16px rgba(0,0,0,0.03); position: relative;"><div style="display: flex; align-items: center; background: {T["bg_card_hover"]}; border-bottom: 1.5px solid {T["border"]}; font-weight: 700; color: {T["text_secondary"]}; font-size: 13.5px; min-width: 1100px;"><div style="width: 240px; flex-shrink: 0; position: sticky; left: 0; z-index: 3; background: {T["bg_card_hover"]}; border-right: 1.5px solid {T["border"]}; padding: 14px 16px; box-shadow: 4px 0 10px rgba(0,0,0,0.02); font-weight: 800; color: {T["text_primary"]};">{hdr_col}</div><div style="flex: 1; display: flex; align-items: center; gap: 4px; padding: 12px 14px;">{header_days_html}</div></div>{gantt_rows}</div>'
        st.markdown(container_html, unsafe_allow_html=True)

        lg1 = t("auto_text_page_kpi_schedule_23")
        lg2 = t("auto_text_page_kpi_schedule_24")
        lg3 = t("auto_text_page_kpi_schedule_25")
        lg4 = t("auto_text_page_kpi_schedule_26")
        legend_html = f'<div style="display: flex; gap: 24px; margin-top: 16px; font-size: 13.5px; color: {T["text_secondary"]}; font-weight: 600; flex-wrap: wrap;"><span style="display: inline-flex; align-items: center; gap: 8px;"><span style="width: 18px; height: 18px; background: {T["bg_card_hover"]}; border: 1px solid {T["border"]}; border-radius: 4px; display: inline-block;"></span> {lg1}</span><span style="display: inline-flex; align-items: center; gap: 8px;"><span style="width: 18px; height: 18px; background: {T["bg_card"]}; border: 1.5px solid {T["warning"]}; border-radius: 4px; display: inline-flex; align-items: center; justify-content: center; font-size: 10px; font-weight: 800; color: {T["warning"]};">AL</span> {lg2}</span><span style="display: inline-flex; align-items: center; gap: 8px;"><span style="width: 18px; height: 18px; background: {T["bg_card_hover"]}; border: 1.5px solid {T["accent"]}; border-radius: 4px; display: inline-flex; align-items: center; justify-content: center; font-size: 10px; font-weight: 800; color: {T["accent"]};">OT</span> {lg3}</span><span style="display: inline-flex; align-items: center; gap: 8px;"><span style="width: 18px; height: 18px; background: transparent; border: 1px dashed {T["border"]}; border-radius: 4px; display: inline-block;"></span> {lg4}</span></div>'
        st.markdown(legend_html, unsafe_allow_html=True)

    with tab2:
        kpi_h4 = t("auto_text_page_kpi_schedule_27")
        st.markdown(f"<h4 style='color: {T['text_primary']}; font-size: 17px; font-weight: 800; margin-bottom: 16px;'>{kpi_h4}</h4>", unsafe_allow_html=True)
        
        w_att = st.session_state.get('kpi_w_att', 30)
        w_mos = st.session_state.get('kpi_w_mos', 40)
        w_okr = st.session_state.get('kpi_w_okr', 30)

        kpi_cards = ""
        lbl_att = t("auto_text_page_kpi_schedule_28")
        lbl_mos = t("auto_text_page_kpi_schedule_29")
        lbl_okr = t("auto_text_page_kpi_schedule_30")
        unit_pt = t("auto_text_page_kpi_schedule_31")
        for idx, emp in enumerate(f_emps):
            emp_code_upper = str(emp['ma']).strip().upper()
            emp_name_clean = str(emp['ten']).strip()
            
            # 1. Điểm Chấm công (att_score): Tính từ bảng records thực tế trong SQLite
            att_score = 100.0
            try:
                conn = sqlite3.connect(DB_FILE)
                df_rec = pd.read_sql_query(
                    "SELECT di_tre, ve_som, ghi_chu FROM records WHERE (UPPER(ma_nv) = ? OR ten_nv = ?) AND (ngay LIKE ? OR ngay LIKE ?)",
                    conn,
                    params=[emp_code_upper, emp_name_clean, f"%/{month_num}/%", f"%-{month_num}-%"]
                )
                conn.close()
                if not df_rec.empty:
                    latenesses = (pd.to_numeric(df_rec['di_tre'], errors='coerce').fillna(0) > 0).sum() + (pd.to_numeric(df_rec['ve_som'], errors='coerce').fillna(0) > 0).sum()
                    anomalies = df_rec['ghi_chu'].astype(str).str.contains('thiếu|lệch|bất thường', case=False, na=False).sum()
                    deduction = (latenesses * 2.0) + (anomalies * 3.0)
                    att_score = max(70.0, min(100.0, 100.0 - deduction))
            except Exception:
                att_score = 100.0

            # 2. Điểm Dự án MOS (mos_score): Tính từ tổng giờ làm MOS thực tế
            mos_score = 95.0
            try:
                df_raw = st.session_state.get('df_mos_raw')
                df_res = st.session_state.get('df_mos_result')
                df_target = df_raw if df_raw is not None and not df_raw.empty else df_res
                if df_target is not None and not df_target.empty:
                    match_rows = df_target[
                        (df_target['ma_nv'].astype(str).str.upper() == emp_code_upper) |
                        (df_target['ten_nv'].astype(str) == emp_name_clean)
                    ]
                    if not match_rows.empty:
                        col_gio = 'tong_gio' if 'tong_gio' in match_rows.columns else ('Giờ làm (h)' if 'Giờ làm (h)' in match_rows.columns else None)
                        if col_gio:
                            actual_hrs = pd.to_numeric(match_rows[col_gio], errors='coerce').sum()
                            std_hrs = st.session_state.get('mos_kpi_std_hours', 160.0)
                            if std_hrs > 0 and actual_hrs > 0:
                                mos_score = max(75.0, min(100.0, round((actual_hrs / std_hrs) * 100.0, 1)))
                            elif actual_hrs > 0:
                                mos_score = 98.0
            except Exception:
                mos_score = 95.0

            # 3. Điểm OKR / Đóng góp (okr_score): Tính từ lượt check-in công tác & đóng góp thực tế
            okr_score = 90.0
            try:
                conn = sqlite3.connect(DB_FILE)
                df_field = pd.read_sql_query(
                    "SELECT COUNT(*) as cnt FROM field_checkins WHERE UPPER(ma_nv) = ? OR ten_nv = ?",
                    conn,
                    params=[emp_code_upper, emp_name_clean]
                )
                conn.close()
                cnt = df_field.iloc[0]['cnt'] if not df_field.empty else 0
                okr_score = min(100.0, 90.0 + (cnt * 2.0))
            except Exception:
                okr_score = 92.0

            # Tùy chỉnh trực tiếp từ session_state nếu quản lý đã lưu điểm ghi đè
            override_dict = st.session_state.get('kpi_custom_scores', {}).get(emp_code_upper, {})
            att_score = round(float(override_dict.get('att', att_score)), 1)
            mos_score = round(float(override_dict.get('mos', mos_score)), 1)
            okr_score = round(float(override_dict.get('okr', okr_score)), 1)
            
            total_kpi = round((att_score * w_att + mos_score * w_mos + okr_score * w_okr) / 100.0, 1)
            badge_color = "#10B981" if total_kpi >= 95 else ("#EC4899" if total_kpi >= 90 else "#F59E0B")
            badge_lbl = ("Xuất sắc (S)" if total_kpi >= 95 else ("Giỏi (A)" if total_kpi >= 90 else "Khá (B)")) if is_vi else ("優秀 (S)" if total_kpi >= 95 else ("優良 (A)" if total_kpi >= 90 else "良好 (B)"))
            char_avt = emp['ten'][0].upper()

            card_html = f'<div style="background: {T["bg_card"]}; border: 1.5px solid {T["border"]}; border-radius: 16px; padding: 18px; margin-bottom: 12px; display: flex; align-items: center; justify-content: space-between; box-shadow: 0 2px 8px rgba(0,0,0,0.02);"><div style="display: flex; align-items: center; gap: 14px; width: 280px;"><div style="width: 46px; height: 46px; border-radius: 50%; background: linear-gradient(135deg, #EC4899, #EC4899); color: white; display: flex; align-items: center; justify-content: center; font-weight: 800; font-size: 18px;">{char_avt}</div><div><div style="font-weight: 800; color: {T["text_primary"]}; font-size: 15.5px;">{emp["ten"]}</div><div style="color: {T["text_secondary"]}; font-size: 13px;">ID: {emp["ma"]} | {emp["pb"]}</div></div></div><div style="display: flex; gap: 30px; text-align: center;"><div><div style="color: {T["text_secondary"]}; font-size: 12px; font-weight: 600;">{lbl_att} ({w_att}%)</div><div style="color: {T["text_primary"]}; font-size: 16px; font-weight: 800; margin-top: 2px;">{att_score} <span style="font-size: 12px; color: #10B981;">{unit_pt}</span></div></div><div><div style="color: {T["text_secondary"]}; font-size: 12px; font-weight: 600;">{lbl_mos} ({w_mos}%)</div><div style="color: {T["text_primary"]}; font-size: 16px; font-weight: 800; margin-top: 2px;">{mos_score} <span style="font-size: 12px; color: #EC4899;">{unit_pt}</span></div></div><div><div style="color: {T["text_secondary"]}; font-size: 12px; font-weight: 600;">{lbl_okr} ({w_okr}%)</div><div style="color: {T["text_primary"]}; font-size: 16px; font-weight: 800; margin-top: 2px;">{okr_score} <span style="font-size: 12px; color: #8B5CF6;">{unit_pt}</span></div></div></div><div style="text-align: right; width: 180px;"><div style="display: inline-block; padding: 6px 14px; background: {badge_color}; color: white; border-radius: 20px; font-weight: 800; font-size: 14px; box-shadow: 0 4px 10px rgba(0,0,0,0.15);">&#9733; {total_kpi} - {badge_lbl}</div></div></div>'
            kpi_cards += card_html

        st.markdown(kpi_cards, unsafe_allow_html=True)
        st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
        btn_export = t("auto_text_page_kpi_schedule_32")
        if st.button(btn_export, type="primary"):
            st.success(t("auto_text_page_kpi_schedule_33"))

    with tab3:
        t3_heading = t("auto_text_page_kpi_schedule_34")
        t3_desc = t("auto_text_page_kpi_schedule_35")
        st.markdown(f"<h4 style='color: #0F172A; font-size: 17px; font-weight: 800; margin-bottom: 12px;'>{t3_heading}</h4>", unsafe_allow_html=True)
        st.markdown(f"<p style='color: #64748B; font-size: 14.5px; margin-bottom: 20px;'>{t3_desc}</p>", unsafe_allow_html=True)
        
        c1, c2, c3 = st.columns(3)
        with c1:
            sl_att = t("auto_text_page_kpi_schedule_36")
            nw_att = st.slider(sl_att, 0, 100, st.session_state.get('kpi_w_att', 30))
        with c2:
            sl_mos = t("auto_text_page_kpi_schedule_37")
            nw_mos = st.slider(sl_mos, 0, 100, st.session_state.get('kpi_w_mos', 40))
        with c3:
            sl_okr = t("auto_text_page_kpi_schedule_38")
            nw_okr = st.slider(sl_okr, 0, 100, st.session_state.get('kpi_w_okr', 30))

        if nw_att + nw_mos + nw_okr != 100:
            warn_msg = f"⚠️ Tổng tỷ trọng hiện tại là **{nw_att + nw_mos + nw_okr}%**. Vui lòng điều chỉnh để tổng đúng bằng **100%**." if is_vi else f"⚠️ 現在の合計ウェイトは **{nw_att + nw_mos + nw_okr}%** です。合計が **100%** になるように調整してください。"
            st.warning(warn_msg)

        st.markdown("<div style='height: 15px;'></div>", unsafe_allow_html=True)
        btn_save = t("auto_text_page_kpi_schedule_40")
        if st.button(btn_save, type="primary"):
            st.session_state.kpi_w_att = nw_att
            st.session_state.kpi_w_mos = nw_mos
            st.session_state.kpi_w_okr = nw_okr
            save_msg = t("auto_text_page_kpi_schedule_41")
            st.success(save_msg)

