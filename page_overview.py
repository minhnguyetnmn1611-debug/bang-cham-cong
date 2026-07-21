import streamlit as st
import pandas as pd
import datetime
import sqlite3
import re
from theme import get_theme
from translations import get_t, translate_name, translate_dia_diem
from email_service import send_email_notification
from db import DB_FILE, get_company_emp_dict, get_field_checkins

def render_enterprise_dashboard():
    t = get_t(st.session_state.get('lang', 'vi'))
    T = get_theme(st.session_state.get('theme_mode', 'light'))
    is_sepia = st.session_state.get('eye_care_sepia', False)

    is_vi = (st.session_state.get('lang', 'vi') == 'vi')
    
    if not st.session_state.get('sidebar_auto_expanded_once', False):
        import streamlit.components.v1 as components
        components.html("""
        <script>
        setTimeout(function() {
            try {
                const doc = window.parent.document;
                const expandBtn = doc.querySelector('[data-testid="collapsedControl"]') || 
                                  doc.querySelector('[data-testid="stSidebarCollapsedControl"]') || 
                                  doc.querySelector('button[title="Expand sidebar"]');
                if (expandBtn) {
                    expandBtn.click();
                }
            } catch(e) {}
        }, 300);
        </script>
        """, height=0)
        st.session_state['sidebar_auto_expanded_once'] = True
    
    # --- TÍNH TOÁN SỐ LIỆU ĐỘNG TỪ DỮ LIỆU THỰC TẾ ---
    emp_dict_display = get_company_emp_dict(st.session_state.get('lang', 'vi'))
    for emp in st.session_state.get('manual_emps', []):
        emp_dict_display[emp['ma']] = emp['ten']
    for d_ma in st.session_state.get('deleted_emps', set()):
        emp_dict_display.pop(d_ma, None)
        
    # Lọc lại một lần nữa để đảm bảo không đếm sếp hoặc dòng trống (phòng trường hợp cache hoặc manual_emps)
    emp_dict_display = {
        ma: ten for ma, ten in emp_dict_display.items()
        if ma.strip() and ten.strip() and not ('GD01' in ma.upper() or 'otaki' in ten.lower() or 'masahide' in ten.lower() or '大滝' in ten)
    }
    
    total_emp_count = len(emp_dict_display)
    
    latest_period_str = t("auto_text_page_overview_1")
    status_sub = t("auto_text_page_overview_2")

    # 1. Kiểm tra dữ liệu kỳ chấm công trong session state (khi vừa upload file Excel)
    if st.session_state.get('df_raw') is not None:
        latest_period_str = t("auto_text_page_overview_3")
        status_sub = t("auto_text_page_overview_4")
    else:
        # 2. Lấy kỳ chấm công gần nhất từ SQLite
        try:
            conn = sqlite3.connect(DB_FILE)
            cursor = conn.cursor()
            
            # Get latest period
            cursor.execute("SELECT ngay FROM records ORDER BY id DESC LIMIT 100")
            dates = cursor.fetchall()
            if dates:
                valid_dates = [d[0] for d in dates if d[0] is not None]
                if valid_dates:
                    import pandas as pd
                    df_dates = pd.DataFrame(valid_dates, columns=['ngay'])
                    df_dates['thang_nam'] = df_dates['ngay'].apply(lambda x: str(x)[3:10] if len(str(x))>=10 else "N/A")
                    periods = sorted([t for t in df_dates['thang_nam'].unique() if t != "N/A"], reverse=True)
                    if periods:
                        jp_date = f"{periods[0].split('/')[1]}年{periods[0].split('/')[0]}月" if '/' in periods[0] else f"{periods[0]}月"
                        latest_period_str = f"Tháng {periods[0]}" if is_vi else jp_date
                        status_sub = t("auto_text_page_overview_5")
            conn.close()
        except Exception:
            pass

    emp_display = f"{total_emp_count}" if total_emp_count > 0 else "0"
    lbl_emp_sub = t("auto_text_page_overview_6")
    lbl_emp_status = (t("auto_text_page_overview_7")) if total_emp_count > 0 else (t("auto_text_page_overview_8"))

    # Thẻ 3: Đơn Nghỉ / OT chờ duyệt từ session_state
    pending_leave = len(st.session_state.get("manual_leave", {})) + len(st.session_state.get("manual_ot_reason", {})) + len([p for p in st.session_state.get('pending_hr_approvals', []) if p['status'] == '⏳ Chờ duyệt'])
    pending_display = f"{pending_leave}"
    lbl_pending_sub = t("auto_text_page_overview_9")

    # Thẻ 4: Dự án MOS (Đồng bộ số liệu từ trang Tổng hợp MOS hoặc Chấm công)
    active_projects = 0
    try:
        # 1. Ưu tiên lấy từ số liệu đã xử lý bên trang Tổng hợp giờ làm MOS (df_mos_edited / df_mos_result / df_mos_raw)
        if 'df_mos_edited' in st.session_state and st.session_state.df_mos_edited is not None and not st.session_state.df_mos_edited.empty:
            active_projects = len(st.session_state.df_mos_edited)
        elif 'df_mos_result' in st.session_state and st.session_state.df_mos_result is not None and not st.session_state.df_mos_result.empty:
            active_projects = len(st.session_state.df_mos_result)
        elif 'df_mos_raw' in st.session_state and st.session_state.df_mos_raw is not None and not st.session_state.df_mos_raw.empty:
            df_raw_tmp = st.session_state.df_mos_raw
            if 'ma_da' in df_raw_tmp.columns:
                def _temp_key(r):
                    code = str(r.get('ma_da', '')).strip().upper()
                    if '00000' in code or re.match(r'^[A-Z]?0+$', code):
                        mgr = str(r.get('ql_nhat') or r.get('khach') or r.get('tanto') or '').strip()
                        return f"{code}___{mgr}"
                    return code
                active_projects = df_raw_tmp.apply(_temp_key, axis=1).nunique()
            else:
                active_projects = len(df_raw_tmp)
        
        # 2. Nếu chưa tải bên MOS, thử tìm trong dữ liệu chấm công df_raw
        if active_projects == 0 and 'df_raw' in st.session_state and st.session_state.df_raw is not None:
            da_col = None
            for c in st.session_state.df_raw.columns:
                if any(x in str(c).lower() for x in ['dự án', 'project', 'phòng ban', 'bộ phận', 'team', 'ma_da']):
                    da_col = c
                    break
            if da_col:
                cnt_da = st.session_state.df_raw[da_col].nunique()
                if cnt_da > 0:
                    active_projects = cnt_da
    except Exception:
        pass

    lbl_mos_status = (t("auto_text_page_overview_10")) if active_projects > 0 else (t("auto_text_page_overview_11"))

    is_sepia = st.session_state.get('eye_care_sepia', False)
    bg_banner = "linear-gradient(135deg, #FDF8ED 0%, #F4ECD8 100%)" if is_sepia else "linear-gradient(135deg, rgba(255,255,255,0.92) 0%, rgba(240,249,255,0.88) 100%)"
    banner_border = f"1.5px solid {T['border']}"
    title_color = T["primary"]
    subtitle_color = T["text_secondary"]
    bg_card = f"{T['bg_card']}E6"
    notif_border = f"1px solid {T['border']}"
    title_notif_color = T["text_primary"]

    st.markdown(f"""
    <div style="background: {bg_banner}; backdrop-filter: blur(20px); padding: 24px 30px; border-radius: 20px; box-shadow: 0 10px 30px rgba(14,165,233,0.25); margin-top: -130px; margin-bottom: 24px; border: {banner_border}; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 15px;">
        <div>
            <h1 style="color: {title_color}; font-size: 26px; margin: 0; font-family: Plus Jakarta Sans, Inter, sans-serif; font-weight: 800;">{ "BẢNG ĐIỀU KHIỂN V.MOS" if st.session_state.lang == 'vi' else "V.MOS DASHBOARD" }</h1>
            <p style="color: {subtitle_color}; font-size: 14.5px; margin-top: 4px; margin-bottom: 0;">{t("auto_text_page_overview_12")}</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(f"""<div style="background: {bg_card}; padding: 18px; border-radius: 16px; box-shadow: inset 5px 0 0 0 #0EA5E9, 0 8px 25px rgba(14,165,233,0.25);">
            <div style="color: {subtitle_color}; font-size: 13px; font-weight: 600;">{t("auto_text_page_overview_14")}</div>
            <div style="color: {title_notif_color}; font-size: 28px; font-weight: 800; margin: 4px 0;">{emp_display} <span style="font-size:14px; font-weight:600; color:#0EA5E9;">{lbl_emp_sub}</span></div>
            <div style="color: #0EA5E9; font-size: 12.5px; font-weight: 600;">{lbl_emp_status}</div>
        </div>""", unsafe_allow_html=True)
    with c2:
        st.markdown(f"""<div style="background: {bg_card}; padding: 18px; border-radius: 16px; box-shadow: inset 5px 0 0 0 #10B981, 0 8px 25px rgba(16,185,129,0.25);">
            <div style="color: {subtitle_color}; font-size: 13px; font-weight: 600;">{t("auto_text_page_overview_15")}</div>
            <div style="color: {title_notif_color}; font-size: 24px; font-weight: 800; margin: 6px 0;">{latest_period_str}</div>
            <div style="color: #10B981; font-size: 12.5px; font-weight: 600;">{status_sub}</div>
        </div>""", unsafe_allow_html=True)
    with c3:
        st.markdown(f"""<div style="background: {bg_card}; padding: 18px; border-radius: 16px; box-shadow: inset 5px 0 0 0 #0EA5E9, 0 8px 25px rgba(14,165,233,0.25);">
            <div style="color: {subtitle_color}; font-size: 13px; font-weight: 600;">{t("auto_text_page_overview_18")}</div>
            <div style="color: {title_notif_color}; font-size: 28px; font-weight: 800; margin: 4px 0;">{active_projects} <span style="font-size:14px; font-weight:600; color:#0EA5E9;">{t("auto_text_page_overview_19")}</span></div>
            <div style="color: #0EA5E9; font-size: 12.5px; font-weight: 600;">{lbl_mos_status}</div>
        </div>""", unsafe_allow_html=True)



def render_overview_page():
    t = get_t(st.session_state.get('lang', 'vi'))
    T = get_theme(st.session_state.get('theme_mode', 'light'))
    is_sepia = st.session_state.get('eye_care_sepia', False)

    render_enterprise_dashboard()

