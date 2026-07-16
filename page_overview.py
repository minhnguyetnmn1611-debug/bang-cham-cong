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
    total_emp_count = 0
    latest_period_str = t("auto_text_page_overview_1")
    status_sub = t("auto_text_page_overview_2")

    # 1. Kiểm tra dữ liệu trong session state (khi vừa upload file Excel)
    if st.session_state.get('df_raw') is not None:
        df_temp = st.session_state.df_raw
        if 'ma_nv' in df_temp.columns:
            total_emp_count = df_temp['ma_nv'].nunique()
        elif 'Mã NV' in df_temp.columns:
            total_emp_count = df_temp['Mã NV'].nunique()
        elif 'ten_nv' in df_temp.columns:
            total_emp_count = df_temp['ten_nv'].nunique()
        latest_period_str = t("auto_text_page_overview_3")
        status_sub = t("auto_text_page_overview_4")

    # 2. Kiểm tra trong cơ sở dữ liệu SQLite nếu chưa có trong session state
    if total_emp_count == 0:
        try:
            conn = sqlite3.connect(DB_FILE)
            cursor = conn.cursor()
            
            # Count unique employees
            cursor.execute("SELECT COUNT(DISTINCT ma_nv) FROM records")
            emp_count_res = cursor.fetchone()
            if emp_count_res and emp_count_res[0] is not None:
                total_emp_count = emp_count_res[0]
                
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
    bg_banner = "linear-gradient(135deg, rgba(254, 252, 232, 0.92) 0%, rgba(253, 246, 178, 0.88) 100%)" if is_sepia else "linear-gradient(135deg, rgba(255,255,255,0.92) 0%, rgba(240,249,255,0.88) 100%)"
    banner_border = f"1.5px solid {T['border']}"
    title_color = T["primary"]
    subtitle_color = T["text_secondary"]
    bg_card = f"{T['bg_card']}E6"
    notif_border = f"1px solid {T['border']}"
    title_notif_color = T["text_primary"]

    st.markdown(f"""
    <div style="background: {bg_banner}; backdrop-filter: blur(20px); padding: 24px 30px; border-radius: 20px; box-shadow: 0 10px 30px rgba(14,165,233,0.25); margin-bottom: 24px; border: {banner_border}; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 15px;">
        <div>
            <h1 style="color: {title_color}; font-size: 26px; margin: 0; font-family: Plus Jakarta Sans, Inter, sans-serif; font-weight: 800;">V.MOS ENTERPRISE DASHBOARD</h1>
            <p style="color: {subtitle_color}; font-size: 14.5px; margin-top: 4px; margin-bottom: 0;">{t("auto_text_page_overview_12")}</p>
        </div>
        <div style="background: linear-gradient(135deg, #0EA5E9, #0284C7); color: white; padding: 6px 16px; border-radius: 30px; font-size: 13px; font-weight: 700; box-shadow: 0 4px 12px rgba(14,165,233,0.25);">{t("auto_text_page_overview_13")}</div>
    </div>
    """, unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)
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
        st.markdown(f"""<div style="background: {bg_card}; padding: 18px; border-radius: 16px; box-shadow: inset 5px 0 0 0 #F59E0B, 0 8px 25px rgba(245,158,11,0.25);">
            <div style="color: {subtitle_color}; font-size: 13px; font-weight: 600;">{t("auto_text_page_overview_16")}</div>
            <div style="color: {title_notif_color}; font-size: 28px; font-weight: 800; margin: 4px 0;">{pending_display} <span style="font-size:14px; font-weight:600; color:#F59E0B;">{lbl_pending_sub}</span></div>
            <div style="color: #D97706; font-size: 12.5px; font-weight: 600;">{t("auto_text_page_overview_17")}</div>
        </div>""", unsafe_allow_html=True)
    with c4:
        st.markdown(f"""<div style="background: {bg_card}; padding: 18px; border-radius: 16px; box-shadow: inset 5px 0 0 0 #0EA5E9, 0 8px 25px rgba(14,165,233,0.25);">
            <div style="color: {subtitle_color}; font-size: 13px; font-weight: 600;">{t("auto_text_page_overview_18")}</div>
            <div style="color: {title_notif_color}; font-size: 28px; font-weight: 800; margin: 4px 0;">{active_projects} <span style="font-size:14px; font-weight:600; color:#0EA5E9;">{t("auto_text_page_overview_19")}</span></div>
            <div style="color: #0EA5E9; font-size: 12.5px; font-weight: 600;">{lbl_mos_status}</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<div style='height: 16px;'></div>", unsafe_allow_html=True)

    pending_hr_reqs = [p for p in st.session_state.get('pending_hr_approvals', []) if p['status'] == '⏳ Chờ duyệt']
    live_hr_alert = ""
    if pending_hr_reqs:
        cur_lang = t("auto_text_page_overview_20")
        emp_hr = translate_name(str(pending_hr_reqs[0]['emp']), cur_lang)
        if is_vi:
            live_hr_alert = f"""<div style="padding: 12px; background: {T['warning']}20; border-left: 4px solid {T['warning']}; border-radius: 8px; margin-bottom: 10px; font-size: 13.5px;">
<b style="color: {T['warning']};">🔔 [HR Yêu cầu phê duyệt] Có {len(pending_hr_reqs)} đơn đăng ký Phép/OT mới từ nhân viên!</b> — Kỹ sư {emp_hr} vừa gửi đơn {pending_hr_reqs[0]['type']}.
</div>\n"""
        else:
            live_hr_alert = f"""<div style="padding: 12px; background: {T['warning']}20; border-left: 4px solid {T['warning']}; border-radius: 8px; margin-bottom: 10px; font-size: 13.5px;">
<b style="color: {T['warning']};">🔔 [人事承認待ち] 従業員から未承認の休暇・残業申請が {len(pending_hr_reqs)} 件あります！</b> — {emp_hr} さんから申請が届きました。
</div>\n"""
    dynamic_notifs = ""
    try:
        df_checkins = get_field_checkins(limit=3)
        if not df_checkins.empty:
            cur_lang = t("auto_text_page_overview_20")
            emp_dict_synced = get_company_emp_dict(cur_lang)
            for _, row in df_checkins.iterrows():
                ma_c = str(row.get('ma_nv', '')).strip()
                if ma_c in emp_dict_synced:
                    ten_synced = emp_dict_synced[ma_c]
                else:
                    ten_synced = translate_name(str(row.get('ten_nv', '')), cur_lang)
                dia_diem_synced = translate_dia_diem(str(row.get('dia_diem', '')), cur_lang)
                
                if is_vi:
                    text = f"Vừa báo cáo công tác tại {dia_diem_synced} lúc {row['thoi_gian']}."
                    title = f"[📍 Công tác] {ten_synced} ({ma_c})"
                else:
                    thoi_gian_jp = str(row['thoi_gian'])
                    if len(thoi_gian_jp) >= 10 and thoi_gian_jp[2] == '/' and thoi_gian_jp[5] == '/':
                        thoi_gian_jp = f"{thoi_gian_jp[6:10]}/{thoi_gian_jp[3:5]}/{thoi_gian_jp[0:2]}" + thoi_gian_jp[10:]
                    text = f"{thoi_gian_jp} に {dia_diem_synced} で出張報告しました。"
                    title = f"[📍 出張] {ten_synced} ({ma_c})"
                dynamic_notifs += f"""<div style="padding: 12px; background: {T['primary']}15; border-left: 4px solid {T['primary']}; border-radius: 8px; margin-bottom: 10px; font-size: 13.5px;">
<b style="color: {T['primary']};">{title}</b> — {text}
</div>\n"""
        else:
            if is_vi:
                dynamic_notifs = f"""<div style="padding: 12px; background: {T['bg_card_hover']}; border-left: 4px solid {T['text_secondary']}; border-radius: 8px; font-size: 13.5px;">
<b style="color: {T['text_secondary']};">[✅ Hệ thống] V.MOS System đang hoạt động ổn định.</b> — Chưa có hoạt động ngoại nghiệp nào hôm nay.
</div>\n"""
            else:
                dynamic_notifs = f"""<div style="padding: 12px; background: {T['bg_card_hover']}; border-left: 4px solid {T['text_secondary']}; border-radius: 8px; font-size: 13.5px;">
<b style="color: {T['text_secondary']};">[✅ システム] V.MOS System は正常に稼働しています。</b> — 今日の現場活動はまだありません。
</div>\n"""
    except Exception as e:
        pass

    title_text = t("auto_text_page_overview_22")
    st.markdown(f"""
<div style="background: {bg_card}; padding: 22px; border-radius: 18px; border: {notif_border}; box-shadow: 0 8px 25px rgba(14,165,233,0.25); height: 100%;">
<div style="font-size: 17px; font-weight: 800; color: {title_notif_color}; margin-bottom: 14px; display: flex; align-items: center; gap: 8px;">{title_text}</div>
{live_hr_alert}{dynamic_notifs}</div>
    """, unsafe_allow_html=True)

def render_overview_page():
    t = get_t(st.session_state.get('lang', 'vi'))
    T = get_theme(st.session_state.get('theme_mode', 'light'))
    is_sepia = st.session_state.get('eye_care_sepia', False)

    render_enterprise_dashboard()

