import streamlit as st
import pandas as pd
import datetime
import sqlite3
from theme import get_theme
from translations import get_t, translate_name
from email_service import send_email_notification
from db import DB_FILE, get_company_emp_dict
from utils import time_to_float

@st.cache_data(ttl=60, show_spinner=False)
def _get_all_history_records():
    try:
        conn = sqlite3.connect(DB_FILE)
        df = pd.read_sql_query("SELECT * FROM records", conn)
        conn.close()
        if not df.empty and 'ngay' in df.columns:
            df['thang_nam'] = df['ngay'].apply(lambda x: str(x)[3:10] if len(str(x))>=10 else "N/A")
        return df
    except Exception as e:
        return pd.DataFrame()

def render_history_page():
    t = get_t(st.session_state.get('lang', 'vi'))
    T = get_theme(st.session_state.get('theme_mode', 'light'))
    is_vi = (st.session_state.get('lang', 'vi') == 'vi')

    tab_cfg, tab1, tab2 = st.tabs([t("auto_text_page_history_2"), t("auto_text_page_history_3"), t("auto_text_page_history_4")])
    
    with tab_cfg:
        render_integrated_settings_content()

    with tab1:
        df_all = _get_all_history_records()
        if df_all.empty:
            st.warning(t("auto_text_page_history_5"))
        else:
            list_thang = sorted([t_val for t_val in df_all['thang_nam'].unique().tolist() if t_val != "N/A"], reverse=True)
            if not list_thang:
                list_thang = ["Tất cả"]
            sel_thang = st.selectbox(t("auto_text_page_history_7"), list_thang, key="sb_hist_tab1_month")
            df_thang = df_all[df_all['thang_nam'] == sel_thang]
            
            c1, c2, c3, c4 = st.columns(4)
            c1.metric(t("auto_text_page_history_8"), f"{df_thang['ma_nv'].nunique()} NV" if is_vi else f"{df_thang['ma_nv'].nunique()} 名")
            c2.metric(t("auto_text_page_history_10"), f"{df_thang['tong_gio'].sum():.1f} h")
            c3.metric(t("auto_text_page_history_11"), f"{df_thang['ot'].sum():.1f} h")
            c4.metric(t("auto_text_page_history_12"), f"{(df_thang['di_tre'] > 0).sum()} lần" if is_vi else f"{(df_thang['di_tre'] > 0).sum()} 回")
            
            df_display = df_thang.copy()
            if 'ten_nv' in df_display.columns:
                df_display['ten_nv'] = [translate_name(x, st.session_state.get('lang', 'vi')) for x in df_display['ten_nv']]
            if not is_vi and 'ngay' in df_display.columns:
                def fmt_jp_d(val):
                    s = str(val).strip()
                    if len(s) == 10 and s[2] == '/' and s[5] == '/':
                        return f"{s[6:10]}/{s[3:5]}/{s[0:2]}"
                    return s
                df_display['ngay'] = df_display['ngay'].apply(fmt_jp_d)
            if 'thang_nam' in df_display.columns:
                df_display = df_display.drop(columns=['thang_nam'])
            for col_c in df_display.columns:
                df_display[col_c] = ["" if str(v).strip().lower() in ['nan', '<na>', 'none', 'null'] else v for v in df_display[col_c]]
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
        df_all = _get_all_history_records()
        if df_all.empty:
            st.warning(t("auto_text_page_history_6"))
        else:
            list_thang = sorted([t_val for t_val in df_all['thang_nam'].unique().tolist() if t_val != "N/A"], reverse=True)
            if len(list_thang) < 2:
                st.info(t("auto_text_page_history_14"))
            else:
                col_k1, col_k2 = st.columns(2)
                k1 = col_k1.selectbox(t("auto_text_page_history_15"), list_thang, index=1 if len(list_thang)>1 else 0, key="sb_hist_tab2_k1")
                k2 = col_k2.selectbox(t("auto_text_page_history_16"), list_thang, index=0, key="sb_hist_tab2_k2")
                
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

def render_integrated_settings_content():
    t = get_t(st.session_state.get('lang', 'vi'))
    T = get_theme(st.session_state.get('theme_mode', 'light'))
    is_sepia = st.session_state.get('eye_care_sepia', False)

    is_vi = (st.session_state.get('lang', 'vi') == 'vi')

    if is_sepia:
        bg_sum = "#FAF5EB"
        bd_sum = "#E6DCCB"
        c_sum = "#5C4434"
        bg_open = "#EFE6D5"
        bd_open = "#C8B9A6"
        c_open = "#4A3525"
    else:
        bg_sum = "#F8FAFC"
        bd_sum = "#E2E8F0"
        c_sum = "#334155"
        bg_open = "#EFF6FF"
        bd_open = "#93C5FD"
        c_open = "#1E40AF"

    st.markdown(f"""
    <style>
    div[data-testid="stExpander"] {{
        max-width: 430px !important;
        border-radius: 14px !important;
        margin-bottom: 26px !important;
        box-shadow: 0 2px 8px rgba(0,0,0,0.03) !important;
    }}
    div[data-testid="stExpander"] summary {{
        justify-content: center !important;
        background: {bg_sum} !important;
        border: 1.5px solid {bd_sum} !important;
        border-radius: 14px !important;
    }}
    div[data-testid="stExpander"] details[open] > summary {{
        background: {bg_open} !important;
        border: 1.5px solid {bd_open} !important;
        border-radius: 14px 14px 0 0 !important;
    }}
    div[data-testid="stExpander"] summary > div {{
        justify-content: center !important;
        width: 100% !important;
    }}
    div[data-testid="stExpander"] summary p {{
        font-size: 14.5px !important;
        font-weight: 700 !important;
        text-align: center !important;
        width: 100% !important;
        color: {c_sum} !important;
    }}
    div[data-testid="stExpander"] details[open] > summary p {{
        color: {c_open} !important;
    }}
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)

    col_l, col_r = st.columns(2)
    with col_l:
        # 0. Quản lý thông báo hệ thống
        with st.expander("🔔 Quản lý thông báo hệ thống" if is_vi else "🔔 お知らせ管理", expanded=False):
            try:
                conn = sqlite3.connect(DB_FILE)
                c = conn.cursor()
                c.execute("SELECT id, date_str, title_vi, type, title_ja FROM system_notifications ORDER BY id DESC")
                notifs_db = c.fetchall()
                
                if notifs_db:
                    st.markdown("**Danh sách thông báo hiện tại:**" if is_vi else "**現在のお知らせ：**")
                    for n in notifs_db:
                        c_n1, c_n2 = st.columns([4, 1])
                        title_display = n[2] if is_vi else (n[4] if len(n) > 4 and n[4] else n[2])
                        
                        date_val = n[1]
                        try:
                            if "/" in date_val:
                                parts = date_val.split("/")
                                if len(parts) == 3:
                                    if len(parts[0]) == 4: # YYYY/MM/DD
                                        d_obj = datetime.datetime(int(parts[0]), int(parts[1]), int(parts[2]))
                                    else: # DD/MM/YYYY
                                        d_obj = datetime.datetime(int(parts[2]), int(parts[1]), int(parts[0]))
                                    date_val = d_obj.strftime("%d/%m/%Y") if is_vi else d_obj.strftime("%Y/%m/%d")
                        except Exception:
                            pass
                            
                        c_n1.markdown(f"`[{date_val}]` {title_display} (*{n[3]}*)")
                        st.markdown(f"""<style>
                        div.st-key-del_notif_{n[0]} button {{
                            padding: 0px 10px !important;
                            height: 32px !important;
                            min-height: 32px !important;
                            line-height: 32px !important;
                            font-size: 13px !important;
                            width: 100% !important;
                        }}
                        </style>""", unsafe_allow_html=True)
                        if c_n2.button("🗑️ Xoá" if is_vi else "🗑️ 削除", key=f"del_notif_{n[0]}"):
                            conn.execute("DELETE FROM system_notifications WHERE id=?", (n[0],))
                            conn.commit()
                            st.rerun()
                
                st.markdown("---")
                st.markdown("**Thêm thông báo mới**" if is_vi else "**新規お知らせ追加**")
                n_date = st.text_input("Ngày (VD: 15/07/2026)" if is_vi else "日付 (例: 2026/07/15)", value=datetime.datetime.now().strftime("%d/%m/%Y") if is_vi else datetime.datetime.now().strftime("%Y/%m/%d"))
                n_type = st.selectbox("Mức độ (Màu sắc)" if is_vi else "重要度 (色)", ["info (Xanh dương)", "secondary (Xám)", "warning (Vàng)"] if is_vi else ["info (ブルー)", "secondary (グレー)", "warning (イエロー)"])
                type_val = n_type.split(" ")[0]
                n_title_vi = st.text_input("Tiêu đề (Tiếng Việt)" if is_vi else "タイトル（ベトナム語）")
                n_content_vi = st.text_area("Nội dung (Tiếng Việt)" if is_vi else "内容（ベトナム語）")
                n_title_ja = st.text_input("Tiêu đề (Tiếng Nhật)" if is_vi else "タイトル（日本語）")
                n_content_ja = st.text_area("Nội dung (Tiếng Nhật)" if is_vi else "内容（日本語）")
                
                if st.button("Thêm thông báo" if is_vi else "追加", use_container_width=True, type="primary"):
                    if n_title_vi and n_content_vi:
                        conn.execute("INSERT INTO system_notifications (date_str, title_vi, content_vi, title_ja, content_ja, type) VALUES (?, ?, ?, ?, ?, ?)",
                                     (n_date, n_title_vi, n_content_vi, n_title_ja, n_content_ja, type_val))
                        conn.commit()
                        st.rerun()
                    else:
                        st.error("Vui lòng nhập đầy đủ Tiêu đề và Nội dung Tiếng Việt!" if is_vi else "ベトナム語のタイトルと内容を入力してください！")
                        
                conn.close()
            except Exception as e:
                st.error(f"Lỗi tải thông báo: {e}")

        # 1. Ngày lễ & Nghỉ bù
        with st.expander(t("exp_holidays") if is_vi else "📅 祝日・代休設定", expanded=False):
            st.markdown(f"<small style='color:#059669; font-weight:700;'>{t('holidays_note')}</small>", unsafe_allow_html=True)
            c_date, c_name = st.columns([1, 2])
            selected_date = c_date.date_input(t("holiday_select"), value=datetime.date.today(), key="set_date_in_holiday")
            holiday_name = c_name.text_input("Tên ngày nghỉ lễ (không bắt buộc)" if is_vi else "祝日名（任意）", placeholder="VD: Nghỉ mát công ty...", key="set_name_holiday")
            
            if st.button(t("btn_add_custom_holiday"), key="set_btn_add_custom_holiday", use_container_width=True):
                # Convert set to dict if needed for backward compatibility
                if isinstance(st.session_state.custom_holidays, set):
                    st.session_state.custom_holidays = {d: "" for d in st.session_state.custom_holidays}
                st.session_state.custom_holidays[selected_date] = holiday_name.strip()
                st.rerun()
                
            if st.session_state.custom_holidays:
                if isinstance(st.session_state.custom_holidays, set):
                    st.session_state.custom_holidays = {d: "" for d in st.session_state.custom_holidays}
                st.markdown(t("custom_holiday_count", count=len(st.session_state.custom_holidays)), unsafe_allow_html=True)
                
                # Hiển thị trực tiếp các ngày đã thêm
                tags_html = ""
                for d in sorted(st.session_state.custom_holidays.keys()):
                    name = st.session_state.custom_holidays[d]
                    display_text = f"<b>{d.strftime('%d/%m/%Y')}</b>: {name}" if name else f"<b>{d.strftime('%d/%m/%Y')}</b>"
                    tags_html += f"<span style='background:#E2E8F0; color:#334155; padding:6px 12px; border-radius:12px; font-size:13.5px; margin: 0 6px 8px 0; display:inline-block; border: 1px solid #CBD5E1;'>{display_text}</span>"
                st.markdown(f"<div style='margin-top:5px; margin-bottom:15px;'>{tags_html}</div>", unsafe_allow_html=True)

                holiday_list = [d.strftime('%d/%m/%Y') for d in sorted(st.session_state.custom_holidays.keys())]
                col_btn1, col_btn2 = st.columns(2)
                with col_btn1:
                    selected_to_remove = st.multiselect(t("holiday_del_sel"), options=holiday_list, placeholder="...", key="set_multi_sel_del_holiday")
                    if st.button(t("btn_del_selected"), key="set_btn_del_sel_holiday", use_container_width=True) and selected_to_remove:
                        for d_str in selected_to_remove:
                            try: del st.session_state.custom_holidays[datetime.datetime.strptime(d_str, '%d/%m/%Y').date()]
                            except Exception as e: logger.warning(f"Lỗi: {e}", exc_info=True); pass
                        st.rerun()
                with col_btn2:
                    if st.button(t("btn_del_all"), key="set_btn_del_all_holiday", use_container_width=True):
                        st.session_state.custom_holidays = {}
                        st.rerun()

        # 2. Tùy chỉnh ngày làm bù
        with st.expander(t("exp_makeup") if is_vi else "🔄 振替出勤日", expanded=False):
            st.markdown(f"<small style='color:#0F172A; font-weight:600;'>{t('makeup_note')}</small>", unsafe_allow_html=True)
            c_mdate, c_mname = st.columns([1, 2])
            selected_makeup = c_mdate.date_input(t("makeup_choose"), value=datetime.date.today(), key="set_date_makeup_input")
            makeup_name = c_mname.text_input("Lý do làm bù (không bắt buộc)" if is_vi else "振替理由（任意）", placeholder="VD: Làm bù cho ngày lễ...", key="set_name_makeup")
            
            if st.button(t("btn_add_makeup"), key="set_btn_add_makeup", use_container_width=True):
                # Convert set to dict if needed for backward compatibility
                if isinstance(st.session_state.custom_workdays, set):
                    st.session_state.custom_workdays = {d: "" for d in st.session_state.custom_workdays}
                st.session_state.custom_workdays[selected_makeup] = makeup_name.strip()
                st.rerun()
                
            if st.session_state.custom_workdays:
                if isinstance(st.session_state.custom_workdays, set):
                    st.session_state.custom_workdays = {d: "" for d in st.session_state.custom_workdays}
                st.markdown(f"<small><b>{t('makeup_count').format(count=len(st.session_state.custom_workdays))}</b></small>", unsafe_allow_html=True)
                
                # Hiển thị trực tiếp các ngày đã thêm
                tags_html_mk = ""
                for d in sorted(st.session_state.custom_workdays.keys()):
                    name = st.session_state.custom_workdays[d]
                    display_text = f"<b>{d.strftime('%d/%m/%Y')}</b>: {name}" if name else f"<b>{d.strftime('%d/%m/%Y')}</b>"
                    tags_html_mk += f"<span style='background:#E0F2FE; color:#0369A1; padding:6px 12px; border-radius:12px; font-size:13.5px; margin: 0 6px 8px 0; display:inline-block; border: 1px solid #BAE6FD;'>{display_text}</span>"
                st.markdown(f"<div style='margin-top:5px; margin-bottom:15px;'>{tags_html_mk}</div>", unsafe_allow_html=True)

                makeup_list = [d.strftime('%d/%m/%Y') for d in sorted(st.session_state.custom_workdays.keys())]
                selected_makeup_remove = st.multiselect(t("makeup_remove_prompt"), options=makeup_list, placeholder="...", key="set_multi_sel_del_makeup")
                c_m1, c_m2 = st.columns(2)
                with c_m1:
                    if st.button(t("btn_del_selected"), key="set_btn_del_sel_makeup", use_container_width=True) and selected_makeup_remove:
                        for d_str in selected_makeup_remove:
                            try: del st.session_state.custom_workdays[datetime.datetime.strptime(d_str, '%d/%m/%Y').date()]
                            except Exception as e: logger.warning(f"Lỗi: {e}", exc_info=True); pass
                        st.rerun()
                with c_m2:
                    if st.button(t("btn_del_all"), key="set_btn_del_all_makeup", use_container_width=True):
                        st.session_state.custom_workdays = {}
                        st.rerun()

        # 3. Tuỳ chỉnh giờ làm chuẩn
        with st.expander(t('sidebar_standard_hours') if is_vi else "⏰ 標準勤務時間設定", expanded=False):
            st.time_input(t("time_in"), datetime.time(8, 0), key="gio_vao_chuan")
            st.time_input(t("time_out"), datetime.time(17, 0), key="gio_ra_chuan")
            st.time_input(t("break_start"), datetime.time(12, 0), key="nghi_trua_bat_dau")
            st.time_input(t("break_end"), datetime.time(13, 0), key="nghi_trua_ket_thuc")
            st.number_input(t("max_hours"), min_value=0.0, max_value=24.0, value=8.0, step=0.5, key="so_gio_toi_da")
            if time_to_float(st.session_state.gio_ra_chuan) <= time_to_float(st.session_state.gio_vao_chuan):
                st.error(t("error_time_out"))
            if time_to_float(st.session_state.nghi_trua_ket_thuc) < time_to_float(st.session_state.nghi_trua_bat_dau):
                st.error(t("error_break"))

    with col_r:
        # 5. Quản lý nhân viên & Thêm mới
        with st.expander(t("auto_text_page_history_21"), expanded=False):
            st.markdown("#### ➕ " + (t("auto_text_page_history_22")))
            # Ngăn tiêu đề rớt dòng để form không bị xô lệch
            st.markdown("<style>div[data-testid='stForm'] label p { white-space: nowrap !important; font-size: 13px !important; }</style>", unsafe_allow_html=True)
            with st.form("form_add_emp_inline", clear_on_submit=True, border=False):
                c1, c2, c3, c4, c5 = st.columns([2, 3, 3, 3, 1])
                new_ma = c1.text_input(t("auto_text_page_history_23"), placeholder="VM001")
                new_ten = c2.text_input(t("auto_text_page_history_24"), placeholder="Nguyễn Văn A")
                new_cv = c3.text_input(t("auto_text_page_history_25"), placeholder="Nhân viên")
                new_pb = c4.text_input(t("auto_text_page_history_26"), placeholder="Kinh doanh")
                
                c5.markdown("<div style='height: 32px;'></div>", unsafe_allow_html=True)
                if c5.form_submit_button("➕", type="tertiary", use_container_width=True):
                    if new_ma and new_ten:
                        m_clean = new_ma.strip().upper()
                        st.session_state.manual_emps.append({"ma": m_clean, "ten": new_ten.strip(), "cv": new_cv.strip(), "pb": new_pb.strip()})
                        try:
                            conn = sqlite3.connect(DB_FILE)
                            conn.execute("INSERT OR REPLACE INTO employees VALUES (?, ?, ?, ?)", (m_clean, new_ten.strip(), new_cv.strip(), new_pb.strip()))
                            conn.commit(); conn.close()
                        except Exception as e: logger.warning(f"Lỗi: {e}", exc_info=True); pass
                        st.rerun()
                    else: st.error(t("auto_text_page_history_27"))

            st.markdown("#### 📋 " + (t("auto_text_page_history_28")))
            
            emp_dict_display = get_company_emp_dict(st.session_state.lang)
            for emp in st.session_state.manual_emps: emp_dict_display[emp['ma']] = emp['ten']
            for d_ma in list(st.session_state.deleted_emps): emp_dict_display.pop(d_ma, None)
            
            db_pb_cv = {}
            try:
                conn = sqlite3.connect(DB_FILE)
                df_emp = pd.read_sql_query("SELECT ma_nv, chuc_vu, phong_ban FROM employees", conn)
                for _, r in df_emp.iterrows():
                    db_pb_cv[str(r['ma_nv']).strip()] = {
                        'cv': str(r['chuc_vu']).strip() if pd.notnull(r['chuc_vu']) and str(r['chuc_vu']).lower() not in ['nan', 'none'] else "",
                        'pb': str(r['phong_ban']).strip() if pd.notnull(r['phong_ban']) and str(r['phong_ban']).lower() not in ['nan', 'none'] else ""
                    }
                conn.close()
            except Exception as e: logger.warning(f"Lỗi: {e}", exc_info=True); pass

            raw_pb_cv = {}
            if 'df_raw' in st.session_state and st.session_state.df_raw is not None:
                df = st.session_state.df_raw
                from utils import auto_detect_columns
                auto_m = auto_detect_columns(df)
                m = st.session_state.get('mapping', {})
                ma_col = m.get('ma_nv', auto_m.get('ma_nv'))
                if ma_col and ma_col in df.columns:
                    cv_col = m.get('chuc_vu', auto_m.get('chuc_vu'))
                    pb_col = m.get('phong_ban', auto_m.get('phong_ban'))
                    cols_to_get = [ma_col]
                    if cv_col and cv_col in df.columns: cols_to_get.append(cv_col)
                    if pb_col and pb_col in df.columns: cols_to_get.append(pb_col)
                    try:
                        uniq = df[cols_to_get].drop_duplicates(subset=[ma_col])
                        for _, r in uniq.iterrows():
                            ma = str(r[ma_col]).strip()
                            if ma and ma.lower() not in ['nan', 'none']:
                                raw_pb_cv[ma] = {
                                    'cv': str(r[cv_col]).strip() if cv_col and cv_col in df.columns and pd.notnull(r[cv_col]) and str(r[cv_col]).lower() not in ['nan', 'none'] else "",
                                    'pb': str(r[pb_col]).strip() if pb_col and pb_col in df.columns and pd.notnull(r[pb_col]) and str(r[pb_col]).lower() not in ['nan', 'none'] else ""
                                }
                    except Exception as e: logger.warning(f"Lỗi: {e}", exc_info=True); pass

            if not emp_dict_display:
                st.info(t("auto_text_page_history_29"))
            else:
                emp_data_list = []
                for ma, ten in sorted(emp_dict_display.items()):
                    pb = st.session_state.edited_emps.get(ma, {}).get('pb', "") if 'edited_emps' in st.session_state else ""
                    cv = st.session_state.edited_emps.get(ma, {}).get('cv', "") if 'edited_emps' in st.session_state else ""
                    
                    if not pb or not cv:
                        for manual_e in st.session_state.manual_emps:
                            if manual_e['ma'] == ma:
                                if not pb: pb = manual_e.get('pb', "")
                                if not cv: cv = manual_e.get('cv', "")
                                
                    if not pb or not cv:
                        if ma in db_pb_cv:
                            if not pb: pb = db_pb_cv[ma]['pb']
                            if not cv: cv = db_pb_cv[ma]['cv']
                            
                    if not pb or not cv:
                        if ma in raw_pb_cv:
                            if not pb: pb = raw_pb_cv[ma]['pb']
                            if not cv: cv = raw_pb_cv[ma]['cv']

                    ma_nv_col = "Mã NV" if is_vi else "社員ID"
                    ten_nv_col = "Tên NV" if is_vi else "氏名"
                    cv_col = "Chức vụ" if is_vi else "役職"
                    pb_col = "Phòng ban" if is_vi else "部署"
                    emp_data_list.append({ma_nv_col: ma, ten_nv_col: ten, cv_col: cv, pb_col: pb})

                st.markdown("<hr style='margin:0px; padding:0px; border-top: 1px solid #E2E8F0; margin-bottom: 10px;'>", unsafe_allow_html=True)
                df_emp_disp = pd.DataFrame(emp_data_list)
                st.dataframe(df_emp_disp, use_container_width=True, hide_index=True, height=250)

                lbl_del_title = "<div style='margin-top: 10px;'><b>🗑️ Xóa nhân viên</b></div>" if is_vi else "<div style='margin-top: 10px;'><b>🗑️ 社員を削除</b></div>"
                st.markdown(lbl_del_title, unsafe_allow_html=True)
                col_del1, col_del2 = st.columns([3, 1])
                
                lbl_select = "Chọn nhân viên" if is_vi else "社員を選択"
                lbl_placeholder = "--- Chọn nhân viên ---" if is_vi else "--- 社員を選択 ---"
                lbl_btn_del = "Xóa nhân viên" if is_vi else "社員を削除"
                
                del_ma = col_del1.selectbox(lbl_select, options=[""] + [r[ma_nv_col] for r in emp_data_list], format_func=lambda x: f"{x} - {emp_dict_display.get(x, '')}" if x else lbl_placeholder, label_visibility="collapsed")
                if col_del2.button(lbl_btn_del, use_container_width=True):
                    if del_ma:
                        ma = del_ma
                        st.session_state.deleted_emps.add(ma)
                        st.session_state.manual_emps = [e for e in st.session_state.manual_emps if e['ma'] != ma]
                        try:
                            conn = sqlite3.connect(DB_FILE)
                            conn.execute("DELETE FROM employees WHERE ma_nv=?", (ma,))
                            conn.commit(); conn.close()
                        except Exception as e: logger.warning(f"Lỗi: {e}", exc_info=True); pass
                        st.rerun()

        # 5.5 Quản lý Form mẫu
        with st.expander(t("auto_text_page_history_35")):
            with st.form("form_upload_templates_settings", clear_on_submit=True):
                uploaded_templates = st.file_uploader(
                    t("auto_text_page_history_36"), 
                    type=['xlsx', 'xls'], 
                    accept_multiple_files=True
                )
                if st.form_submit_button(t("auto_text_page_history_37"), use_container_width=True):
                    if uploaded_templates:
                        import os
                        if not os.path.exists("templates"): os.makedirs("templates")
                        for uploaded_file in uploaded_templates:
                            with open(os.path.join("templates", uploaded_file.name), "wb") as f:
                                f.write(uploaded_file.getbuffer())
                        st.rerun()

            import os
            if not os.path.exists("templates"): os.makedirs("templates")
            t_files = os.listdir("templates")
            if t_files:
                st.markdown("<hr style='margin:10px 0'>", unsafe_allow_html=True)
                st.markdown(t("auto_text_page_history_38"))
                for tf in t_files:
                    col_tf1, col_tf2 = st.columns([5, 1])
                    col_tf1.markdown(f"<span style='font-size:13px; word-break: break-all;'>{tf}</span>", unsafe_allow_html=True)
                    st.markdown(f"""<style>
                    div.st-key-del_tpl_set_{tf} button {{
                        padding: 0px 10px !important;
                        height: 32px !important;
                        min-height: 32px !important;
                        line-height: 32px !important;
                        font-size: 13px !important;
                        width: 100% !important;
                    }}
                    </style>""", unsafe_allow_html=True)
                    if col_tf2.button("❌", key=f"del_tpl_set_{tf}", help="Xóa"):
                        os.remove(os.path.join("templates", tf))
                        st.rerun()





