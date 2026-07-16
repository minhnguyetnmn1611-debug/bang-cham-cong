import streamlit as st
import pandas as pd
import datetime
import sqlite3
from theme import get_theme
from translations import get_t, translate_name
from email_service import send_email_notification
from db import DB_FILE, get_company_emp_options

@st.cache_data(ttl=15, show_spinner=False)
def fetch_firebase_leave_ot():
    try:
        from firebase_config import get_db_ref
        ref = get_db_ref("leave_ot_requests")
        return ref.get() or {}
    except Exception:
        return {}

def render_leave_ot_page():
    t = get_t(st.session_state.get('lang', 'vi'))
    T = get_theme(st.session_state.get('theme_mode', 'light'))
    is_sepia = st.session_state.get('eye_care_sepia', False)

    is_vi = (st.session_state.get('lang', 'vi') == 'vi')
    is_sepia = st.session_state.get('eye_care_sepia', False)
    from streamlit_autorefresh import st_autorefresh
    st.markdown("""<style>
    iframe[title="streamlit_autorefresh.streamlit_autorefresh"] { display: none !important; height: 0 !important; width: 0 !important; }
    div[data-testid="stFrame"]:has(iframe[title="streamlit_autorefresh.streamlit_autorefresh"]) { display: none !important; }
    div[data-testid="element-container"]:has(iframe[title="streamlit_autorefresh.streamlit_autorefresh"]) { display: none !important; height: 0 !important; }
    </style>""", unsafe_allow_html=True)
    st_autorefresh(interval=15000, limit=None, key="leave_ot_refresh")
    
    fb_data = fetch_firebase_leave_ot()
    st.session_state.pending_hr_approvals = list(fb_data.values()) if isinstance(fb_data, dict) else []

    bg_banner = "linear-gradient(135deg, rgba(254, 252, 232, 0.95) 0%, rgba(253, 246, 178, 0.9) 100%)" if is_sepia else "linear-gradient(135deg, rgba(255,255,255,0.92) 0%, rgba(255,255,255,0.85) 100%)"
    border_banner = f"1.5px solid {T['border']}"
    shadow_banner = "0 10px 30px rgba(120, 53, 15, 0.15)" if is_sepia else "0 10px 30px rgba(14,165,233,0.25)"
    title_color = T["primary"]
    text_color = T["text_secondary"]

    st.markdown(f"""
    <div style="background: {bg_banner}; backdrop-filter: blur(20px); padding: 30px; border-radius: 20px; box-shadow: {shadow_banner}; margin-bottom: 24px; border: {border_banner};">
        <h1 style="color: {title_color}; font-size: 26px; margin: 0; font-family: Plus Jakarta Sans, Inter, sans-serif; font-weight: 800;">🌴 Đăng ký Nghỉ phép & Tăng ca (OT)</h1>
        <p style="color: {text_color}; font-size: 14.5px; margin-top: 6px; margin-bottom: 0;">Cổng thông tin phân quyền định hướng: Nhân viên tự đăng ký gửi yêu cầu & HR Quản lý xét duyệt trực tuyến.</p>
    </div>
    """ if is_vi else f"""
    <div style="background: {bg_banner}; backdrop-filter: blur(20px); padding: 30px; border-radius: 20px; box-shadow: {shadow_banner}; margin-bottom: 24px; border: {border_banner};">
        <h1 style="color: {title_color}; font-size: 26px; margin: 0; font-family: Plus Jakarta Sans, Inter, sans-serif; font-weight: 800;">🌴 休暇＆残業(OT)の申請登録</h1>
        <p style="color: {text_color}; font-size: 14.5px; margin-top: 6px; margin-bottom: 0;">従業員による申請と人事マネージャーによる承認権限を明確に分離したワークフロー。</p>
    </div>
    """, unsafe_allow_html=True)
    
    emp_options = get_company_emp_options(st.session_state.lang)

    # --- ROLE SWITCHER BOX ---
    title_text = t("auto_text_page_leave_ot_2")
    desc_text = t("auto_text_page_leave_ot_3")
    st.markdown(f"""
    <div style="background: {T['bg_card_hover']}; padding: 16px 20px; border-radius: 16px; border: 1.5px solid {'#D97706' if is_sepia else '#CBD5E1'}; margin-bottom: 24px;">
        <b style="color: {T['text_primary']}; font-size: 15px;">🔑 {title_text}</b>
        <div style="color: {T['text_secondary']}; font-size: 13px; margin-top: 4px;">{desc_text}</div>
    </div>
    """, unsafe_allow_html=True)
    
    role_mode = st.radio(
        t("auto_text_page_leave_ot_4"),
        ["👤 Chế độ Nhân viên (Tự gửi đơn xin Phép/OT lên Bảng thông báo HR)", "🛡️ Chế độ Quản lý / HR Admin (Duyệt đơn từ nhân viên & Nhập trực tiếp)"] if is_vi else ["👤 従業員モード (人事へ申請送信)", "🛡️ 人事・管理者モード (承認・直接入力)"],
        horizontal=True,
        key="role_mode_switcher"
    )
    is_emp_mode = role_mode.startswith("👤")

    if is_emp_mode:
        st.info(t("auto_text_page_leave_ot_5"))
        my_account = st.selectbox(t("auto_text_page_leave_ot_6"), emp_options, key="emp_my_acc")
        
        tab_l, tab_o = st.tabs([t("auto_text_page_leave_ot_7"), t("auto_text_page_leave_ot_8")])
        with tab_l:
            c1, c2 = st.columns(2)
            with c1:
                d_leave = st.date_input(t("auto_text_page_leave_ot_9"), value=datetime.date.today(), key="emp_d_leave")
            with c2:
                reason_l = st.text_input(t("auto_text_page_leave_ot_10"), value=t("auto_text_page_leave_ot_11"), key="emp_r_leave")
            
            if st.button(t("auto_text_page_leave_ot_12"), type="primary", use_container_width=True, key="emp_btn_send_l"):
                emp_code = my_account.split(" - ")[0].strip().upper() if " - " in my_account else my_account
                import uuid
                req_id = f"REQ-{uuid.uuid4().hex[:8].upper()}"
                req_item = {
                    "id": req_id,
                    "type": "Nghỉ phép",
                    "emp": emp_code,
                    "date": d_leave.strftime("%d/%m/%Y"),
                    "reason": reason_l,
                    "status": "⏳ Chờ duyệt"
                }
                st.session_state.pending_hr_approvals.append(req_item)
                try:
                    from firebase_config import get_db_ref
                    get_db_ref(f"leave_ot_requests/{req_id}").set(req_item)
                    fetch_firebase_leave_ot.clear()
                except Exception: pass
                
                
                # Tính năng 3: Gửi email tự động cho HR
                html_body = f"<p>Nhân viên <b>{emp_code}</b> vừa gửi đơn <b>Nghỉ phép</b> cho ngày {d_leave.strftime('%d/%m/%Y')}.</p><p>Lý do: {reason_l}</p><p>Vui lòng đăng nhập hệ thống V.MOS để duyệt đơn.</p>"
                send_email_notification(to_email="hr@vietmos.com", subject=f"[V.MOS] Đơn xin nghỉ phép mới từ {emp_code}", html_content=html_body)
                
                st.success(f"✅ Đã gửi đơn nghỉ phép cho {emp_code} ngày {d_leave.strftime('%d/%m/%Y')}! Vui lòng chuyển sang 'Chế độ Quản lý/HR' để xem bảng thông báo." if is_vi else "✅ 申請を送信しました！人事モードで確認してください。")
        
        with tab_o:
            co1, co2 = st.columns(2)
            with co1:
                d_ot = st.date_input(t("auto_text_page_leave_ot_14"), value=datetime.date.today(), key="emp_d_ot")
            with co2:
                ot_reasons_list = ["Xử lý sự cố khẩn cấp", "Bảo trì công trình", "Chạy thử máy mới", "Họp MOS định kỳ", "Khác"] if is_vi else ["緊急トラブル対応", "設備定期メンテナンス", "新機種テスト運転", "定例MOSミーティング", "その他"]
                reason_o = st.selectbox(t("auto_text_page_leave_ot_15"), ot_reasons_list, key="emp_r_ot")
            
            if st.button(t("auto_text_page_leave_ot_16"), type="primary", use_container_width=True, key="emp_btn_send_o"):
                emp_code = my_account.split(" - ")[0].strip().upper() if " - " in my_account else my_account
                import uuid
                req_id = f"REQ-{uuid.uuid4().hex[:8].upper()}"
                req_item = {
                    "id": req_id,
                    "type": "OT",
                    "emp": emp_code,
                    "date": d_ot.strftime("%d/%m/%Y"),
                    "reason": reason_o,
                    "status": "⏳ Chờ duyệt"
                }
                st.session_state.pending_hr_approvals.append(req_item)
                try:
                    from firebase_config import get_db_ref
                    get_db_ref(f"leave_ot_requests/{req_id}").set(req_item)
                    fetch_firebase_leave_ot.clear()
                except Exception: pass
                
                # Tính năng 3: Gửi email tự động cho HR
                html_body = f"<p>Nhân viên <b>{emp_code}</b> vừa gửi đơn <b>Tăng ca (OT)</b> cho ngày {d_ot.strftime('%d/%m/%Y')}.</p><p>Lý do: {reason_o}</p><p>Vui lòng đăng nhập hệ thống V.MOS để duyệt đơn.</p>"
                send_email_notification(to_email="hr@vietmos.com", subject=f"[V.MOS] Đơn xin OT mới từ {emp_code}", html_content=html_body)
                
                st.success(f"✅ Đã gửi đơn xin Tăng ca (OT) cho {emp_code} ngày {d_ot.strftime('%d/%m/%Y')}! Chờ HR duyệt." if is_vi else "✅ 残業申請を送信しました！人事の承認をお待ちください。")

        st.divider()
        st.markdown(t("auto_text_page_leave_ot_18"))
        my_code = my_account.split(" - ")[0].strip().upper() if " - " in my_account else my_account
        my_reqs = [p for p in st.session_state.pending_hr_approvals if p['emp'] == my_code]
        if my_reqs:
            for r in reversed(my_reqs):
                status_map = {"⏳ Chờ duyệt": "⏳ 承認待ち", "✅ Đã duyệt": "✅ 承認済み", "❌ Từ chối": "❌ 却下"}
                reason_map = {"Nghỉ phép năm / Việc cá nhân": "年次有給休暇 / 私用", "Nghỉ ốm (có giấy viện)": "病気休暇（診断書あり）", "Nghỉ không lương": "無給休暇", "Khác": "その他"}
                type_map = {"Nghỉ phép": "休暇", "Tăng ca (OT)": "残業 (OT)"}
                r_type_disp = r['type'] if is_vi else type_map.get(r['type'], r['type'])
                r_reason_disp = r['reason'] if is_vi else reason_map.get(r['reason'], r['reason'])
                r_status_disp = r['status'] if is_vi else status_map.get(r['status'], r['status'])
                
                badge_bg = "#FEF3C7" if r['status'] == '⏳ Chờ duyệt' else ("#D1FAE5" if "Duyệt" in r['status'] or "duyệt" in r['status'] else "#FEE2E2")
                badge_col = "#D97706" if r['status'] == '⏳ Chờ duyệt' else ("#059669" if "Duyệt" in r['status'] or "duyệt" in r['status'] else "#DC2626")
                st.markdown(f"""
                <div style='background:rgba(255,255,255,0.9); padding:12px 18px; border-radius:12px; margin-bottom:10px; border:1px solid #E2E8F0; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px;'>
                    <div>
                        <b style='color:#0F172A;'>[{r_type_disp}] {t("auto_text_page_leave_ot_19")} {r['date']}</b> — <span style='color:#64748B;'>{t("auto_text_page_leave_ot_20")}: {r_reason_disp}</span>
                    </div>
                    <span style='background:{badge_bg}; color:{badge_col}; padding:4px 12px; border-radius:20px; font-weight:700; font-size:12.5px;'>{r_status_disp}</span>
                </div>
                """, unsafe_allow_html=True)
        else:
            st.markdown(f"""<div style='text-align:center; padding: 40px; background: rgba(255,255,255,0.6); border-radius: 16px; border: 2px dashed #CBD5E1; margin: 20px 0;'>
                <div style='font-size: 54px; margin-bottom: 12px;'>🍃</div>
                <b style="color:#64748B; font-size:16px;">{t("auto_text_page_leave_ot_21")}</b>
            </div>""", unsafe_allow_html=True)

    else:
        # --- HR / MANAGERS MODE ---
        st.markdown(t("auto_text_page_leave_ot_22"))
        pending_reqs = [p for p in st.session_state.get('pending_hr_approvals', []) if p['status'] == '⏳ Chờ duyệt']
        if pending_reqs:
            st.warning(f"⚠️ Đang có **{len(pending_reqs)} yêu cầu** chờ phê duyệt từ nhân viên!" if is_vi else f"⚠️ **{len(pending_reqs)}件**の未承認申請があります。")
            
            c_app_all, c_space = st.columns([1, 4])
            with c_app_all:
                if st.button(t("auto_text_page_leave_ot_24"), type="primary", use_container_width=True, key="btn_appr_all"):
                    st.session_state.show_approve_all_confirm = True
                    
            if st.session_state.get('show_approve_all_confirm', False):
                title_appr = t("auto_text_page_leave_ot_25")
                desc_appr = t("auto_text_page_leave_ot_26")
                st.markdown(f"<div style='background: rgba(254, 252, 232, 0.95); backdrop-filter: blur(10px); padding: 25px; border-radius: 16px; border: 2px dashed #F59E0B; margin: 20px 0; text-align: center; box-shadow: 0 10px 30px rgba(245,158,11,0.2);'><h4 style='color:#B45309; margin-top:0; font-size: 20px;'>{title_appr}</h4><p style='color:#78350F; font-size: 15px;'>{desc_appr}</p></div>", unsafe_allow_html=True)
                ca, cb, cc = st.columns([2,1,1])
                with cb:
                    if st.button(t("auto_text_page_leave_ot_27"), use_container_width=True, type="primary", key="btn_confirm_appr_all"):
                        for req in pending_reqs:
                            req['status'] = "✅ Đã duyệt"
                            if req['type'] == 'Nghỉ phép':
                                st.session_state.manual_leave[(req['emp'], req['date'])] = True
                            else:
                                st.session_state.manual_ot_reason[(req['emp'], req['date'])] = req['reason']
                            try:
                                from firebase_config import get_db_ref
                                get_db_ref(f"leave_ot_requests/{req['id']}/status").set("✅ Đã duyệt")
                                fetch_firebase_leave_ot.clear()
                            except Exception: pass
                        st.session_state.show_approve_all_confirm = False
                        st.success(t("auto_text_page_leave_ot_28"))
                        st.rerun()
                with cc:
                    if st.button(t("auto_text_page_leave_ot_29"), use_container_width=True, key="btn_cancel_appr_all"):
                        st.session_state.show_approve_all_confirm = False
                        st.rerun()
            
            st.markdown("<br>", unsafe_allow_html=True)
            for idx, req in enumerate(pending_reqs):
                is_leave = (req['type'] == 'Nghỉ phép')
                bg_color = "linear-gradient(135deg, rgba(255,255,255,0.95) 0%, rgba(255,241,242,0.9) 100%)" if is_leave else "linear-gradient(135deg, rgba(255,255,255,0.95) 0%, rgba(240,249,255,0.9) 100%)"
                border_color = "#F43F5E" if is_leave else "#0EA5E9"
                icon = "🌴" if is_leave else "⏰"
                type_label = "Nghỉ phép" if is_leave else "Tăng ca (OT)"
                
                status_map = {"⏳ Chờ duyệt": "⏳ 承認待ち", "✅ Đã duyệt": "✅ 承認済み", "❌ Từ chối": "❌ 却下"}
                reason_map = {"Nghỉ phép năm / Việc cá nhân": "年次有給休暇 / 私用", "Nghỉ ốm (có giấy viện)": "病気休暇（診断書あり）", "Nghỉ không lương": "無給休暇", "Khác": "その他"}
                type_map = {"Nghỉ phép": "休暇", "Tăng ca (OT)": "残業 (OT)"}
                type_label_disp = type_label if is_vi else type_map.get(req['type'], req['type'])
                req_reason_disp = req['reason'] if is_vi else reason_map.get(req['reason'], req['reason'])
                req_status_disp = req.get('status', '⏳ Chờ duyệt')
                req_status_disp = req_status_disp if is_vi else status_map.get(req_status_disp, req_status_disp)

                badge_bg = "#10B981" if req.get('status') == '✅ Đã duyệt' else ("#F43F5E" if req.get('status') == '❌ Từ chối' else "#F59E0B")
                st.markdown(f"""
                <div style='background: {bg_color}; backdrop-filter: blur(10px); padding: 20px 24px; border-radius: 16px; border-left: 6px solid {border_color}; margin-bottom: 16px; box-shadow: 0 8px 25px rgba(0,0,0,0.05); transition: transform 0.2s ease, box-shadow 0.2s ease;' onmouseover='this.style.transform="translateY(-2px)"; this.style.boxShadow="0 12px 30px rgba(0,0,0,0.08)";' onmouseout='this.style.transform="translateY(0)"; this.style.boxShadow="0 8px 25px rgba(0,0,0,0.05)";'>
                    <div style='display: flex; justify-content: space-between; align-items: flex-start;'>
                        <div>
                                <span style='font-size: 20px;'>{icon}</span>
                                <div>
                                    <div style='font-weight: 800; color: #0F172A; font-size: 16px;'>{type_label_disp}</div>
                                    <div style='color: #64748B; font-size: 13px; font-weight: 600;'>&#128197; {req['date']} &nbsp;|&nbsp; &#128100; {req['emp']}</div>
                                </div>
                            </div>
                            <div style='background: #F8FAFC; border-radius: 10px; padding: 10px 14px; font-size: 14px; color: #334155; margin-bottom: 12px;'>
                                <strong>&#128221;</strong> {req_reason_disp}
                            </div>
                        </div>
                        <div style='display: flex; gap: 8px; flex-direction: column; min-width: 120px;'>
                            <div style='padding: 6px 10px; background: #10B981; color: white; border-radius: 8px; font-size: 13px; font-weight: 700; text-align: center;'>{req_status_disp}</div>
                        </div>
                    </div>
                """, unsafe_allow_html=True)

                badge_bg = "#10B981" if req['status'] == '✅ Đã duyệt' else ("#F43F5E" if req['status'] == '❌ Từ chối' else "#F59E0B")
                ca_btn, cb_btn = st.columns(2)
                btn_approve_label = "✅ Duyệt" if is_vi else "✅ 承認"
                btn_reject_label = "❌ Từ chối" if is_vi else "❌ 却下"
                msg_approved = "✅ Đã duyệt!" if is_vi else "✅ 承認しました！"
                msg_rejected = "❌ Đã từ chối!" if is_vi else "❌ 却下しました！"
                
                with ca_btn:
                    if req['status'] == '⏳ Chờ duyệt' and st.button(btn_approve_label, key=f'app_{idx}', type='primary', use_container_width=True):
                        req['status'] = '✅ Đã duyệt'
                        if req['type'] == 'Nghỉ phép':
                            st.session_state.manual_leave[(req['emp'], req['date'])] = True
                        else:
                            st.session_state.manual_ot_reason[(req['emp'], req['date'])] = req['reason']
                        try:
                            from firebase_config import get_db_ref
                            get_db_ref(f"leave_ot_requests/{req['id']}/status").set("✅ Đã duyệt")
                            fetch_firebase_leave_ot.clear()
                        except Exception: pass
                        st.success(msg_approved)
                        st.rerun()
                with cb_btn:
                    if req['status'] == '⏳ Chờ duyệt' and st.button(btn_reject_label, key=f'rej_{idx}', use_container_width=True):
                        req['status'] = '❌ Từ chối'
                        try:
                            from firebase_config import get_db_ref
                            get_db_ref(f"leave_ot_requests/{req['id']}/status").set("❌ Từ chối")
                            fetch_firebase_leave_ot.clear()
                        except Exception: pass
                        st.warning('❌ Đã từ chối!')
                        st.rerun()

