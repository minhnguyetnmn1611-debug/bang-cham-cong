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
    bg_banner = "linear-gradient(135deg, #FDF8ED 0%, #F4ECD8 100%)" if is_sepia else "linear-gradient(135deg, #F8FAFC 0%, #F1F5F9 100%)"
    banner_border = "1.5px solid rgba(217, 119, 6, 0.45)" if is_sepia else "1.5px solid #E2E8F0"
    title_color = "#451A03" if is_sepia else "#1E293B"
    subtitle_color = "#78350F" if is_sepia else "#64748B"
    bg_card = "rgba(254, 252, 232, 0.96)" if is_sepia else "rgba(255, 255, 255, 0.98)"

    # =========================================================================
    # DESIGN TOKENS & 2-TONE CARD LAYERING SYSTEM (UNIFIED BORDER-RADIUS: 20PX)
    # =========================================================================

    # 1. THẺ CHÍNH THẾ HỆ MỚI (HERO PRIMARY CARDS): BO GÓC CHUẨN 20PX + VIỀN CYAN PHÁT SÁNG NỔI BẬT
    hero_card_css = "background: #FFFFFF; border: 1px solid #E2E8F0; border-top: 2.5px solid #F472B6; border-radius: 20px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03);"

    # 2. THẺ PHỤ CHÌM BÊN DƯỚI (SECONDARY SUPPORTING CONTAINERS): BO GÓC CHUẨN 20PX + NỀN TỐI HƠN CHÌM MỜ
    secondary_card_css = "background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 20px; box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.1), 0 1px 2px 0 rgba(0, 0, 0, 0.06);"

    # =========================================================================
    # UNIFIED VECTOR ICON SET & PROFESSIONAL STATUS INDICATORS (LUCIDE STANDARD)
    # =========================================================================

    icon_users_svg = """<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="#F472B6" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" style="vertical-align: middle; filter: drop-shadow(0 0 6px rgba(244,114,182,0.5));"><path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M22 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>"""
    icon_rocket_svg = """<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="#F472B6" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" style="vertical-align: middle; filter: drop-shadow(0 0 6px rgba(244,114,182,0.5));"><path d="M4.5 16.5c-1.5 1.26-2 5-2 5s3.74-.5 5-2c.71-.71 1.26-1.5 1.5-2.5l2 2s1.5-1.5 0-3l-2-2c1-.24 1.79-.79 2.5-1.5 1.5-1.26 5-2 5-2s-.5 3.74-2 5c-.71.71-1.5 1.26-2.5 1.5l2 2s1.5-1.5 0-3l-2-2z"/><path d="M12 15l-3-3"/><circle cx="15" cy="9" r="2"/></svg>"""
    icon_calendar_svg = """<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#F472B6" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" style="vertical-align: middle;"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>"""
    icon_zap_svg = """<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#F472B6" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" style="vertical-align: middle; filter: drop-shadow(0 0 6px rgba(244,114,182,0.5));"><polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/></svg>"""
    icon_hourglass_svg = """<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#F59E0B" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" style="vertical-align: middle; filter: drop-shadow(0 0 6px rgba(245,158,11,0.5));"><path d="M5 22h14"/><path d="M5 2h14"/><path d="M17 22v-4.172a2 2 0 0 0-.586-1.414L12 12l-4.414 4.414A2 2 0 0 0 7 17.828V22"/><path d="M7 2v4.172a2 2 0 0 0 .586 1.414L12 12l4.414-4.414A2 2 0 0 0 17 6.172V2"/></svg>"""
    icon_arrow_up_svg = """<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#F472B6" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" style="vertical-align: middle;"><polyline points="23 6 13.5 15.5 8.5 10.5 1 18"/><polyline points="17 6 23 6 23 12"/></svg>"""

    # Badge trạng thái dự án MOS
    if active_projects > 0:
        badge_mos_html = f"""<div style="color: #F472B6; font-size: 12px; font-weight: 700; display: inline-flex; align-items: center; gap: 6px;">{icon_arrow_up_svg} <span>{lbl_mos_status}</span></div>"""
    else:
        badge_mos_html = f"""<div style="color: #F59E0B; font-size: 12px; font-weight: 700; display: inline-flex; align-items: center; gap: 6px;">{icon_hourglass_svg} <span>{lbl_mos_status}</span></div>"""

    # Badge trạng thái chấm công
    if st.session_state.get('df_raw') is not None:
        badge_chart_status_html = f"""<div style="color: #F472B6; font-size: 14px; font-weight: 700; display: inline-flex; align-items: center; gap: 6px;"><span style="display: inline-block; width: 8px; height: 8px; background: #F472B6; border-radius: 50%; box-shadow: 0 0 10px #F472B6;"></span> <span>{status_sub}</span></div>"""
    else:
        badge_chart_status_html = f"""<div style="color: #F59E0B; font-size: 14px; font-weight: 700; display: inline-flex; align-items: center; gap: 6px;">{icon_hourglass_svg} <span>{status_sub}</span></div>"""

    # 1. TIÊU ĐỀ CHÍNH CỦA ỨNG DỤNG (CHUYỂN LẠI THÀNH HÀNG RIÊNG FULL-WIDTH, BỎ KHUNG VIỀN ĐỂ HÒA VÀO NỀN)
    main_header_title = "BẢNG ĐIỀU KHIỂN V.MOS" if is_vi else "V.MOS ダッシュボード"
    main_header_desc = "Trung tâm Giám sát Điều hành, Kê khai Công số & Tự động hóa Chấm công AI." if is_vi else "AI自動勤怠管理およびオペレーション監視センター"
    
    st.markdown(f"""<div style="padding: 10px 0px; margin-top: -130px; margin-bottom: 32px; position: relative; text-align: left;">
<h1 style="color: #64748B; font-size: 23px; font-weight: 800; margin: 0; font-family: 'Plus Jakarta Sans', sans-serif; letter-spacing: 0.5px; line-height: 1.35; background: linear-gradient(135deg, #1E293B 0%, #475569 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent;">{main_header_title}</h1>
<p style="color: #64748B; font-size: 14px; margin-top: 6px; margin-bottom: 0; font-weight: 500; letter-spacing: 0.2px;">{main_header_desc}</p>
</div>""", unsafe_allow_html=True)

    # (Removed KPI cards per user request)

    st.markdown("<div style='height: 16px;'></div>", unsafe_allow_html=True)

    # 3. MÔ-ĐỦN KHỐI GIỮA: TRẠNG THÁI CHỨC NĂNG CHÍNH (2-COLUMN GRID)
    col1, col2 = st.columns(2, gap="small")
    
    icon_calendar_48 = """<svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="#64748B" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"></rect><line x1="16" y1="2" x2="16" y2="6"></line><line x1="8" y1="2" x2="8" y2="6"></line><line x1="3" y1="10" x2="21" y2="10"></line></svg>"""
    icon_clock_48 = """<svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="#64748B" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><polyline points="12 6 12 12 16 14"></polyline></svg>"""
    icon_calendar_small = """<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" style="vertical-align: middle;"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"></rect><line x1="16" y1="2" x2="16" y2="6"></line><line x1="8" y1="2" x2="8" y2="6"></line><line x1="3" y1="10" x2="21" y2="10"></line></svg>"""
    icon_rocket_small = """<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" style="vertical-align: middle;"><path d="M4.5 16.5c-1.5 1.26-2 5-2 5s3.74-.5 5-2c.71-.71 1.26-1.5 1.5-2.5l2 2s1.5-1.5 0-3l-2-2c1-.24 1.79-.79 2.5-1.5 1.5-1.26 5-2 5-2s-.5 3.74-2 5c-.71.71-1.5 1.26-2.5 1.5l2 2s1.5-1.5 0-3l-2-2z"></path><path d="M12 15l-3-3"></path><circle cx="15" cy="9" r="2"></circle></svg>"""

    circle_blue = f"""<div style="position: absolute; right: -24px; top: -24px; width: 120px; height: 120px; border-radius: 50%; background: rgba(236, 72, 153, 0.04); border: 1px solid rgba(236, 72, 153, 0.08); pointer-events: none;"></div><div style="position: absolute; right: 10px; top: 10px; width: 80px; height: 80px; border-radius: 50%; border: 1px solid rgba(236, 72, 153, 0.06); pointer-events: none;"></div>"""
    circle_orange = f"""<div style="position: absolute; right: -24px; top: -24px; width: 120px; height: 120px; border-radius: 50%; background: rgba(245, 158, 11, 0.04); border: 1px solid rgba(245, 158, 11, 0.08); pointer-events: none;"></div><div style="position: absolute; right: 10px; top: 10px; width: 80px; height: 80px; border-radius: 50%; border: 1px solid rgba(245, 158, 11, 0.06); pointer-events: none;"></div>"""

    has_df = st.session_state.get('df_raw') is not None
    if has_df:
        b1_title = f"{latest_period_str}"
        b1_desc = f"Đã ghi nhận dữ liệu chấm công. Thực hiện xem chi tiết tại mục <b>[Chấm công]</b> ở thanh bên trái." if is_vi else f"勤怠データが記録されました。左側のメニューの<b>[勤怠データ処理]</b>で詳細を確認してください。"
        b1_pill = f"""<div style="color: #10B981; font-size: 12.5px; font-weight: 700; display: inline-flex; align-items: center; gap: 6px; background: rgba(16, 185, 129, 0.1); padding: 4px 10px; border-radius: 50px; border: 0.5px solid rgba(16,185,129,0.2);">✓ <span>{"Đã cập nhật" if is_vi else "更新済み"}</span></div>"""
    else:
        b1_title = "Chưa có dữ liệu kỳ này" if is_vi else "今期のデータはありません"
        b1_desc = "Vui lòng tải lên tập tin Excel bảng chấm công. Thực hiện tại mục <b>[Chấm công]</b> ở thanh bên trái." if is_vi else "勤怠エクセルファイルをアップロードしてください。左側のメニューの<b>[勤怠データ処理]</b>で実行します。"
        b1_pill = f"""<div style="color: #F59E0B; font-size: 12.5px; font-weight: 700; display: inline-flex; align-items: center; gap: 6px; background: rgba(245, 158, 11, 0.1); padding: 4px 10px; border-radius: 50px; border: 0.5px solid rgba(245,158,11,0.2);">⏳ <span>{"Đang chờ dữ liệu" if is_vi else "データ待機中"}</span></div>"""

    lbl_b1_head = "XỬ LÝ BẢNG CHẤM CÔNG" if is_vi else "勤怠データ処理"
    html_block1 = f"""
    <div style="background: #FFFFFF; border: 0.5px solid #E2E8F0; border-radius: 12px; padding: 24px; position: relative; overflow: hidden; height: 100%; min-height: 220px; display: flex; flex-direction: column;">
        {circle_blue}
        <div style="color: #64748B; font-size: 12px; font-weight: 700; display: flex; align-items: center; gap: 6px; margin-bottom: 24px; letter-spacing: 0.5px; text-transform: uppercase;">
            {icon_calendar_small} <span>{lbl_b1_head}</span>
        </div>
        <div style="display: flex; gap: 18px; margin-bottom: auto;">
            <div style="width: 48px; height: 48px; border-radius: 50%; background: #F1F5F9; display: flex; align-items: center; justify-content: center; flex-shrink: 0; margin-top: 2px;">
                {icon_calendar_48}
            </div>
            <div>
                <div style="color: #1E293B; font-size: 16px; font-weight: 700; margin-bottom: 6px; font-family: 'Plus Jakarta Sans', sans-serif;">{b1_title}</div>
                <div style="color: #64748B; font-size: 13.5px; line-height: 1.5; margin-bottom: 16px;">{b1_desc}</div>
                {b1_pill}
            </div>
        </div>
    </div>
    """

    has_mos = active_projects > 0
    if has_mos:
        b2_title = f"Đang xử lý {active_projects} dự án" if is_vi else f"{active_projects}プロジェクトを処理中"
        b2_desc = f"Đã đồng bộ giờ làm MOS thành công. Thực hiện xem chi tiết tại mục <b>[Giờ làm MOS]</b> ở thanh bên trái." if is_vi else f"MOSプロジェクト時間が正常に同期されました。左側のメニューの<b>[MOSプロジェクト集計]</b>で詳細を確認してください。"
        b2_pill = f"""<div style="color: #10B981; font-size: 12.5px; font-weight: 700; display: inline-flex; align-items: center; gap: 6px; background: rgba(16, 185, 129, 0.1); padding: 4px 10px; border-radius: 50px; border: 0.5px solid rgba(16,185,129,0.2);">✓ <span>{"Đã cập nhật" if is_vi else "更新済み"}</span></div>"""
    else:
        b2_title = "Chưa có dữ liệu kỳ này" if is_vi else "今期のデータはありません"
        b2_desc = "Vui lòng tải lên báo cáo MOS để tổng hợp. Thực hiện tại mục <b>[Giờ làm MOS]</b> ở thanh bên trái." if is_vi else "MOSレポートをアップロードして集計してください。左側のメニューの<b>[MOSプロジェクト集計]</b>で実行します。"
        b2_pill = f"""<div style="color: #F59E0B; font-size: 12.5px; font-weight: 700; display: inline-flex; align-items: center; gap: 6px; background: rgba(245, 158, 11, 0.1); padding: 4px 10px; border-radius: 50px; border: 0.5px solid rgba(245,158,11,0.2);">⏳ <span>{"Đang chờ dữ liệu" if is_vi else "データ待機中"}</span></div>"""

    lbl_b2_head = "GIỜ LÀM MOS" if is_vi else "MOSプロジェクト集計"
    html_block2 = f"""
    <div style="background: #FFFFFF; border: 0.5px solid #E2E8F0; border-radius: 12px; padding: 24px; position: relative; overflow: hidden; height: 100%; min-height: 220px; display: flex; flex-direction: column;">
        {circle_orange}
        <div style="color: #64748B; font-size: 12px; font-weight: 700; display: flex; align-items: center; gap: 6px; margin-bottom: 24px; letter-spacing: 0.5px; text-transform: uppercase;">
            {icon_rocket_small} <span>{lbl_b2_head}</span>
        </div>
        <div style="display: flex; gap: 18px; margin-bottom: auto;">
            <div style="width: 48px; height: 48px; border-radius: 50%; background: #F1F5F9; display: flex; align-items: center; justify-content: center; flex-shrink: 0; margin-top: 2px;">
                {icon_clock_48}
            </div>
            <div>
                <div style="color: #1E293B; font-size: 16px; font-weight: 700; margin-bottom: 6px; font-family: 'Plus Jakarta Sans', sans-serif;">{b2_title}</div>
                <div style="color: #64748B; font-size: 13.5px; line-height: 1.5; margin-bottom: 16px;">{b2_desc}</div>
                {b2_pill}
            </div>
        </div>
    </div>
    """

    with col1:
        st.markdown(html_block1, unsafe_allow_html=True)
    with col2:
        st.markdown(html_block2, unsafe_allow_html=True)




def render_overview_page():
    t = get_t(st.session_state.get('lang', 'vi'))
    T = get_theme(st.session_state.get('theme_mode', 'light'))
    is_sepia = st.session_state.get('eye_care_sepia', False)

    render_enterprise_dashboard()

