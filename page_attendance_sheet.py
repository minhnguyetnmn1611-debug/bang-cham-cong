import streamlit as st
import pandas as pd
from datetime import datetime
from theme import get_theme
from db import get_company_emp_options, get_field_checkins, DB_FILE
import sqlite3
import re

def render_attendance_sheet_page():
    is_vi = (st.session_state.get('lang', 'vi') == 'vi')
    T = get_theme(st.session_state.get('theme_mode', 'light'))
    
    st.markdown(f"""
    <div style="background: {T['bg_card']}; padding: 24px; border-radius: 16px; margin-bottom: 24px; border: 1.5px solid {T['border']}; box-shadow: {T['shadow']};">
        <h2 style="margin: 0; font-family: 'Plus Jakarta Sans', sans-serif; font-weight: 800; font-size: 24px; color: {T['text_primary']};">📄 { 'Bảng Công Cá Nhân' if is_vi else '個人の勤怠表' }</h2>
        <p style="margin: 8px 0 0 0; font-size: 14.5px; color: {T['text_secondary']};">{'Xem và tải xuống bảng chấm công chi tiết hàng tháng của bạn.' if is_vi else '毎月の詳細な勤怠表を確認・ダウンロードします。'}</p>
    </div>
    """, unsafe_allow_html=True)

    # 1. Select Employee & Month
    col1, col2 = st.columns(2)
    with col1:
        months = [f"Tháng {i}/2026" if is_vi else f"2026年{i}月" for i in range(1, 13)]
        selected_month = st.selectbox("Tháng (Kỳ công)" if is_vi else "勤怠月", months, index=datetime.now().month - 1)
    with col2:
        emp_options = get_company_emp_options(st.session_state.get('lang', 'vi'))
        selected_emp = st.selectbox("Nhân viên" if is_vi else "従業員", emp_options)

    st.markdown("<hr style='margin: 20px 0;'>", unsafe_allow_html=True)

    emp_name = selected_emp.split(" - ")[1] if " - " in selected_emp else selected_emp
    emp_code = selected_emp.split(" - ")[0] if " - " in selected_emp else selected_emp

    # 2. Extract Data from DB
    month_match = re.search(r'\d+', selected_month)
    month_num = int(month_match.group()) if month_match else 6
    month_str = f"/{month_num:02d}/2026"

    total_days = 0
    ot_hours = 0.0
    leave_days = 0.0
    late_count = 0

    try:
        conn = sqlite3.connect(DB_FILE)
        df_records = pd.read_sql_query(
            "SELECT * FROM records WHERE TRIM(ma_nv) = ? AND ngay LIKE ?", 
            conn, 
            params=(emp_code.strip(), f"%{month_str}")
        )
        conn.close()

        if not df_records.empty:
            total_days = len(df_records[df_records['tong_gio'] > 0])
            ot_hours = float(df_records['ot'].sum())
            
            # Tính số lần đi muộn/về sớm, bỏ qua những ngày có ghi chú "có phép" (đã xin phép)
            late_mask = (df_records['di_tre'] > 0) | (df_records['ve_som'] > 0)
            permitted_mask = df_records['ghi_chu'].str.contains('có phép|co phep', case=False, na=False)
            late_count = len(df_records[late_mask & ~permitted_mask])
            
            # Tính số ngày nghỉ phép
            leave_days = float(len(df_records[df_records['ghi_chu'].str.contains('nghỉ|phép|leave', case=False, na=False)]))
    except Exception as e:
        st.error(f"Lỗi truy xuất dữ liệu: {e}")

    if total_days == 0 and ot_hours == 0 and late_count == 0:
        st.warning("⚠️ " + ("Không tìm thấy dữ liệu chấm công nào trong tháng này! Hãy chọn đúng tháng (ví dụ: Tháng 5 hoặc Tháng 6) hoặc đảm bảo bạn đã bấm 'Lưu Database' ở trang Xử lý Bảng chấm công." if is_vi else "この月の勤怠データが見つかりません。月を正しく選択するか、打刻データ処理ページで「データベースに保存」を押してください。"))

    if late_count == 0:
        rating = "TUYỆT VỜI / EXCELLENT" if is_vi else "素晴らしい / EXCELLENT"
        rating_color = "#10B981"
    elif late_count <= 2:
        rating = "TỐT / GOOD" if is_vi else "良好 / GOOD"
        rating_color = "#EC4899"
    else:
        rating = "CẦN CẢI THIỆN / NEEDS IMPROVEMENT" if is_vi else "要改善 / NEEDS IMPROVEMENT"
        rating_color = "#F59E0B"
        
    try:
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute("SELECT phong_ban FROM employees WHERE ma_nv=?", (emp_code,))
        row = c.fetchone()
        department_val = row[0].strip() if row and row[0] else "Khối Kỹ Thuật"
        conn.close()
        dept_map = {"Khối Kỹ Thuật": "技術部", "Nhân sự": "人事部", "Kế toán": "経理部", "Kinh doanh": "営業部", "Hành chính": "総務部", "IT": "IT部"}
        department_display = department_val if is_vi else dept_map.get(department_val, department_val)
    except:
        department_display = "Khối Kỹ Thuật" if is_vi else "技術部"

    t_title = "BẢNG CÔNG CÁ NHÂN / ATTENDANCE SHEET" if is_vi else "個人の勤怠表 / ATTENDANCE SHEET"
    t_name = "Họ và tên / Name:" if is_vi else "氏名 / Name:"
    t_emp_code = "Mã NV / Emp Code:" if is_vi else "社員ID / Emp Code:"
    t_dept = "Phòng ban / Department:" if is_vi else "部署 / Department:"
    t_std_days = "Kỳ công chuẩn / Standard Days:" if is_vi else "標準稼働日数 / Standard Days:"
    t_desc = "MÔ TẢ / DESCRIPTION" if is_vi else "項目 / DESCRIPTION"
    t_qty = "SỐ LƯỢNG / QUANTITY" if is_vi else "数量 / QUANTITY"
    t_act_days = "Tổng ngày đi làm / Actual Work Days" if is_vi else "総出勤日数 / Actual Work Days"
    t_ot_hrs = "Tổng giờ tăng ca / Total OT Hours" if is_vi else "総残業時間 / Total OT Hours"
    t_leave = "Số ngày nghỉ phép / Leave Days Taken" if is_vi else "取得有給休暇日数 / Leave Days Taken"
    t_late = "Số lần đi muộn, về sớm / Lateness" if is_vi else "遅刻・早退回数 / Lateness"
    t_rating = "ĐÁNH GIÁ / RATING" if is_vi else "評価 / RATING"
    
    t_unit_day = "ngày" if is_vi else "日"
    t_unit_hr = "giờ" if is_vi else "時間"
    t_unit_time = "lần" if is_vi else "回"
    
    t_footer1 = "Đây là tài liệu nội bộ của công ty. Vui lòng kiểm tra và phản hồi nếu có sai sót." if is_vi else "これは社内文書です。内容を確認し、誤りがある場合はお知らせください。"
    t_footer2 = "This is an internal document. Please review and report any discrepancies."
    t_footer3 = "Ngày xuất:" if is_vi else "発行日:"

    # 3. HTML Payslip Template
    html_payslip = f"""
    <html>
    <head>
    <meta charset="utf-8">
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;600;800&display=swap" rel="stylesheet">
    <style>
        body {{ font-family: 'Plus Jakarta Sans', sans-serif; background: {T['bg_app']}; padding: 40px; display: flex; justify-content: center; color: {T['text_primary']}; margin: 0; }}
        .payslip-container {{ background: {T['bg_card']}; padding: 40px; border-radius: 12px; box-shadow: {T['shadow']}; width: 100%; max-width: 600px; border: 1px solid {T['border']}; }}
        .header {{ display: flex; justify-content: space-between; border-bottom: 2px solid {T['border']}; padding-bottom: 20px; margin-bottom: 20px; }}
        .logo {{ font-size: 24px; font-weight: 800; color: {T['primary']}; }}
        .company-info {{ text-align: right; font-size: 12px; color: {T['text_tertiary']}; }}
        .title {{ text-align: center; font-size: 20px; font-weight: 800; color: {T['text_primary']}; margin-bottom: 24px; letter-spacing: 1px; text-transform: uppercase; }}
        .title span {{ font-size: 14px; color: {T['text_tertiary']}; font-weight: 600; }}
        .emp-info {{ margin-bottom: 30px; background: {T['bg_card_hover']}; padding: 16px; border-radius: 8px; border: 1px solid {T['border']}; }}
        .emp-info table {{ width: 100%; font-size: 14px; color: {T['text_secondary']}; border-collapse: collapse; }}
        .emp-info td {{ padding: 6px 0; }}
        .emp-info td:nth-child(even) {{ font-weight: 600; text-align: right; color: {T['text_primary']}; }}
        .salary-table {{ width: 100%; border-collapse: collapse; margin-bottom: 30px; }}
        .salary-table th {{ background: {T['primary']}; color: white; padding: 12px; text-align: left; font-size: 13px; font-weight: 600; }}
        .salary-table td {{ padding: 12px; border-bottom: 1px solid {T['border']}; font-size: 14px; color: {T['text_secondary']}; }}
        .salary-table td.amount {{ text-align: right; font-weight: 600; color: {T['text_primary']}; }}
        .net-pay {{ background: {rating_color}; color: white; padding: 16px; border-radius: 8px; display: flex; justify-content: space-between; align-items: center; }}
        .net-pay-title {{ font-size: 16px; font-weight: 600; text-transform: uppercase; }}
        .net-pay-amount {{ font-size: 20px; font-weight: 800; text-transform: uppercase; }}
        .footer {{ text-align: center; font-size: 11px; color: {T['text_tertiary']}; margin-top: 40px; border-top: 1px dashed {T['border']}; padding-top: 20px; }}
        
        @media print {{
            body {{ background: white !important; color: black !important; padding: 0 !important; }}
            .payslip-container {{ background: white !important; box-shadow: none !important; border: none !important; padding: 0 !important; }}
            .title, .emp-info td:nth-child(even), .salary-table td.amount {{ color: black !important; }}
            .title span, .company-info, .footer {{ color: #64748b !important; }}
            .emp-info {{ background: #f8fafc !important; border: 1px solid #e2e8f0 !important; }}
            .emp-info table {{ color: #334155 !important; }}
            .salary-table th {{ background: #0f172a !important; color: white !important; }}
            .salary-table td {{ color: #334155 !important; border-bottom: 1px solid #e2e8f0 !important; }}
        }}
    </style>
    </head>
    <body>
        <div class="payslip-container">
            <div class="header">
                <div class="logo">VIET.MOS</div>
                <div class="company-info">
                    <strong>VIET.MOS COMPANY LIMITED</strong><br>
                    Tầng 4, Tòa nhà ABC, Hà Nội<br>
                    MST: 0101234567
                </div>
            </div>
            <div class="title">{t_title}<br><span style="font-size: 14px; color: #64748b; font-weight: 600;">{selected_month}</span></div>
            
            <div class="emp-info">
                <table>
                    <tr><td>{t_name}</td><td>{emp_name}</td></tr>
                    <tr><td>{t_emp_code}</td><td>{emp_code}</td></tr>
                    <tr><td>{t_dept}</td><td>{department_display}</td></tr>
                    <tr><td>{t_std_days}</td><td>22 {t_unit_day}</td></tr>
                </table>
            </div>

            <table class="salary-table">
                <tr><th>{t_desc}</th><th style="text-align: right;">{t_qty}</th></tr>
                <tr><td>{t_act_days}</td><td class="amount">{total_days} {t_unit_day}</td></tr>
                <tr><td>{t_ot_hrs}</td><td class="amount">{ot_hours} {t_unit_hr}</td></tr>
                <tr><td>{t_leave}</td><td class="amount" style="color: #ec4899;">{leave_days} {t_unit_day}</td></tr>
                <tr><td>{t_late}</td><td class="amount" style="color: #ef4444;">{late_count} {t_unit_time}</td></tr>
            </table>

            <div class="net-pay">
                <div class="net-pay-title">{t_rating}</div>
                <div class="net-pay-amount">{rating}</div>
            </div>

            <div class="footer">
                {t_footer1}<br>
                {t_footer2}<br>
                {t_footer3} {datetime.now().strftime('%d/%m/%Y %H:%M')}
            </div>
        </div>
    </body>
    </html>
    """

    # 4. Show HTML in Streamlit
    st.markdown(f"<div style='font-weight: 800; font-size: 16px; margin-bottom: 12px; color: {T['text_primary']};'>📄 {'Xem trước Bảng công' if is_vi else '勤怠表のプレビュー'}</div>", unsafe_allow_html=True)
    
    st.components.v1.html(html_payslip, height=750, scrolling=True)

    # 5. Download Button
    st.markdown("<div style='height: 16px;'></div>", unsafe_allow_html=True)
    col_dl, _ = st.columns([1, 1])
    with col_dl:
        st.download_button(
            label="📥 Tải Xuống Bảng Công (HTML)" if is_vi else "📥 勤怠表をダウンロード (HTML)",
            data=html_payslip,
            file_name=f"Attendance_{emp_name.replace(' ', '_')}_{selected_month.replace('/', '_')}.html",
            mime="text/html",
            type="primary",
            use_container_width=True
        )
    st.info("💡 Mẹo: Bạn có thể mở file HTML vừa tải xuống trên trình duyệt Chrome/Edge và bấm Ctrl+P (hoặc Print) để lưu lại dưới dạng file PDF chuẩn đẹp!" if is_vi else "💡 ヒント: ダウンロードしたHTMLファイルをChrome/Edgeで開き、Ctrl+P (Print) を押してPDFとして保存できます！")
