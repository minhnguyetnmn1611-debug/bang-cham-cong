import streamlit as st
import pandas as pd
from datetime import datetime
import base64
from theme import get_theme

def render_payslip_page():
    is_vi = (st.session_state.get('lang', 'vi') == 'vi')
    T = get_theme(st.session_state.get('theme_mode', 'light'))
    
    st.markdown(f"""
    <div style="background: {T['bg_card']}; padding: 24px; border-radius: 16px; margin-bottom: 24px; border: 1.5px solid {T['border']}; box-shadow: {T['shadow']};">
        <h2 style="margin: 0; font-family: 'Plus Jakarta Sans', sans-serif; font-weight: 800; font-size: 24px; color: {T['text_primary']};">💳 { 'Phiếu Lương Cá Nhân (Payslip)' if is_vi else '個人の給与明細 (Payslip)' }</h2>
        <p style="margin: 8px 0 0 0; font-size: 14.5px; color: {T['text_secondary']};">{'Xem và tải xuống phiếu lương hàng tháng của bạn.' if is_vi else '毎月の給与明細を確認・ダウンロードします。'}</p>
    </div>
    """, unsafe_allow_html=True)

    # 1. Select Employee & Month
    col1, col2 = st.columns(2)
    with col1:
        months = [f"Tháng {i}/2026" if is_vi else f"2026年{i}月" for i in range(1, 13)]
        selected_month = st.selectbox("Tháng (Kỳ lương)" if is_vi else "給与月", months, index=datetime.now().month - 1)
    with col2:
        # Giả lập danh sách nhân viên
        emp_list = ["Hà Văn Đạo", "Nguyễn Đăng Hưng", "Hồ Bá Long", "Khuất Tín Nghĩa", "Nguyễn Cảnh Phương", "Lê Văn Nam"]
        selected_emp = st.selectbox("Nhân viên" if is_vi else "従業員", emp_list)

    st.markdown("<hr style='margin: 20px 0;'>", unsafe_allow_html=True)

    # 2. Simulate Payroll Data
    base_salary = 25000000
    ot_hours = 12.5
    ot_rate = 150000
    ot_pay = ot_hours * ot_rate
    allowance = 1500000
    tax = (base_salary + ot_pay + allowance) * 0.1 # 10% tax mock
    net_pay = base_salary + ot_pay + allowance - tax

    # 3. HTML Payslip Template
    html_payslip = f"""
    <html>
    <head>
    <meta charset="utf-8">
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;600;800&display=swap" rel="stylesheet">
    <style>
        body {{ font-family: 'Plus Jakarta Sans', sans-serif; background: #f1f5f9; padding: 40px; display: flex; justify-content: center; }}
        .payslip-container {{ background: white; padding: 40px; border-radius: 12px; box-shadow: 0 10px 30px rgba(0,0,0,0.08); width: 100%; max-width: 600px; }}
        .header {{ display: flex; justify-content: space-between; border-bottom: 2px solid #e2e8f0; padding-bottom: 20px; margin-bottom: 20px; }}
        .logo {{ font-size: 24px; font-weight: 800; color: #db2777; }}
        .company-info {{ text-align: right; font-size: 12px; color: #64748b; }}
        .title {{ text-align: center; font-size: 20px; font-weight: 800; color: #0f172a; margin-bottom: 24px; letter-spacing: 1px; text-transform: uppercase; }}
        .emp-info {{ margin-bottom: 30px; background: #f8fafc; padding: 16px; border-radius: 8px; border: 1px solid #e2e8f0; }}
        .emp-info table {{ width: 100%; font-size: 14px; color: #334155; }}
        .emp-info td {{ padding: 4px 0; }}
        .emp-info td:nth-child(even) {{ font-weight: 600; text-align: right; color: #0f172a; }}
        .salary-table {{ width: 100%; border-collapse: collapse; margin-bottom: 30px; }}
        .salary-table th {{ background: #0f172a; color: white; padding: 12px; text-align: left; font-size: 13px; font-weight: 600; }}
        .salary-table td {{ padding: 12px; border-bottom: 1px solid #e2e8f0; font-size: 14px; color: #334155; }}
        .salary-table td.amount {{ text-align: right; font-weight: 600; color: #0f172a; }}
        .net-pay {{ background: #db2777; color: white; padding: 16px; border-radius: 8px; display: flex; justify-content: space-between; align-items: center; }}
        .net-pay-title {{ font-size: 16px; font-weight: 600; }}
        .net-pay-amount {{ font-size: 24px; font-weight: 800; }}
        .footer {{ text-align: center; font-size: 11px; color: #94a3b8; margin-top: 40px; border-top: 1px dashed #cbd5e1; padding-top: 20px; }}
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
            <div class="title">PHIẾU LƯƠNG / PAYSLIP<br><span style="font-size: 14px; color: #64748b; font-weight: 600;">{selected_month}</span></div>
            
            <div class="emp-info">
                <table>
                    <tr><td>Họ và tên / Name:</td><td>{selected_emp}</td></tr>
                    <tr><td>Chức vụ / Position:</td><td>Kỹ sư thiết kế</td></tr>
                    <tr><td>Phòng ban / Department:</td><td>Khối Kỹ Thuật</td></tr>
                    <tr><td>Số ngày công / Work Days:</td><td>22 ngày</td></tr>
                </table>
            </div>

            <table class="salary-table">
                <tr><th>MÔ TẢ / DESCRIPTION</th><th style="text-align: right;">SỐ TIỀN / AMOUNT (VND)</th></tr>
                <tr><td>Lương cơ bản / Base Salary</td><td class="amount">{base_salary:,.0f}</td></tr>
                <tr><td>Tiền OT / Overtime Pay ({ot_hours} giờ)</td><td class="amount">{ot_pay:,.0f}</td></tr>
                <tr><td>Phụ cấp / Allowances</td><td class="amount">{allowance:,.0f}</td></tr>
                <tr><td>Khấu trừ Thuế & BH / Deductions</td><td class="amount" style="color: #ef4444;">-{tax:,.0f}</td></tr>
            </table>

            <div class="net-pay">
                <div class="net-pay-title">THỰC LÃNH / NET PAY</div>
                <div class="net-pay-amount">{net_pay:,.0f} VND</div>
            </div>

            <div class="footer">
                Đây là tài liệu mật của công ty. Vui lòng không chia sẻ với người ngoài.<br>
                This is a confidential document. Please do not share.<br>
                Ngày xuất: {datetime.now().strftime('%d/%m/%Y %H:%M')}
            </div>
        </div>
    </body>
    </html>
    """

    # 4. Show HTML in Streamlit
    st.markdown(f"<div style='font-weight: 800; font-size: 16px; margin-bottom: 12px; color: {T['text_primary']};'>📄 {'Xem trước Phiếu lương' if is_vi else '給与明細のプレビュー'}</div>", unsafe_allow_html=True)
    
    st.components.v1.html(html_payslip, height=750, scrolling=True)

    # 5. Download Button
    st.markdown("<div style='height: 16px;'></div>", unsafe_allow_html=True)
    col_dl, _ = st.columns([1, 1])
    with col_dl:
        st.download_button(
            label="📥 Tải Xuống Phiếu Lương (HTML)" if is_vi else "📥 給与明細をダウンロード (HTML)",
            data=html_payslip,
            file_name=f"Payslip_{selected_emp.replace(' ', '_')}_{selected_month.replace('/', '_')}.html",
            mime="text/html",
            type="primary",
            use_container_width=True
        )
    st.info("💡 Mẹo: Bạn có thể mở file HTML vừa tải xuống trên trình duyệt Chrome/Edge và bấm Ctrl+P (hoặc Print) để lưu lại dưới dạng file PDF chuẩn đẹp!" if is_vi else "💡 ヒント: ダウンロードしたHTMLファイルをChrome/Edgeで開き、Ctrl+P (Print) を押してPDFとして保存できます！")
