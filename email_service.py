import streamlit as st
import time

def send_email_notification(to_email, subject, html_content):
    """
    Mock Email Service
    Trong thực tế, hàm này sẽ sử dụng smtplib và email.mime để gửi email qua SMTP server.
    Ở phiên bản hiện tại, hệ thống sẽ mô phỏng việc gửi email và hiển thị thông báo Toast/Alert
    trên giao diện Streamlit để người dùng (HR/Quản lý) biết rằng luồng Workflow đang hoạt động.
    """
    # 1. Giả lập delay gửi mạng
    time.sleep(1)
    
    # 2. Hiển thị Toast message góc dưới
    st.toast(f"📧 Đã gửi Email tới **{to_email}**\nChủ đề: {subject}", icon="✅")
    
    # 3. (Tùy chọn) Lưu log email vào session_state để hiển thị ở trang Admin nếu cần
    if 'email_logs' not in st.session_state:
        st.session_state.email_logs = []
    
    st.session_state.email_logs.append({
        "time": time.strftime("%Y-%m-%d %H:%M:%S"),
        "to": to_email,
        "subject": subject,
        "content": html_content
    })
    
    return True
