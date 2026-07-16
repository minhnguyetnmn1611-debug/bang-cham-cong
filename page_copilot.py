import streamlit as st
import time
import json
import os
from datetime import datetime
from theme import get_theme

def render_copilot_page():
    is_vi = (st.session_state.get('lang', 'vi') == 'vi')
    T = get_theme(st.session_state.get('theme_mode', 'light'))
    
    st.markdown(f"""
    <div style="background: linear-gradient(135deg, {T['primary']}, {T['accent']}); padding: 24px; border-radius: 16px; margin-bottom: 24px; box-shadow: 0 4px 16px rgba(0,0,0,0.1); color: white;">
        <h2 style="margin: 0; font-family: 'Plus Jakarta Sans', sans-serif; font-weight: 800; font-size: 28px;">🤖 V.MOS Copilot</h2>
        <p style="margin: 8px 0 0 0; font-size: 15px; opacity: 0.9;">{'Trợ lý AI nhân sự thông minh của bạn. Hãy đặt câu hỏi về ngày phép, quy định, hoặc thông tin nhân sự!' if is_vi else 'あなたのスマートなHR AIアシスタント。有給休暇、規定、または人事情報について質問してください！'}</p>
    </div>
    """, unsafe_allow_html=True)

    if "copilot_messages" not in st.session_state:
        st.session_state.copilot_messages = [
            {"role": "assistant", "content": "👋 Xin chào! Tôi là V.MOS Copilot, trợ lý AI của bạn. Tôi có thể giúp gì cho bạn hôm nay?" if is_vi else "👋 こんにちは！私はあなたのAIアシスタント、V.MOS Copilotです。今日はどのようなご用件でしょうか？"}
        ]

    # Display chat messages
    for msg in st.session_state.copilot_messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    # Chat input
    prompt = st.chat_input("Hỏi V.MOS Copilot..." if is_vi else "V.MOS Copilotに質問する...")
    
    if prompt:
        # Add user message to state and display it
        st.session_state.copilot_messages.append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)

        # Simulate AI thinking and response
        with st.chat_message("assistant"):
            message_placeholder = st.empty()
            full_response = ""
            
            # Simulated Rule-Based AI Logic
            p_lower = prompt.lower()
            if any(kw in p_lower for kw in ["nghỉ", "phép", "vacation", "leave"]):
                reply = "🌴 Theo hệ thống, bạn hiện còn **12 ngày phép thường niên** trong năm nay. Để xin nghỉ phép, bạn vui lòng vào mục **Đăng ký nghỉ & OT** ở thanh menu bên trái nhé!" if is_vi else "🌴 システムによると、今年の残りの**年次有給休暇は12日**です。休暇を申請するには、左側のメニューの**休暇＆残業申請**にアクセスしてください。"
            elif any(kw in p_lower for kw in ["ot", "tăng ca", "overtime"]):
                reply = "⏰ Bạn muốn đăng ký OT? Vui lòng sử dụng Form đăng ký OT ở mục **Đăng ký nghỉ & OT**. Lưu ý: Quy định công ty yêu cầu đăng ký OT trước 16:00 chiều hàng ngày để được tính cơm chiều." if is_vi else "⏰ 残業を申請しますか？**休暇＆残業申請**セクションの残業申請フォームをご利用ください。注意：夕食の補助を受けるには、毎日午後4時までに残業を申請するという会社の規定があります。"
            elif any(kw in p_lower for kw in ["ai", "who", "đang nghỉ", "vắng", "absent"]):
                reply = "🔍 Dữ liệu hôm nay cho thấy:\n- **Anh Hà Văn Đạo** (Nửa ngày sáng)\n- **Anh Nguyễn Văn Nam** (Cả ngày)\nNếu bạn cần liên hệ gấp, vui lòng gọi trực tiếp hoặc liên hệ HR." if is_vi else "🔍 今日のデータによると：\n- **Ha Van Dao さん** (午前半休)\n- **Nguyen Van Nam さん** (終日)\n緊急の連絡が必要な場合は、直接電話するか人事部にご連絡ください。"
            elif any(kw in p_lower for kw in ["chào", "hello", "hi"]):
                reply = "Chào bạn! Chúc bạn một ngày làm việc vui vẻ và hiệu quả! 😊" if is_vi else "こんにちは！楽しく生産的な一日をお過ごしください！😊"
            else:
                reply = "🤖 Xin lỗi, tôi là phiên bản thử nghiệm nên chưa thể trả lời câu hỏi này. Hiện tại tôi có thể trả lời các thông tin về: **Ngày phép, Xin OT, và Ai đang nghỉ hôm nay**. Bạn thử hỏi lại nhé!" if is_vi else "🤖 申し訳ありませんが、私はベータ版であり、まだこの質問に答えることができません。現在、**有給休暇、残業申請、今日の欠勤者**に関する情報にはお答えできます。もう一度質問してみてください！"
            
            # Typewriter effect
            for chunk in reply.split():
                full_response += chunk + " "
                time.sleep(0.05)
                message_placeholder.markdown(full_response + "▌")
            
            message_placeholder.markdown(full_response)
        
        # Add assistant response to state
        st.session_state.copilot_messages.append({"role": "assistant", "content": full_response})
