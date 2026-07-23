import streamlit as st
import httpx
import time
import json
import urllib.parse
import pandas as pd
from datetime import datetime
from theme import get_theme

def load_saved_api_key():
    import os
    all_keys = []
    if os.path.exists(".streamlit/secrets.toml"):
        try:
            import toml
            with open(".streamlit/secrets.toml", "r", encoding="utf-8") as f:
                sec = toml.load(f)
                if "GEMINI_API_KEYS" in sec:
                    all_keys = sec["GEMINI_API_KEYS"]
                elif "GEMINI_API_KEY" in sec:
                    all_keys = [sec["GEMINI_API_KEY"]]
        except Exception as e:
            from log_config import logger
            logger.error(f"Lỗi khi đọc file secrets.toml: {e}", exc_info=True)
            
    if not all_keys:
        if "GEMINI_API_KEYS" in st.secrets:
            all_keys = st.secrets["GEMINI_API_KEYS"]
        elif "GEMINI_API_KEY" in st.secrets:
            all_keys = [st.secrets["GEMINI_API_KEY"]]
            
    if isinstance(all_keys, str):
        all_keys = [all_keys]
    return all_keys[0] if all_keys else ""

def save_api_key(api_key: str) -> bool:
    import os
    try:
        import toml
        os.makedirs(".streamlit", exist_ok=True)
        secrets = {}
        if os.path.exists(".streamlit/secrets.toml"):
            try:
                with open(".streamlit/secrets.toml", "r", encoding="utf-8") as f:
                    secrets = toml.load(f)
            except Exception:
                secrets = {}
        
        keys = [k.strip() for k in str(api_key).split(",") if k.strip()]
        if len(keys) > 1:
            secrets["GEMINI_API_KEYS"] = keys
            if "GEMINI_API_KEY" in secrets:
                del secrets["GEMINI_API_KEY"]
        elif len(keys) == 1:
            secrets["GEMINI_API_KEY"] = keys[0]
            if "GEMINI_API_KEYS" in secrets:
                del secrets["GEMINI_API_KEYS"]
        else:
            if "GEMINI_API_KEY" in secrets:
                del secrets["GEMINI_API_KEY"]
            if "GEMINI_API_KEYS" in secrets:
                del secrets["GEMINI_API_KEYS"]
                
        with open(".streamlit/secrets.toml", "w", encoding="utf-8") as f:
            toml.dump(secrets, f)
        return True
    except Exception as e:
        from log_config import logger
        logger.error(f"Lỗi khi lưu API key vào file secrets.toml: {e}", exc_info=True)
        st.error(f"❌ Lỗi khi ghi file cấu hình API Key: {e}")
        return False

def render_chatbot():
    # 7. CHATBOT BONG BÓNG LƠ LỬNG
    # ==========================================
    lang = st.session_state.get('lang', 'vi')
    lbl_online = "Đang trực tuyến" if lang == 'vi' else "オンライン"
    lbl_api_warn = "⚠️ Vui lòng nhập Gemini API Key để Chatbot hoạt động!" if lang == 'vi' else "⚠️ チャットボットを機能させるにはGemini APIキーを入力してください！"
    lbl_api_tip = "💡 **Mẹo:** Bạn có thể nhập nhiều API Key cách nhau bằng dấu phẩy (,) để hệ thống tự động luân chuyển khi quá tải!" if lang == 'vi' else "💡 **ヒント:** 複数のAPIキーをカンマ(,)で区切って入力すると、過負荷時に自動で切り替わります！"
    lbl_api_placeholder = "Nhập Gemini API Key tại đây..." if lang == 'vi' else "ここにAPIキーを入力..."
    lbl_api_label = "Nhập Gemini API Key" if lang == 'vi' else "Gemini APIキーを入力"
    lbl_settings = "⚙️ Cài đặt Chatbot" if lang == 'vi' else "⚙️ チャットボット設定"
    lbl_change_api = "🔄 Đổi mã API Key" if lang == 'vi' else "🔄 APIキーを変更する"
    lbl_chat_input = "Hỏi AI..." if lang == 'vi' else "AIに質問する..."
    lbl_hello = "Xin chào! Tôi có thể giúp gì cho bạn về dữ liệu chấm công?" if lang == 'vi' else "こんにちは！勤怠データについて何かお手伝いしましょうか？"
    lang_instruction = "TIẾNG VIỆT" if lang == 'vi' else "TIẾNG NHẬT (日本語)"

    lbl_bubble_btn = "🌸"
    lbl_bubble_help = "Trợ lý ảo thông minh giải đáp dữ liệu chấm công 24/7" if lang == 'vi' else "勤怠データについて24時間365日回答するAIアシスタント"

    theme_mode = st.session_state.get('theme_mode', 'light')
    T = get_theme(theme_mode)
    cb_btn_bg = T['primary_gradient']
    cb_btn_bg_hover = T['accent_gradient']
    cb_btn_shadow = T['shadow']
    cb_btn_shadow_hover = T['shadow_glow']
    cb_header_bg = T['primary']
    cb_icon_bg = T['accent_gradient']

    hide_chatbot = st.session_state.get('hide_chatbot', False)
    display_css = "display: none !important;" if hide_chatbot else ""


    # Chatbot bám sát góc dưới phải màn hình và luôn đi theo khung nhìn khi cuộn trang
    pos_rules = """
    position: fixed !important;
    bottom: 90px !important;
    right: 24px !important;
    z-index: 9999999 !important;
    width: auto !important;
    height: auto !important;
    margin: 0 !important;
    padding: 0 !important;
    """

    css = f"""
    <style>
    .st-key-chat_bubble_wrap button[data-testid="stPopoverButton"] {{
        border-radius: 50% !important;
        width: 60px !important; height: 60px !important;
        background: transparent !important;
        color: white !important; border: none !important;
        box-shadow: none !important;
        filter: drop-shadow(0 8px 16px rgba(255, 117, 140, 0.4)) !important;
        font-size: 26px !important; 
        padding: 0 !important; 
        display: flex !important; align-items: center !important; justify-content: center !important;
        transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1) !important;
    }}
    .st-key-chat_bubble_wrap button[data-testid="stPopoverButton"]:hover {{
        transform: translateY(-3px) scale(1.1) !important;
        filter: drop-shadow(0 14px 24px rgba(255, 117, 140, 0.6)) !important;
        background: transparent !important;
    }}
    .st-key-chat_bubble_wrap button[data-testid="stPopoverButton"] p {{
        font-size: 45px !important; margin: 0 !important; position: relative !important; display: inline-block !important;
    }}
    @keyframes sakura_dance {{
        0%, 100% {{ transform: translateY(0px) rotate(0deg) scale(1); }}
        33% {{ transform: translateY(-5px) rotate(12deg) scale(1.05); }}
        66% {{ transform: translateY(2px) rotate(-10deg) scale(0.98); }}
    }}
    .st-key-chat_bubble_wrap button[data-testid="stPopoverButton"] p {{
        animation: sakura_dance 3.5s ease-in-out infinite !important;
        transform-origin: center center !important;
    }}
    /* Vô hiệu hóa animation và transform trên các container cha để position: fixed bám thẳng vào cửa sổ trình duyệt (viewport) */
    div[data-testid="element-container"]:has(.st-key-chat_bubble_wrap),
    div[data-testid="stVerticalBlock"]:has(.st-key-chat_bubble_wrap),
    div[data-testid="stVerticalBlockBorderWrapper"]:has(.st-key-chat_bubble_wrap),
    div[class*="stMainBlockContainer"]:has(.st-key-chat_bubble_wrap),
    [data-testid="stMain"]:has(.st-key-chat_bubble_wrap),
    [data-testid="stAppViewContainer"]:has(.st-key-chat_bubble_wrap),
    .stApp:has(.st-key-chat_bubble_wrap) {{
        animation: none !important;
        transform: none !important;
        filter: none !important;
        perspective: none !important;
        contain: none !important;
    }}
    div[data-testid="element-container"]:has(.st-key-chat_bubble_wrap),
    .st-key-chat_bubble_wrap {{
        {pos_rules}
        {display_css}
    }}
    div[data-testid="stPopoverBody"] {{
        width: min(340px, 88vw) !important;
        max-height: calc(100vh - 110px) !important;
        background: {T['bg_app']} !important;
        border-radius: 20px !important;
        box-shadow: 0 25px 60px rgba(0,0,0,0.25) !important;
        padding: 0 !important;
        border: 1px solid {T['border']} !important;
        overflow: hidden !important;
    }}


    div[data-testid="stPopoverBody"] [data-testid="stChatMessage"] {{
        background: {T['bg_content']} !important; border: 1px solid {T['border']} !important;
        border-radius: 16px 16px 16px 4px !important; padding: 12px !important;
        box-shadow: 0 2px 8px rgba(0,0,0,0.02) !important; margin: 12px !important; width: auto !important;
    }}
    div[data-testid="stPopoverBody"] [data-testid="stChatMessage"] [data-testid="stMarkdownContainer"] p {{
        font-size: 13.5px !important; color: {T['text_secondary']} !important; line-height: 1.5 !important;
    }}
    div[data-testid="stPopoverBody"] [data-testid="stChatInput"] {{
        background: white !important; padding: 12px 16px !important; border-top: 1px solid #E2E8F0 !important;
    }}
    div[data-testid="stPopoverBody"] [data-testid="stChatInput"] textarea {{
        background: #F1F5F9 !important; border-radius: 24px !important; border: 1px solid #E2E8F0 !important; font-size: 13px !important;
    }}
    </style>
    """
    st.markdown(css, unsafe_allow_html=True)

    bubble = st.container(key="chat_bubble_wrap")
    with bubble:
        with st.popover(lbl_bubble_btn, help=lbl_bubble_help):
            st.markdown(f"""
            <div style="background:{cb_header_bg}; color:white; padding:16px; display:flex; align-items:center; gap:12px;">
                <div style="width:40px; height:40px; background:{cb_icon_bg}; border-radius:50%; display:flex; align-items:center; justify-content:center; font-size:24px; box-shadow:0 4px 10px rgba(0,0,0,0.1); flex-shrink:0;">🌸</div>
                <div style="display:flex; flex-direction:column;">
                    <div style="font-weight:700; font-size:15px; line-height:1.2;">V-MOS Assistant</div>
                    <div style="font-size:12px; display:flex; align-items:center; gap:6px; opacity:0.9; margin-top:4px;">
                        <span style="width:8px; height:8px; background:#10B981; border-radius:50%; display:inline-block;"></span> {lbl_online}
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            saved_key = load_saved_api_key()
            if saved_key:
                st.session_state['gemini_configured'] = True
            else:
                st.session_state['gemini_configured'] = False
                
            if not st.session_state['gemini_configured']:
                st.warning("⚠️ Vui lòng cấu hình API Key trong mục Cài đặt Hệ Thống để sử dụng Chatbot!" if lang == 'vi' else "⚠️ チャットボットを使用するには設定でAPIキーを構成してください！")
            else:
                if 'chat_messages' not in st.session_state:
                    st.session_state['chat_messages'] = [{"role": "assistant", "content": lbl_hello}]
                elif len(st.session_state['chat_messages']) == 1 and st.session_state['chat_messages'][0]["role"] == "assistant":
                    st.session_state['chat_messages'][0]["content"] = lbl_hello

                chat_container = st.container(height=480, key="chat_msg_area")
                with chat_container:
                    for msg in st.session_state['chat_messages']:
                        with st.chat_message(msg["role"], avatar="🌸" if msg["role"] == "assistant" else "👤"):
                            st.markdown(msg["content"])

                prompt = st.chat_input(lbl_chat_input)
                if prompt:
                    st.session_state['chat_messages'].append({"role": "user", "content": prompt})
                    with chat_container:
                        with st.chat_message("user", avatar="👤"):
                            st.markdown(prompt)
                        with st.chat_message("assistant", avatar="🌸"):
                            with st.spinner("⏳ Đang kết nối AI Gemini để phân tích dữ liệu..."):
                                try:
                                    import requests
                                    saved_key = load_saved_api_key()
                                    if not saved_key:
                                        raise Exception("Không tìm thấy API Key. Vui lòng nhập lại API Key ở menu.")

                                    df_ctx = "Chưa có dữ liệu"
                                    if st.session_state.get('app_page') == 'mos':
                                        if 'df_mos_edited' in st.session_state and st.session_state.df_mos_edited is not None:
                                            df_ctx = "--- DỮ LIỆU TỔNG HỢP DỰ ÁN MOS ---\n" + st.session_state.df_mos_edited.to_csv(index=False)
                                        else:
                                            df_ctx = "Chưa có dữ liệu dự án MOS nào được xử lý."
                                    else:
                                        if 'df_filtered_for_chat' in st.session_state and st.session_state.df_filtered_for_chat is not None:
                                            dff = st.session_state.df_filtered_for_chat.copy()
                                            ten = st.session_state.mapping.get('ten_nv', dff.columns[1])
                                            dff['Số giờ làm thực tế'] = pd.to_numeric(dff['Số giờ làm thực tế'], errors='coerce').fillna(0)
                                            dff['Giờ OT'] = pd.to_numeric(dff['Giờ OT'], errors='coerce').fillna(0)
                                            dff['Giờ hành chính'] = pd.to_numeric(dff['Giờ hành chính'], errors='coerce').fillna(0)
    
                                            df_detail = pd.DataFrame({
                                                "Tên NV": dff[ten],
                                                "Ngày": dff["Ngày"],
                                                "Giờ làm": dff["Giờ hành chính"].round(1),
                                                "OT": dff["Giờ OT"].round(1),
                                                "Tổng": dff["Số giờ làm thực tế"].round(1),
                                                "Ghi chú": dff["Ghi chú"]
                                            })
    
                                            if len(df_detail) <= 500:
                                                df_ctx = f"--- DỮ LIỆU CHI TIẾT (Toàn bộ {len(df_detail)} dòng) ---\n" + df_detail.to_csv(index=False)
                                            else:
                                                agg = df_detail.groupby("Tên NV").agg({'Giờ làm': 'sum', 'OT': 'sum', 'Tổng': 'sum'}).reset_index().round(1)
                                                df_ctx = f"--- BẢNG TỔNG HỢP (TẤT CẢ NHÂN VIÊN) ---\n{agg.to_csv(index=False)}\n\n"
    
                                                # Nén dữ liệu cực đại để AI đọc được toàn bộ không bị mất dòng nào
                                                compressed_lines = []
                                                for name, group in df_detail.groupby("Tên NV"):
                                                    days = []
                                                    for _, row in group.iterrows():
                                                        d = str(row["Ngày"])[:5] # "15/05/2026" -> "15/05"
                                                        hc = row["Giờ làm"]
                                                        ot = row["OT"]
                                                        note = str(row["Ghi chú"]).strip()
                                                        if ot == 0 and note == "":
                                                            if hc == 8.0: days.append(d)
                                                            elif hc == 0.0: days.append(f"{d}(Nghỉ)")
                                                            else: days.append(f"{d}({hc}h)")
                                                        else:
                                                            days.append(f"{d}({hc}h|OT:{ot}|{note})")
                                                    compressed_lines.append(f"{name}: " + ", ".join(days))
    
                                                df_ctx += "--- CHI TIẾT LỊCH SỬ TỪNG NGÀY (DẠNG NÉN SIÊU TỐI ƯU) ---\n"
                                                df_ctx += "Quy ước: Nếu chỉ ghi ngày (VD: 15/05) nghĩa là làm đủ 8h không OT. Các ngày khác ghi rõ (Giờ làm|OT|Ghi chú).\n"
                                                df_ctx += "\n".join(compressed_lines)
                                        elif 'df_result' in st.session_state and st.session_state.df_result is not None:
                                            df_ctx = st.session_state.df_result.head(50).to_csv(index=False)
                                        else:
                                            df_ctx = "Chưa có dữ liệu nào được tải lên."

                                    sys_prompt = f"""Bạn là V-MOS AI, một siêu trợ lý trí tuệ nhân tạo toàn năng.
Bạn là trợ lý chính cho hệ thống "Quản lý Chấm công & Dự án Nội bộ MOS". Bạn am hiểu sâu sắc về cả việc tính toán giờ làm việc (chấm công) và tổng hợp giờ làm dự án MOS.
Bạn SẴN SÀNG TRẢ LỜI BẤT KỲ CÂU HỎI NÀO CỦA NGƯỜI DÙNG (từ phân tích số liệu bảng chấm công, báo cáo dự án MOS, cho đến viết email, làm toán, lập trình, tư vấn).
Hãy thoải mái trò chuyện và hỗ trợ người dùng như một AI thực thụ. Tuyệt đối không bao giờ từ chối trả lời vì lý do "tôi chỉ là trợ lý chấm công".
BẠN ĐANG GIAO TIẾP VỚI NGƯỜI DÙNG. NGÔN NGỮ GIAO TIẾP VÀ TRẢ LỜI BẮT BUỘC LÀ: {lang_instruction}. BẠN PHẢI TUYỆT ĐỐI TUÂN THỦ NGÔN NGỮ NÀY TRONG CÂU TRẢ LỜI CỦA MÌNH.

[DỮ LIỆU HIỆN TẠI MÀ NGƯỜI DÙNG ĐANG TƯƠNG TÁC]:
{df_ctx}

Luôn ưu tiên trả lời tự nhiên, thân thiện và chính xác."""

                                    try:
                                        import httpx
                                        import time
                                        
                                        # Lịch sử hội thoại gửi lên Gemini API (chuẩn hóa luồng user <-> model)
                                        gemini_history = []
                                        last_role = None
                                        for msg in st.session_state.get('chat_messages', [])[-10:]:
                                            text = str(msg.get("content", "")).strip()
                                            if not text:
                                                continue
                                            
                                            role = "user" if msg.get("role") == "user" else "model"
                                            
                                            # Bỏ qua các thông báo lỗi trước đó để tránh làm hỏng cấu trúc history
                                            if role == "model" and ("Lỗi hệ thống" in text or "Lỗi kết nối" in text or "Hệ thống AI đang bận" in text or "API Key" in text):
                                                continue
                                                
                                            if role == last_role:
                                                if gemini_history:
                                                    gemini_history[-1]["parts"][0]["text"] += "\n" + text
                                            else:
                                                gemini_history.append({"role": role, "parts": [{"text": text}]})
                                                last_role = role

                                        # Đảm bảo gemini_history bắt đầu và kết thúc chuẩn
                                        if not gemini_history or gemini_history[-1]["role"] != "user":
                                            gemini_history.append({"role": "user", "parts": [{"text": prompt}]})

                                        data = {
                                            "systemInstruction": {"parts": [{"text": sys_prompt}]},
                                            "contents": gemini_history,
                                            "generationConfig": {"temperature": 0.7}
                                        }
                                        
                                        # Lấy danh sách keys từ biến môi trường
                                        all_keys = []
                                        
                                        if "GEMINI_API_KEYS" in st.secrets:
                                            all_keys = st.secrets["GEMINI_API_KEYS"]
                                        elif "GEMINI_API_KEY" in st.secrets:
                                            all_keys = [st.secrets["GEMINI_API_KEY"]]
                                            
                                        if isinstance(all_keys, str):
                                            all_keys = [all_keys]
                                            
                                        if not all_keys:
                                            answer = "⚠️ Vui lòng cấu hình API Key."
                                        else:
                                            # Thử qua tất cả các key hoặc thử lại nhiều lần nếu chỉ có 1 key
                                            max_retries = max(3, len(all_keys) * 2)
                                            answer = None
                                            for attempt in range(max_retries):
                                                current_key = all_keys[attempt % len(all_keys)]
                                                url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent"
                                                
                                                res = httpx.post(
                                                    url, 
                                                    headers={"Content-Type": "application/json", "x-goog-api-key": current_key}, 
                                                    json=data, 
                                                    timeout=30.0
                                                )
                                                
                                                if res.status_code == 200:
                                                    res_data = res.json()
                                                    candidates = res_data.get("candidates", [])
                                                    if candidates:
                                                        cand = candidates[0]
                                                        content = cand.get("content", {})
                                                        parts = content.get("parts", [])
                                                        if parts and len(parts) > 0 and "text" in parts[0]:
                                                            answer = parts[0]["text"].strip()
                                                        else:
                                                            finish_reason = cand.get("finishReason", "UNKNOWN")
                                                            answer = f"Rất tiếc, AI không thể tạo phản hồi cho nội dung này (Trạng thái: {finish_reason}). Vui lòng hỏi lại với từ khóa khác!"
                                                    else:
                                                        answer = "Không nhận được phản hồi từ AI."
                                                    break
                                                elif res.status_code == 400 and "API_KEY_INVALID" in res.text:
                                                    if attempt == max_retries - 1:
                                                        answer = "Tất cả các API Key đều không hợp lệ. Vui lòng kiểm tra lại Cài đặt Chatbot."
                                                    continue
                                                elif res.status_code in [429, 503, 500, 502, 504]:
                                                    if attempt < max_retries - 1:
                                                        import time
                                                        time.sleep(1.0) # Sleep 1s trước khi thử lại
                                                        continue
                                                    else:
                                                        answer = f"Hệ thống AI đang bận (Lỗi {res.status_code}). Vui lòng đợi một lát rồi thử lại nhé!"
                                                        break
                                                else:
                                                    answer = f"Lỗi kết nối Gemini: {res.text}" 
                                                    break
                                                
                                        if not answer:
                                            answer = "Lỗi không xác định khi kết nối."
                                    except Exception as e:
                                        answer = f"Lỗi hệ thống: {str(e)}"

                                    st.markdown(answer)
                                    st.session_state['chat_messages'].append({"role": "assistant", "content": answer})
                                except Exception as e:
                                    st.error(f"Lỗi: {str(e)}")


_MOS_AI_CACHE = {}
_GEMINI_RATE_LIMITED_UNTIL = 0

def smart_offline_summarize(ma_da: str, ten_da: str, tasks_list: list, phan_vung: str = "") -> str:
    """
    Tóm tắt thông minh offline (không cần AI) khi có nhiều nội dung ủy thác khác nhau.
    Gộp thành 1 nội dung duy nhất chuẩn văn phong thương mại Nhật - Việt.
    """
    if not tasks_list:
        return ""
    if len(tasks_list) == 1:
        return str(tasks_list[0]).strip()
        
    # Thay vì đè bằng tên phân vùng chung chung, hãy giữ nguyên các nhiệm vụ thực tế và nối bằng dấu phẩy
    return ', '.join([str(t).strip() for t in tasks_list])

def summarize_tasks_with_ai(ma_da: str, ten_da: str, tasks: list) -> str:
    """
    Gọi Gemini API để tóm tắt danh sách task thành 1 câu nội dung chung.
    Có bộ nhớ đệm (cache) và tự động ngắt (fallback) khi vượt giới hạn RPM.
    """
    global _GEMINI_RATE_LIMITED_UNTIL
    import time
    if not tasks:
        return ''
    
    tasks_unique = list(dict.fromkeys([t.strip() for t in tasks if t.strip()]))
    if len(tasks_unique) == 1 and not ten_da:
        return tasks_unique[0]
        
    cache_key = (str(ma_da), str(ten_da), tuple(tasks_unique))
    if cache_key in _MOS_AI_CACHE:
        return _MOS_AI_CACHE[cache_key]
    
    fallback_res = smart_offline_summarize(ma_da, ten_da, tasks_unique)
    if time.time() < _GEMINI_RATE_LIMITED_UNTIL:
        return fallback_res

    prompt = f"""Dưới đây là danh sách các nội dung công việc của dự án 
"{ten_da}" (mã: {ma_da}) bằng tiếng Nhật:

{chr(10).join(f'- {t}' for t in tasks_unique)}

Nhiệm vụ của bạn:
1. Dịch tên dự án "{ten_da}" sang tiếng Việt.
2. Tóm tắt danh sách công việc trên thành 1 câu súc tích bằng tiếng Nhật, sau đó dịch câu đó sang tiếng Việt.

Trả về kết quả DƯỚI DẠNG JSON với định dạng sau (không giải thích thêm):
{{
  "ten_da_song_ngu": "Tên tiếng Nhật \\n Tên tiếng Việt",
  "noi_dung_song_ngu": "Tóm tắt tiếng Nhật \\n Tóm tắt tiếng Việt"
}}
"""

    try:
        api_key = load_saved_api_key()
        if not api_key:
            return fallback_res
            
        url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent"
        response = httpx.post(
            url,
            headers={"Content-Type": "application/json", "x-goog-api-key": api_key},
            json={
                "contents": [{"parts": [{"text": prompt}]}],
                "generationConfig": {"temperature": 0.2}
            },
            timeout=6.0
        )
        if response.status_code == 429:
            _GEMINI_RATE_LIMITED_UNTIL = time.time() + 30
            return fallback_res
        if response.status_code != 200:
            return fallback_res
            
        data = response.json()
        result_text = data["candidates"][0]["content"]["parts"][0]["text"].strip()
        
        import json
        if "```json" in result_text:
            result_text = result_text.split("```json")[1].split("```")[0]
        elif "```" in result_text:
            result_text = result_text.split("```")[1].split("```")[0]
            
        res_json = json.loads(result_text)
        _MOS_AI_CACHE[cache_key] = res_json
        return res_json
    except Exception:
        return fallback_res


def batch_summarize_projects(projects: list) -> dict:
    """
    Tóm tắt hàng loạt dự án trong 1 lần gọi API để không bị lỗi vượt giới hạn 15 RPM khi tải nhiều file.
    Trả về dict: { ma_da: { 'ten_da_song_ngu': ..., 'noi_dung_song_ngu': ... } }
    """
    global _GEMINI_RATE_LIMITED_UNTIL
    import time
    results = {}
    uncached = []
    
    for p in projects:
        ma_da = str(p.get('ma_da', '')).strip()
        ten_da = str(p.get('ten_da', '')).strip()
        tasks = p.get('tasks', [])
        tasks_unique = list(dict.fromkeys([t.strip() for t in tasks if t.strip()]))
        ckey = (ma_da, ten_da, tuple(tasks_unique))
        if ckey in _MOS_AI_CACHE:
            results[ma_da] = _MOS_AI_CACHE[ckey]
        else:
            uncached.append((ma_da, ten_da, tasks_unique, ckey))
            
    if not uncached:
        return results
        
    api_key = load_saved_api_key()
    if not api_key or time.time() < _GEMINI_RATE_LIMITED_UNTIL:
        for ma_da, ten_da, tasks_u, ckey in uncached:
            res = {'ten_da_song_ngu': ten_da, 'noi_dung_song_ngu': smart_offline_summarize(ma_da, ten_da, tasks_u)}
            results[ma_da] = res
            _MOS_AI_CACHE[ckey] = res
        return results

    # Chunk into batches of up to 8 projects để AI không bỏ sót
    batch_size = 8
    for i in range(0, len(uncached), batch_size):
        chunk = uncached[i:i+batch_size]
        prompt_items = []
        for ma_da, ten_da, tasks_u, _ in chunk:
            ts_str = '; '.join(tasks_u)
            prompt_items.append(f'- Mã: {ma_da} | Tên: {ten_da} | Tasks: {ts_str}')
            
        prompt = f"""Dưới đây là danh sách {len(chunk)} dự án (Mã, Tên tiếng Nhật, Tasks tiếng Nhật):
{chr(10).join(prompt_items)}

Nhiệm vụ: Dịch tên dự án sang tiếng Việt và tóm tắt tasks thành 1 câu tiếng Nhật kèm bản dịch tiếng Việt.
YÊU CẦU BẮT BUỘC: Bạn PHẢI trả về bản dịch cho ĐẦY ĐỦ {len(chunk)} dự án trên, KHÔNG ĐƯỢC BỎ SÓT BẤT KỲ DỰ ÁN NÀO.
Trả về duy nhất JSON dictionary theo định dạng sau (với key là Mã dự án):
{{
  "MÃ_DỰ_ÁN_1": {{
    "ten_da_song_ngu": "Tên Nhật \\n Tên Việt",
    "noi_dung_song_ngu": "Tóm tắt Nhật \\n Tóm tắt Việt"
  }},
  "MÃ_DỰ_ÁN_2": {{ ... }}
}}"""
        try:
            url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent"
            response = httpx.post(
                url,
                headers={"Content-Type": "application/json", "x-goog-api-key": api_key},
                json={
                    "contents": [{"parts": [{"text": prompt}]}],
                    "generationConfig": {"temperature": 0.2}
                },
                timeout=12.0
            )
            if response.status_code == 429:
                _GEMINI_RATE_LIMITED_UNTIL = time.time() + 30
                raise Exception("Rate limited")
            if response.status_code == 200:
                data = response.json()
                result_text = data["candidates"][0]["content"]["parts"][0]["text"].strip()
                import json
                if "```json" in result_text:
                    result_text = result_text.split("```json")[1].split("```")[0]
                elif "```" in result_text:
                    result_text = result_text.split("```")[1].split("```")[0]
                parsed = json.loads(result_text)
                for ma_da, ten_da, tasks_u, ckey in chunk:
                    if ma_da in parsed and isinstance(parsed[ma_da], dict):
                        res = parsed[ma_da]
                    else:
                        res = {'ten_da_song_ngu': ten_da, 'noi_dung_song_ngu': smart_offline_summarize(ma_da, ten_da, tasks_u)}
                    results[ma_da] = res
                    _MOS_AI_CACHE[ckey] = res
            else:
                raise Exception(f"HTTP {response.status_code}")
        except Exception:
            for ma_da, ten_da, tasks_u, ckey in chunk:
                res = {'ten_da_song_ngu': ten_da, 'noi_dung_song_ngu': smart_offline_summarize(ma_da, ten_da, tasks_u)}
                results[ma_da] = res
                _MOS_AI_CACHE[ckey] = res

    return results


