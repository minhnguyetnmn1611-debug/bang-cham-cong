import base64
import os
import streamlit as st


def get_splash_img_base64():
    """Lấy dữ liệu base64 của ảnh mở đầu V.MOS từ thư mục assets một cách an toàn trên mọi hệ điều hành (Windows/Linux/GitHub)."""
    app_dir = os.path.dirname(os.path.abspath(__file__))
    possible_paths = [
        os.path.join(app_dir, "assets", "splash_japan_trip.png"),
        os.path.join(app_dir, "assets", "splash_japan_trip.PNG"),
        os.path.join(os.getcwd(), "assets", "splash_japan_trip.png"),
        os.path.join(os.getcwd(), "assets", "splash_japan_trip.PNG"),
    ]
    
    # Tìm kiếm linh hoạt không phân biệt chữ hoa/thường trên Linux/GitHub Server
    assets_dir = os.path.join(app_dir, "assets")
    if os.path.exists(assets_dir):
        for fname in os.listdir(assets_dir):
            if fname.lower() == "splash_japan_trip.png" or "splash" in fname.lower():
                possible_paths.insert(0, os.path.join(assets_dir, fname))

    for p in possible_paths:
        if os.path.exists(p):
            try:
                with open(p, "rb") as f:
                    return base64.b64encode(f.read()).decode("utf-8")
            except Exception:
                continue
    return ""


def render_splash_screen():
    """
    Màn hình chào mừng V.MOS áp dụng CHÍNH XÁC 100% đoạn mã CSS do người dùng chỉ định:
    .hero-image-wrapper {
      width: 100%;
      height: 100vh;
      overflow: hidden;
    }
    .hero-image {
      width: 100%;
      height: 100%;
      object-fit: cover;
      object-position: center right;
      display: block;
    }
    """
    is_vi = st.session_state.get('lang', 'vi') == 'vi'
    img_b64 = get_splash_img_base64()

    # ─── Ẩn toàn bộ khung mặc định của Streamlit & áp dụng CSS ───
    st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800;900&display=swap');

        [data-testid="stSidebar"], [data-testid="stHeader"], #MainMenu, footer {
            display: none !important;
        }
        .block-container {
            padding: 0 !important;
            max-width: 100% !important;
        }
        html, body, .stApp, .block-container {
            padding: 0 !important;
            margin: 0 !important;
            width: 100vw !important;
            height: 100vh !important;
            max-width: 100vw !important;
            max-height: 100vh !important;
            overflow: hidden !important;
            background: #FFFFFF !important;
        }

        /* Container Hero chính */
        .hero {
            position: fixed !important;
            top: 0 !important;
            left: 0 !important;
            width: 100vw !important;
            height: 100vh !important;
            padding: 0 !important;
            margin: 0 !important;
            z-index: 999999 !important;
            background-color: #FFFFFF !important;
        }

        /* Wrapper chứa bức ảnh theo đúng CSS người dùng chỉ định */
        .hero-image-wrapper {
            width: 100%;
            height: 100vh;
            overflow: hidden;
            position: relative;
        }

        /* Bức ảnh sắc nét cao được tối ưu hiển thị HD */
        .hero-image {
            width: 100%;
            height: 100%;
            object-fit: cover;
            object-position: center right;
            display: block;
            image-rendering: -webkit-optimize-contrast !important;
            image-rendering: crisp-edges !important;
            filter: contrast(107%) saturate(106%) brightness(101%) !important;
            transform: translateZ(0) !important;
            backface-visibility: hidden !important;
        }

        /* Lớp nội dung chữ phủ rộng rãi trên mảng trắng bên trái */
        .vmos-text-box {
            position: absolute;
            top: 24%;
            left: 8.5%;
            width: 44%;
            z-index: 50;
            line-height: normal;
            font-family: 'Plus Jakarta Sans', sans-serif;
        }

        .vmos-overlay-logo {
            font-size: 28px;
            font-weight: 900;
            color: #0F172A;
            letter-spacing: -0.5px;
            margin-bottom: 44px;
        }

        .vmos-overlay-eyebrow {
            font-size: 15px;
            font-weight: 800;
            letter-spacing: 3px;
            color: #1E293B;
            text-transform: uppercase;
            margin-bottom: 8px;
        }

        .vmos-overlay-title {
            font-size: 78px;
            font-weight: 900;
            line-height: 0.95;
            color: #DB2777;
            background: linear-gradient(135deg, #DB2777 0%, #E91E8C 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            letter-spacing: -1.5px;
            margin-bottom: 20px;
        }

        .vmos-overlay-desc {
            font-size: 16px;
            line-height: 1.65;
            color: #475569;
            margin-bottom: 34px;
            max-width: 420px;
        }

        /* Đưa nút Bắt đầu nổi trực tiếp lên trên màn hình Hero */
        .st-key-splash_start_btn {
            position: fixed !important;
            left: 8.5% !important;
            top: 52% !important;
            z-index: 10000000 !important;
            width: auto !important;
        }

        .st-key-splash_start_btn button {
            background: linear-gradient(135deg, #F472B6 0%, #EC4899 100%) !important;
            color: #FFFFFF !important;
            border: none !important;
            border-radius: 999px !important;
            padding: 15px 44px !important;
            font-weight: 800 !important;
            font-size: 15px !important;
            letter-spacing: 1px !important;
            text-transform: uppercase !important;
            box-shadow: 0 10px 24px rgba(236, 72, 153, 0.45) !important;
            transition: all 0.25s ease !important;
            cursor: pointer !important;
        }
        .st-key-splash_start_btn button:hover {
            transform: translateY(-2px) scale(1.02) !important;
            box-shadow: 0 14px 32px rgba(236, 72, 153, 0.6) !important;
        }

        @media (max-width: 768px) {
            .vmos-text-box {
                width: 85%;
                left: 6%;
                top: 5%;
            }
            .vmos-overlay-title { font-size: 48px; }
            .st-key-splash_start_btn {
                left: 6% !important;
                top: 60% !important;
                width: 88vw !important;
            }
        }
        </style>
    """, unsafe_allow_html=True)

    btn_label = "GET STARTED" if not is_vi else "🚀 Bắt đầu / GET STARTED"
    desc_content = "Hệ thống quản trị nhân sự, chấm công và kê khai công số dự án MOS — kết nối Việt Nam & Nhật Bản." if is_vi else "Discover the serenity, culture, and beauty of Japan through curated experiences."

    if img_b64:
        st.markdown(f"""
        <div class="hero">
          <div class="hero-image-wrapper">
            <img src="data:image/png;base64,{img_b64}" class="hero-image" alt="" />
            <div class="vmos-text-box">
              <div class="vmos-overlay-eyebrow">WELCOME TO</div>
              <div class="vmos-overlay-title">V. MOS</div>
              <div class="vmos-overlay-desc">{desc_content}</div>
            </div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        if st.button(btn_label, key="splash_start_btn"):
            st.session_state['splash_done'] = True
            st.session_state['app_page'] = 'overview'
            try:
                st.query_params["page"] = "overview"
            except Exception:
                pass
            st.rerun()
    else:
        # Fallback
        st.markdown("""
        <div class="hero">
          <div style="background:#FFF; padding:60px; border-radius:28px; text-align:center; max-width:600px;">
            <h1 style="color:#0F172A; font-size:46px; margin-bottom:16px;">V.MOS</h1>
            <p style="color:#475569; font-size:16px; margin-bottom:30px;">Discover the serenity, culture, and beauty of Japan through curated experiences.</p>
          </div>
        </div>
        """, unsafe_allow_html=True)

        if st.button(btn_label, key="splash_start_btn"):
            st.session_state['splash_done'] = True
            st.session_state['app_page'] = 'overview'
            try:
                st.query_params["page"] = "overview"
            except Exception:
                pass
            st.rerun()
