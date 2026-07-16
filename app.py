# -*- coding: utf-8 -*-
# Streamlit App Entry Point
import streamlit as st # Force reload 2
from log_config import logger
import streamlit.components.v1 as components
import pandas as pd
import datetime
import calendar
import io
import os
import sys
import importlib

# @st.cache_resource removed temporarily to force reload
def force_reload_once():
    for m in list(sys.modules.keys()):
        if m.startswith('page_') or m in ['theme', 'translations']:
            try:
                importlib.reload(sys.modules[m])
            except:
                pass
    return True

force_reload_once()

import base64
import math
from PIL import Image as PILImage

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from copy import copy
import sqlite3
import toml
from translations import get_t, get_data_t, translate_name, translate_dia_diem
from utils import *
from excel_export import *
from db import init_db, get_company_emp_options, get_company_emp_dict, save_field_checkin, get_field_checkins, save_to_db
from page_mos import render_mos_page
from page_checkin import render_checkin_page
from page_gamification import render_gamification_dashboard
from theme import get_theme
from email_service import send_email_notification

APP_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(APP_DIR, "assets")
FAVICON_PATH = os.path.join(ASSETS_DIR, "favicon.png")

def load_favicon():
    try:
        return PILImage.open(FAVICON_PATH)
    except Exception:
        return "📊"

st.set_page_config(page_title="V.MOS SYSTEM", page_icon=load_favicon(), layout="wide", initial_sidebar_state="expanded")
global_css_container = st.empty()

import sys
if 'ai_chat' in sys.modules:
    del sys.modules['ai_chat']
import ai_chat

if st.session_state.get('theme_just_toggled', False):
    st.session_state.theme_just_toggled = False

if st.session_state.get('save_success', False):
    st.toast("✅ Lưu thành công!")
    st.session_state.save_success = False

components.html("""
<script>
(function() {
    try {
        const parentDoc = window.parent.document;
        // --- TỰ ĐỘNG CĂN GIỮA VÀ LÀM MỜ NỀN MỌI DIALOG BẰNG CSS ---
        if (!parentDoc.getElementById('vmos-dialog-style')) {
            const ds = parentDoc.createElement('style');
            ds.id = 'vmos-dialog-style';
            ds.textContent = `
                div[data-testid="stModal"] {
                    background-color: rgba(15, 23, 42, 0.65) !important;
                    backdrop-filter: blur(8px) !important;
                    -webkit-backdrop-filter: blur(8px) !important;
                }
                div[role="dialog"] {
                    margin-top: 15vh !important;
                    margin-bottom: auto !important;
                    border-radius: 16px !important;
                    box-shadow: 0 20px 60px rgba(0,0,0,0.5) !important;
                    border: 1px solid rgba(255,255,255,0.2) !important;
                }
            `;
            parentDoc.head.appendChild(ds);
        }
    } catch (e) {}
    
    // ── 6. SMOOTH PAGE TRANSITION – fade-out khi click nav, fade-in sau rerun ──
    (function setupPageTransition() {
        const parentDoc = window.frameElement
            ? window.frameElement.ownerDocument
            : document;

        // Inject transition CSS vào parent document
        if (!parentDoc.getElementById('vmos-transition-style')) {
            const s = parentDoc.createElement('style');
            s.id = 'vmos-transition-style';
            s.textContent = `
                [data-testid="stMain"],
                [data-testid="stSidebarContent"] {
                    transition: opacity 0.15s ease, transform 0.15s ease !important;
                }
                .vmos-page-leaving [data-testid="stMain"] {
                    opacity: 0 !important;
                    transform: translateY(-4px) !important;
                    pointer-events: none !important;
                }
            `;
            parentDoc.head.appendChild(s);
        }

        // Lắng nghe click trên các nút chuyển trang (bỏ qua nút ngầm ORG_BTN_)
        const bindNavClicks = () => {
            const sidebar = parentDoc.querySelector('[data-testid="stSidebar"]');
            if (!sidebar) return;
            sidebar.querySelectorAll('[data-testid="stButton"] > button').forEach(btn => {
                if (btn.__vmos_nav_bound) return;
                if (btn.innerText && btn.innerText.includes('ORG_BTN_')) return;
                btn.__vmos_nav_bound = true;
                btn.addEventListener('click', () => {
                    parentDoc.body.classList.add('vmos-page-leaving');
                    // Tự động xóa class sau tối đa 400ms để tuyệt đối không bị kẹt ẩn giao diện
                    setTimeout(() => {
                        parentDoc.body.classList.remove('vmos-page-leaving');
                    }, 400);
                    const observer = new MutationObserver(() => {
                        parentDoc.body.classList.remove('vmos-page-leaving');
                        observer.disconnect();
                    });
                    observer.observe(
                        parentDoc.querySelector('[data-testid="stMain"]') || parentDoc.body,
                        { childList: true, subtree: true }
                    );
                });
            });
        };
        
        const bindLangClick = () => {
            const langBtn = parentDoc.querySelector('.st-key-lang_switch_btn button');
            if (!langBtn || langBtn.__vmos_lang_bound) return;
            langBtn.__vmos_lang_bound = true;
            langBtn.addEventListener('click', () => {
                let overlay = parentDoc.getElementById('vmos-lang-overlay');
                if (!overlay) {
                    overlay = parentDoc.createElement('div');
                    overlay.id = 'vmos-lang-overlay';
                    overlay.innerHTML = `
                        <div style="display:flex; flex-direction:column; align-items:center; justify-content:center; height:100%;">
                            <div style="width:40px; height:40px; border:4px solid rgba(14,165,233,0.3); border-top-color:#0ea5e9; border-radius:50%; animation:vmos-spin 0.8s linear infinite;"></div>
                            <div style="margin-top:16px; font-weight:600; color:#334155; font-size:15px; font-family:sans-serif; text-shadow:0 1px 2px rgba(255,255,255,0.8);">Đang đổi ngôn ngữ / 言語を変更中...</div>
                        </div>
                        <style>
                            @keyframes vmos-spin { to { transform: rotate(360deg); } }
                            #vmos-lang-overlay {
                                position: fixed; top: 0; left: 0; width: 100vw; height: 100vh;
                                background: rgba(255,255,255,0.7); backdrop-filter: blur(12px); -webkit-backdrop-filter: blur(12px);
                                z-index: 9999999; display: flex; align-items: center; justify-content: center;
                                opacity: 1; transition: opacity 0.3s ease;
                            }
                        </style>
                    `;
                    parentDoc.body.appendChild(overlay);
                } else {
                    overlay.style.opacity = '1';
                    overlay.style.display = 'flex';
                }
                
                const observer = new MutationObserver(() => {
                    overlay.style.opacity = '0';
                    setTimeout(() => { overlay.style.display = 'none'; }, 300);
                    observer.disconnect();
                });
                observer.observe(parentDoc.querySelector('[data-testid="stMain"]') || parentDoc.body, { childList: true, subtree: true });
                
                setTimeout(() => {
                    overlay.style.opacity = '0';
                    setTimeout(() => { overlay.style.display = 'none'; }, 300);
                    observer.disconnect();
                }, 2000);
            });
        };

        // Re-bind sau mỗi lần Streamlit rerender
        const sidebarObserver = new MutationObserver(bindNavClicks);
        const waitForSidebar = setInterval(() => {
            const sb = parentDoc.querySelector('[data-testid="stSidebar"]');
            if (sb) {
                clearInterval(waitForSidebar);
                sidebarObserver.observe(sb, { childList: true, subtree: true });
                bindNavClicks();
            }
        }, 200);
        
        const globalObserver = new MutationObserver(bindLangClick);
        globalObserver.observe(parentDoc.body, { childList: true, subtree: true });
        bindLangClick();
    })();
})();
</script>
""", height=0, width=0)




# ==========================================
# GLOBAL STYLES & BACKGROUND
# ==========================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=Inter:wght@400;500;600;700&display=swap');

/* ===================================================================
   PAGE TRANSITION – Fade-in mượt mà khi chuyển trang
   =================================================================== */

/* 1. Phần nội dung chính fade-in từ nhẹ */
[data-testid="stMain"] {
    animation: vmos-fadein 0.28s ease forwards;
}
[data-testid="stSidebarContent"] {
    animation: vmos-fadein 0.22s ease forwards;
}

@keyframes vmos-fadein {
    from { opacity: 0; transform: translateY(6px); }
    to   { opacity: 1; transform: translateY(0);   }
}

/* 2. Mỗi block-container element xuất hiện tuần tự (stagger) */
[data-testid="stVerticalBlock"] > [data-testid="stVerticalBlockBorderWrapper"],
[data-testid="stVerticalBlock"] > div[data-testid="element-container"],
[data-testid="stVerticalBlock"] > [data-testid="stHorizontalBlock"] {
    animation: vmos-fadein 0.3s ease both;
}

/* 3. Ngăn flash trắng giữa các lần rerun — giữ màu nền nhất quán */
html, body, .stApp, [data-testid="stAppViewContainer"] {
    background-color: #F0F9FF !important; /* fallback = brand-50 */
}

/* 4. Transition mượt khi hover/active các nút sidebar */
[data-testid="stSidebar"] [data-testid="stButton"] > button {
    transition: background 0.18s ease, transform 0.15s ease, box-shadow 0.18s ease !important;
}

/* 5. Smooth scroll toàn trang */
html {
    scroll-behavior: smooth !important;
}

/* 6. Prevent layout-shift flash trên nội dung có chiều cao cố định */
[data-testid="stSpinner"] {
    display: none !important;
}


/* Đồng bộ phông chữ hiện đại toàn bộ hệ thống */
html, body, [class*="css"], h1, h2, h3, h4, h5, h6, p, div, span, button, input, select, textarea, a, li, ul, ol, table, td, th {
    font-family: 'Plus Jakarta Sans', 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
}
h1, h2, h3, h4, .chamcong-upload-title, .welcome-title {
    font-family: 'Plus Jakarta Sans', 'Inter', sans-serif !important;
}

/* Ensure all buttons center their text perfectly & modern styling */
div[data-testid="stButton"] > button {
    display: flex !important;
    align-items: center !important;
    justify-content: center !important;
    text-align: center !important;
    font-family: 'Plus Jakarta Sans', 'Inter', sans-serif !important;
    border-radius: 12px !important;
    font-weight: 700 !important;
    transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1) !important;
}
div[data-testid="stButton"] > button[kind="primary"] {
    background: var(--brand-gradient) !important;
    border: 1px solid rgba(255, 255, 255, 0.2) !important;
    box-shadow: var(--shadow) !important;
}
div[data-testid="stButton"] > button[kind="primary"]:hover {
    transform: translateY(-2px) !important;
    box-shadow: var(--shadow-glow) !important;
    background: var(--accent-gradient) !important;
}
div[data-testid="stButton"] > button[kind="primary"] p {
    color: #FFFFFF !important;
    text-shadow: 0 1px 2px rgba(0,0,0,0.2) !important;
}
div[data-testid="stButton"] > button[kind="secondary"] {
    background: var(--bg-card) !important;
    border: 1.5px solid var(--border) !important;
    box-shadow: 0 2px 6px rgba(0, 0, 0, 0.04) !important;
}
div[data-testid="stButton"] > button[kind="secondary"]:hover {
    border-color: var(--brand-primary) !important;
    background: var(--bg-card-hover) !important;
    transform: translateY(-1px) !important;
    box-shadow: 0 4px 12px rgba(59, 130, 246, 0.15) !important;
}
.stButton > button[kind="secondary"] p {
    color: var(--text-primary) !important;
    text-shadow: none !important;
}
.stButton > button p {
    text-align: center !important;
    margin: 0 !important;
    font-weight: 700 !important;
    font-family: 'Plus Jakarta Sans', 'Inter', sans-serif !important;
}

/* Chuẩn hóa giao diện các nút bấm/hộp thoại dạng chọn (Selectbox, Multiselect, Date Input, Text Input) theo chuẩn SaaS Card */
div[data-baseweb="select"] > div,
div[data-baseweb="input"] > div,
div[data-testid="stSelectbox"] div[data-baseweb="select"] > div,
div[data-testid="stMultiSelect"] div[data-baseweb="select"] > div,
div[data-testid="stDateInput"] div[data-baseweb="input"] > div,
div[data-testid="stTextInput"] div[data-baseweb="input"] > div,
div[data-testid="stNumberInput"] div[data-baseweb="input"] > div {
    background-color: var(--bg-card) !important;
    border: 1.5px solid var(--border) !important;
    border-radius: 12px !important;
    padding: 2px 6px !important;
    min-height: 40px !important;
    box-shadow: 0 1px 3px rgba(0, 0, 0, 0.03) !important;
    transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1) !important;
}

/* Khi hover vào hộp thoại chọn */
div[data-baseweb="select"] > div:hover,
div[data-baseweb="input"] > div:hover,
div[data-testid="stSelectbox"] div[data-baseweb="select"] > div:hover,
div[data-testid="stMultiSelect"] div[data-baseweb="select"] > div:hover,
div[data-testid="stDateInput"] div[data-baseweb="input"] > div:hover,
div[data-testid="stTextInput"] div[data-baseweb="input"] > div:hover,
div[data-testid="stNumberInput"] div[data-baseweb="input"] > div:hover {
    border-color: var(--brand-primary) !important;
    background-color: var(--bg-card-hover) !important;
    box-shadow: 0 4px 12px rgba(59, 130, 246, 0.12) !important;
}

/* Khi focus / click vào hộp thoại chọn */
div[data-baseweb="select"] > div:focus-within,
div[data-baseweb="input"] > div:focus-within,
div[data-testid="stSelectbox"] div[data-baseweb="select"] > div:focus-within,
div[data-testid="stMultiSelect"] div[data-baseweb="select"] > div:focus-within {
    border-color: var(--brand-primary) !important;
    box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.18) !important;
    background-color: var(--bg-card) !important;
}

/* Chuẩn hóa màu chữ và phông chữ bên trong hộp chọn */
div[data-baseweb="select"] span,
div[data-baseweb="select"] div,
div[data-baseweb="input"] input,
div[data-baseweb="base-input"] input {
    color: var(--text-primary) !important;
    font-weight: 600 !important;
    font-size: 14px !important;
    font-family: 'Plus Jakarta Sans', 'Inter', sans-serif !important;
}

/* Màu mũi tên dropdown (Chevron) */
div[data-baseweb="select"] svg {
    fill: var(--text-secondary) !important;
    transition: transform 0.2s ease !important;
}
div[data-baseweb="select"] > div:hover svg {
    fill: var(--brand-primary) !important;
}

/* ĐỒNG BỘ TỶ LỆ BO GÓC 12PX CHO TOÀN BỘ GIAO DIỆN VÀ NÚT BẤM TOÀN HỆ THỐNG */
.stButton > button,
button,
input,
select,
textarea,
div[data-baseweb="select"] > div,
div[data-baseweb="input"] > div,
div[data-baseweb="base-input"],
div[data-testid="stSelectbox"] div[data-baseweb="select"] > div,
div[data-testid="stMultiSelect"] div[data-baseweb="select"] > div,
div[data-testid="stDateInput"] div[data-baseweb="input"] > div,
div[data-testid="stTextInput"] div[data-baseweb="input"] > div,
div[data-testid="stNumberInput"] div[data-baseweb="input"] > div,
div[data-testid="stExpander"],
div[data-testid="stExpander"] details,
div[data-testid="stExpander"] summary,
div[data-testid="stPopoverBody"],
div[data-testid="stPopoverButton"],
div[data-testid="stMetric"],
div[data-testid="stDataFrame"],
div[data-testid="stTable"],
div[data-testid="stAlert"],
div[data-testid="stNotification"],
div[data-testid="stFileUploader"],
div[data-testid="stFileUploaderDropzone"],
div[data-testid="stVerticalBlockBorderWrapper"],
section[data-testid="stSidebar"] div[data-testid="stButton"] > button,
.card,
.upload-hint,
[class*="st-key-"] [data-testid="stExpander"] {
    border-radius: 12px !important;
}

/* Chuẩn hóa tỷ lệ giao diện gọn gàng, chống phì đại (oversized) trên các máy tính/màn hình có DPI/Scaling cao (125%, 150%) */
html, body, [class*="css"], .stMarkdown, .stDataFrame, div[data-testid="stText"] {
    font-size: clamp(12px, 0.75vw, 14.5px) !important;
}
section[data-testid="stSidebar"],
section[data-testid="stSidebar"] > div,
[data-testid="stSidebarContent"],
.main .block-container {
    overflow-x: hidden !important;
}
.main .block-container {
    max-width: 1650px !important; /* Giới hạn chiều rộng tối đa để giao diện không bị bè to trên màn hình lớn/máy mới */
    margin: 0 auto !important;
    padding-top: 1.5rem !important;
    padding-left: clamp(1rem, 2vw, 2.5rem) !important;
    padding-right: clamp(1rem, 2vw, 2.5rem) !important;
}

/* Thiết kế Responsive chuẩn (Gập cột trên điện thoại thay vì zoom) */
@media screen and (max-width: 768px) {
    div[data-testid="stHorizontalBlock"] {
        flex-wrap: wrap !important;
    }
    div[data-testid="column"] {
        min-width: 100% !important;
        margin-bottom: 0.8rem !important;
    }
    .main .block-container {
        padding: 1rem !important;
    }
}

/* ===== Mobile Nav Bar – thu gọn 5 nút top fixed < 640px ===== */
@media screen and (max-width: 640px) {
    /* Ẩn chữ, chỉ giữ emoji của 4 nav buttons */
    .st-key-nav_btn_notif div[data-testid="stButton"] > button p,
    .st-key-nav_btn_docs div[data-testid="stButton"] > button p,
    .st-key-nav_btn_support div[data-testid="stButton"] > button p,
    .st-key-nav_btn_profile div[data-testid="stButton"] > button p {
        font-size: 18px !important;
        letter-spacing: -0.5px !important;
    }
    /* Co lại width và vị trí right cho mobile */
    .st-key-btn_top_eyecare_fixed { right: 230px !important; width: 80px !important; }
    .st-key-nav_btn_notif         { right: 178px !important; width: 48px !important; }
    .st-key-nav_btn_docs          { right: 130px !important; width: 44px !important; }
    .st-key-nav_btn_support       { right: 82px !important;  width: 44px !important; }
    .st-key-nav_btn_profile       { right: 30px !important;  width: 48px !important; }
}

/* ===== Mobile cực nhỏ < 480px – ẩn hẳn eyecare + docs + support, giữ notif + profile ===== */
@media screen and (max-width: 480px) {
    .st-key-btn_top_eyecare_fixed { display: none !important; }
    .st-key-nav_btn_docs          { display: none !important; }
    .st-key-nav_btn_support       { display: none !important; }
    .st-key-nav_btn_notif         { right: 64px !important; width: 44px !important; }
    .st-key-nav_btn_profile       { right: 12px !important; width: 44px !important; }
}

/* ===== Gamification cards – giảm min-width trên mobile ===== */
@media screen and (max-width: 640px) {
    .vmos-badge-card {
        min-width: 140px !important;
        padding: 12px 10px !important;
    }
    .vmos-badge-card .badge-title {
        font-size: 13px !important;
    }
    .vmos-leaderboard-row {
        flex-wrap: wrap !important;
        gap: 6px !important;
    }
    /* Org chart tree cuộn ngang thay vì vỡ */
    .org-tree {
        overflow-x: auto !important;
        -webkit-overflow-scrolling: touch !important;
    }
    /* App header banner gọn trên mobile */
    .app-header {
        padding: 20px 18px !important;
        border-radius: 16px !important;
        flex-wrap: wrap !important;
        gap: 16px !important;
    }
}

/* Mở rộng thanh chức năng bên trái linh hoạt (responsive) */
section[data-testid="stSidebar"] {
    background: rgba(255, 255, 255, 0.9) !important;
    backdrop-filter: blur(20px) !important;
    -webkit-backdrop-filter: blur(20px) !important;
    border-right: 1.5px solid rgba(14, 165, 233, 0.35) !important;
}

/* Glassmorphism Cards for Expanders (Khắc phục chìm nền) */
div[data-testid="stExpander"] {
    background: rgba(255, 255, 255, 0.95) !important;
    backdrop-filter: blur(20px) !important;
    -webkit-backdrop-filter: blur(20px) !important;
    border: 1.5px solid rgba(14, 165, 233, 0.45) !important;
    border-radius: 18px !important;
    box-shadow: 0 10px 30px rgba(15, 23, 42, 0.15) !important;
    overflow: hidden !important;
    margin-bottom: 18px !important;
}
div[data-testid="stExpander"]:hover {
    box-shadow: 0 16px 40px rgba(14, 165, 233, 0.22) !important;
    border-color: rgba(14, 165, 233, 0.6) !important;
}

/* Expander Headers - High Contrast & Sleek Icons */
div[data-testid="stExpander"] details > summary {
    padding: 16px 22px !important;
    background: linear-gradient(90deg, #F8FAFC 0%, #EFF6FF 100%) !important;
    border-bottom: 1px solid rgba(203, 213, 225, 0.9) !important;
    font-weight: 800 !important;
    color: #0F172A !important;
    font-size: 16.5px !important;
}
div[data-testid="stExpander"] details[open] > summary {
    background: linear-gradient(135deg, #0F172A 0%, #1E3A8A 100%) !important;
    color: #FFFFFF !important;
    border-bottom: 2.5px solid #0EA5E9 !important;
}
div[data-testid="stExpander"] details[open] > summary p,
div[data-testid="stExpander"] details[open] > summary span {
    color: #FFFFFF !important;
}
div[data-testid="stExpander"] details > summary svg {
    color: #0284C7 !important;
    fill: #0284C7 !important;
}
div[data-testid="stExpander"] details[open] > summary svg {
    color: #0EA5E9 !important;
    fill: #0EA5E9 !important;
}

/* Tabs Styling - Modern Pill Design */
div[data-testid="stTabs"] button[data-baseweb="tab"] {
    background: rgba(255, 255, 255, 0.92) !important;
    backdrop-filter: blur(14px) !important;
    border-radius: 12px !important;
    padding: 12px 26px !important;
    margin-right: 12px !important;
    border: 1.5px solid #CBD5E1 !important;
    font-weight: 800 !important;
    color: #1E293B !important;
    box-shadow: 0 4px 12px rgba(0,0,0,0.08) !important;
    transition: all 0.25s ease !important;
}
div[data-testid="stTabs"] button[data-baseweb="tab"]:not([aria-selected="true"]):hover {
    border-color: #0EA5E9 !important;
    box-shadow: 0 16px 40px rgba(14, 165, 233, 0.22) !important;
    transform: translateY(-2px);
    color: #0EA5E9 !important;
}
div[data-testid="stTabs"] button[data-baseweb="tab"][aria-selected="true"] {
    background: linear-gradient(135deg, #0EA5E9 0%, #0284C7 100%) !important;
    color: #FFFFFF !important;
    border-color: transparent !important;
    box-shadow: 0 8px 25px rgba(14, 165, 233, 0.42) !important;
}
div[data-testid="stTabs"] button[data-baseweb="tab"][aria-selected="true"] p {
    color: #FFFFFF !important;
}

/* Inputs & Selectboxes - Super Clear Contrast */
div[data-baseweb="select"] > div,
div[data-baseweb="input"] > div,
div[data-testid="stTimeInput"] > div > div,
div[data-testid="stNumberInput"] > div > div,
input[type="text"] {
    background-color: #FFFFFF !important;
    border: 1.5px solid #64748B !important;
    border-radius: 10px !important;
    color: #0F172A !important;
    font-weight: 700 !important;
}

/* Dataframe & Table Card Backdrop */
div[data-testid="stDataFrame"] {
    background: rgba(255, 255, 255, 0.96) !important;
    border-radius: 14px !important;
    padding: 10px !important;
    box-shadow: 0 6px 20px rgba(0,0,0,0.1) !important;
    border: 1px solid #E2E8F0 !important;
}

/* =================================================================
   CẢI TIẾN 1 – TYPE SCALE nhất quán toàn hệ thống
   ================================================================= */
h1, .vmos-h1 { font-size: 32px !important; font-weight: 800 !important; line-height: 1.2 !important; letter-spacing: -0.5px !important; }
h2, .vmos-h2 { font-size: 24px !important; font-weight: 700 !important; line-height: 1.3 !important; }
h3, .vmos-h3 { font-size: 18px !important; font-weight: 700 !important; line-height: 1.4 !important; }
h4, .vmos-h4 { font-size: 15px !important; font-weight: 600 !important; }
p, .vmos-body  { font-size: 15px !important; line-height: 1.65 !important; }
small, .vmos-small { font-size: 13px !important; }

/* =================================================================
   CẢI TIẾN 2 – KPI CARDS: accent color bar + glassmorphism
   ================================================================= */
/* Các metric card của Streamlit */
[data-testid="stMetric"] {
    background: rgba(255,255,255,0.88) !important;
    backdrop-filter: blur(16px) !important;
    -webkit-backdrop-filter: blur(16px) !important;
    border-radius: 16px !important;
    border: 1px solid rgba(14,165,233,0.18) !important;
    box-shadow: 0 4px 20px rgba(14,165,233,0.08) !important;
    padding: 18px 20px 14px 20px !important;
    position: relative !important;
    overflow: hidden !important;
    transition: transform 0.2s ease, box-shadow 0.2s ease !important;
}
[data-testid="stMetric"]::before {
    content: '' !important;
    position: absolute !important;
    top: 0 !important; left: 0 !important; right: 0 !important;
    height: 4px !important;
    background: linear-gradient(90deg, #0EA5E9, #0284C7) !important;
    border-radius: 16px 16px 0 0 !important;
}
[data-testid="stMetric"]:hover {
    transform: translateY(-3px) !important;
    box-shadow: 0 10px 32px rgba(14,165,233,0.18) !important;
}
[data-testid="stMetricValue"] {
    font-size: 28px !important;
    font-weight: 800 !important;
    color: #0F172A !important;
    line-height: 1.1 !important;
}
[data-testid="stMetricLabel"] {
    font-size: 13px !important;
    font-weight: 600 !important;
    color: #64748B !important;
    text-transform: uppercase !important;
    letter-spacing: 0.5px !important;
}
[data-testid="stMetricDelta"] {
    font-size: 13px !important;
    font-weight: 700 !important;
}

/* =================================================================
   CẢI TIẾN 3 – SEPIA / DARK MODE: smooth background transition
   ================================================================= */
.stApp, [data-testid="stAppViewContainer"],
[data-testid="stMain"], [data-testid="stSidebar"],
html, body {
    transition: background-color 0.35s ease, background 0.35s ease,
                border-color 0.25s ease, color 0.25s ease !important;
}

/* =================================================================
   CẢI TIẾN 4 – SIDEBAR: tooltip cho icon mode + hover glow
   ================================================================= */
[data-testid="stSidebar"].vmos-mini [data-testid="stButton"] > button {
    position: relative !important;
}
[data-testid="stSidebar"].vmos-mini [data-testid="stButton"] > button::after {
    content: attr(title) !important;
    position: absolute !important;
    left: calc(100% + 10px) !important;
    top: 50% !important;
    transform: translateY(-50%) !important;
    background: #0F172A !important;
    color: white !important;
    font-size: 12px !important;
    font-weight: 600 !important;
    padding: 5px 10px !important;
    border-radius: 8px !important;
    white-space: nowrap !important;
    pointer-events: none !important;
    opacity: 0 !important;
    transition: opacity 0.2s ease !important;
    z-index: 9999 !important;
    box-shadow: 0 4px 12px rgba(0,0,0,0.3) !important;
}
[data-testid="stSidebar"].vmos-mini [data-testid="stButton"] > button:hover::after {
    opacity: 1 !important;
}
/* Glow active sidebar button */
[data-testid="stSidebar"] [data-testid="stButton"] > button[kind="primary"] {
    box-shadow: 0 0 0 2px rgba(14,165,233,0.35), 0 4px 14px rgba(14,165,233,0.25) !important;
}

/* =================================================================
   CẢI TIẾN 5 – ORG CHART: swipe indicator + scroll container
   ================================================================= */
.org-tree, [data-testid="stHtml"] iframe {
    overflow-x: auto !important;
    -webkit-overflow-scrolling: touch !important;
}
/* Scroll hint gradient on right edge for org chart on mobile */
@media screen and (max-width: 768px) {
    .org-tree::after {
        content: '← scroll →' !important;
        display: block !important;
        text-align: center !important;
        font-size: 12px !important;
        color: #94A3B8 !important;
        padding: 6px 0 0 0 !important;
        letter-spacing: 1px !important;
    }
}

/* =================================================================
   CẢI TIẾN 6 – UPLOAD / STEP INDICATOR: visual progress steps
   ================================================================= */
.vmos-step-bar {
    display: flex;
    align-items: center;
    gap: 0;
    margin-bottom: 28px;
}
.vmos-step {
    display: flex;
    align-items: center;
    gap: 8px;
    flex: 1;
}
.vmos-step-dot {
    width: 32px; height: 32px;
    border-radius: 50%;
    display: flex; align-items: center; justify-content: center;
    font-weight: 800; font-size: 14px;
    flex-shrink: 0;
    transition: all 0.2s ease;
}
.vmos-step-dot.active  { background: linear-gradient(135deg,#0EA5E9,#0284C7); color: white; box-shadow: 0 4px 12px rgba(14,165,233,0.4); }
.vmos-step-dot.done    { background: #10B981; color: white; }
.vmos-step-dot.pending { background: #E2E8F0; color: #94A3B8; }
.vmos-step-label { font-size: 13px; font-weight: 600; color: #334155; }
.vmos-step-label.active { color: #0EA5E9; }
.vmos-step-label.done   { color: #10B981; }
.vmos-step-connector { flex: 1; height: 2px; background: #E2E8F0; margin: 0 8px; }
.vmos-step-connector.done { background: linear-gradient(90deg,#10B981,#0EA5E9); }

/* =================================================================
   CẢI TIẾN 7 – TOP NAV GROUP: tất cả nút fixed top-right căn thẳng
   ================================================================= */
/* Prevent individual nav buttons from overlapping each other on wide screens */
.st-key-btn_top_eyecare_fixed,
.st-key-nav_btn_notif,
.st-key-nav_btn_docs,
.st-key-nav_btn_support,
.st-key-nav_btn_profile,
.st-key-lang_switch_btn {
    /* Shared alignment baseline & un-block clicks */
    top: 13px !important;
    z-index: 2147483647 !important;
    pointer-events: auto !important;
}
/* Disable pointer events on Streamlit header to prevent it from blocking clicks */
/* (Moved to theme styling block below to add dynamic background) */
/* All nav buttons share same hover lift */
.st-key-nav_btn_notif button:hover,
.st-key-nav_btn_docs button:hover,
.st-key-nav_btn_support button:hover,
.st-key-nav_btn_profile button:hover {
    transform: translateY(-2px) !important;
}
/* Text contrast for nav labels */
.st-key-nav_btn_notif p,
.st-key-nav_btn_docs p,
.st-key-nav_btn_support p,
.st-key-nav_btn_profile p {
    font-size: 13.5px !important;
    font-weight: 700 !important;
}
</style>
""", unsafe_allow_html=True)




# ==========================================
# KHỞI TẠO LƯU TRỮ API KEY
# ==========================================


# ==========================================
# CƠ SỞ DỮ LIỆU SQLITE
# ==========================================
DB_FILE = 'attendance.db'


init_db()






# ==========================================
# 1. CẤU HÌNH & ASSETS
# ==========================================
LOGO_HEADER_PATH = os.path.join(ASSETS_DIR, "logo_header.png")

def _logo_img_tag(b64_val: str, style: str = "height:60px;", extra_class: str = "") -> str:
    class_str = f' class="{extra_class}"' if extra_class else ""
    if not b64_val:
        return f'<div style="font-size:32px">📊</div>'
    return f'<img src="data:image/png;base64,{b64_val}" style="{style}"{class_str}>'

LOGO_HEADER_B64 = "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"


def load_bg_base64():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    for fname in [os.path.join("assets", "85a6bc46a8ad5cd80c4ad77e55f2ed64.jpg"), os.path.join("assets", "14b2755d8dc29cc8b1b1288c3d0cfeab.jpg"), os.path.join("assets", "ae425316cef00e69341b6e8ace02c575.jpg"), os.path.join("assets", "340fc3694b764c741a1da80b4c54a18c.jpg"), os.path.join("assets", "media__1782271302118.png"), "bg_fuji.png"]:
        p = os.path.join(base_dir, fname)
        if os.path.exists(p):
            with open(p, "rb") as f:
                mime = "image/jpeg" if fname.lower().endswith((".jpg", ".jpeg")) else "image/png"
                return base64.b64encode(f.read()).decode("utf-8"), mime
    return "", "image/png"

BG_B64, BG_MIME = load_bg_base64()

def load_otaki_base64():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    img_path = os.path.join(base_dir, "assets", "otaki.png")
    if os.path.exists(img_path):
        try:
            with open(img_path, "rb") as f:
                return base64.b64encode(f.read()).decode("utf-8")
        except Exception:
            pass
    return ""

OTAKI_B64 = load_otaki_base64()

# ----- KHỞI TẠO BIẾN CẤU HÌNH MẶC ĐỊNH -----
if "gio_vao_chuan" not in st.session_state: st.session_state.gio_vao_chuan = datetime.time(8, 0)
if "gio_ra_chuan" not in st.session_state: st.session_state.gio_ra_chuan = datetime.time(17, 0)
if "nghi_trua_bat_dau" not in st.session_state: st.session_state.nghi_trua_bat_dau = datetime.time(12, 0)
if "nghi_trua_ket_thuc" not in st.session_state: st.session_state.nghi_trua_ket_thuc = datetime.time(13, 0)
if "so_gio_toi_da" not in st.session_state: st.session_state.so_gio_toi_da = 8.0
if "custom_holidays" not in st.session_state: st.session_state.custom_holidays = set()
if "custom_workdays" not in st.session_state: st.session_state.custom_workdays = set()
if "manual_ot" not in st.session_state: st.session_state.manual_ot = {}
if "manual_leave" not in st.session_state: st.session_state.manual_leave = {}
if "manual_ot_reason" not in st.session_state: st.session_state.manual_ot_reason = {}
if "manual_emps" not in st.session_state: st.session_state.manual_emps = []
if "deleted_emps" not in st.session_state: st.session_state.deleted_emps = set()
if "edited_emps" not in st.session_state: st.session_state.edited_emps = {}
if "leave_files" not in st.session_state: st.session_state.leave_files = []
if "theme_mode" not in st.session_state: st.session_state.theme_mode = 'light'

T = get_theme(st.session_state.theme_mode)
is_sepia_global = (st.session_state.theme_mode == 'sepia')
is_dark_global = (st.session_state.theme_mode == 'dark')
btn_grad = T['primary_gradient']
btn_grad_hover = T['accent_gradient']
btn_shadow = T['shadow']
btn_shadow_hover = T['shadow_glow']

if BG_B64:
    bg_img_css = f'url("data:{BG_MIME};base64,{BG_B64}")' if not (is_sepia_global or is_dark_global) else 'none'
    bg_col_css = '#F8FAFC' if not (is_sepia_global or is_dark_global) else T['bg_app']
    overlay_css = 'linear-gradient(135deg, rgba(255, 255, 255, 0.05) 0%, rgba(248, 250, 252, 0.1) 100%)' if not (is_sepia_global or is_dark_global) else 'transparent'
    opacity_css = '1.0' if not (is_sepia_global or is_dark_global) else '1.0'

    trans_css = "transition: background 0.6s cubic-bezier(0.4, 0, 0.2, 1), background-color 0.6s cubic-bezier(0.4, 0, 0.2, 1), border-color 0.6s cubic-bezier(0.4, 0, 0.2, 1) !important;" if st.session_state.get('theme_just_toggled', False) else ""
    
    # Reset immediately so it only applies for one rerun
    if st.session_state.get('theme_just_toggled', False):
        st.session_state.theme_just_toggled = False
        
    global_css_container.markdown(f"""
    <style>
    [data-testid="stAppViewContainer"]::before {{
        content: "" !important;
        position: fixed !important;
        top: 0 !important; left: 0 !important; right: 0 !important; bottom: 0 !important;
        width: 100vw !important; max-width: 100% !important; height: 100vh !important;
        background-color: {bg_col_css} !important;
        background-image: {bg_img_css} !important;
        background-size: clamp(450px, 45vw, 800px) auto !important;
        background-position: right top !important;
        background-repeat: no-repeat !important;
        opacity: {opacity_css} !important;
        image-rendering: -webkit-optimize-contrast !important;
        image-rendering: crisp-edges !important;
        filter: contrast(108%) saturate(105%) !important;
        z-index: 0 !important;
        pointer-events: none !important;
    }}
    .stApp::before, .vimos-fuji-wallpaper {{
        display: none !important;
    }}
    .vimos-fuji-overlay {{
        position: fixed !important;
        top: 0 !important; left: 0 !important; right: 0 !important; bottom: 0 !important;
        width: 100% !important; height: 100% !important;
        background: {overlay_css} !important;
        z-index: -99998 !important;
        pointer-events: none !important;
        {trans_css}
    }}
    
    /* Vô hiệu hóa lớp phủ mờ trắng "stale" mặc định của Streamlit khi chạy hàm tốn thời gian (ngăn chớp trắng) */
    [data-testid="stAppViewContainer"], [data-testid="stMain"], .stApp {{
        opacity: 1 !important;
    }}
    [data-stale="true"] {{
        opacity: 1 !important;
        transition: none !important;
        filter: none !important;
    }}
    
    .stApp, [data-testid="stAppViewContainer"], [data-testid="stMain"], .main, div[class*="stMainBlockContainer"] {{
        background: transparent !important;
        background-color: transparent !important;
        {trans_css}
    }}
    
    /* Thanh Navbar trên cùng (chứa các nút chức năng) với nền mờ chống che khuất nội dung */
    header[data-testid="stHeader"] {{
        background: transparent !important;
        z-index: 999999 !important;
        pointer-events: auto !important;
        {trans_css}
    }}

    /* ==== BẢNG MÀU XANH BIỂN CHUYÊN NGHIỆP TỐI GIẢN ==== */
    :root {{
        --fuji-gradient: linear-gradient(135deg, #0EA5E9 0%, #0284C7 100%);
        --fuji-gradient-hover: linear-gradient(135deg, #0284C7 0%, #0369A1 100%);
    }}

    /* Tất cả các nút bấm chức năng trên toàn trang web */
    div[data-testid="stMain"] .stButton > button, div[data-testid="stMain"] button[kind="primary"],
    [data-testid="stSidebar"] .stButton > button[kind="primary"] {{
        background: {btn_grad} !important;
        color: white !important;
        border: none !important;
        border-radius: 12px !important;
        font-weight: 700 !important;
        letter-spacing: 0.2px !important;
        box-shadow: 0 4px 14px {btn_shadow} !important;
        transition: all 0.25s ease !important;
    }}

    div[data-testid="stMain"] .stButton > button:hover, div[data-testid="stMain"] button[kind="primary"]:hover,
    [data-testid="stSidebar"] .stButton > button[kind="primary"]:hover {{
        background: {btn_grad_hover} !important;
        transform: translateY(-1.5px) !important;
        box-shadow: 0 6px 20px {btn_shadow_hover} !important;
    }}
    div[data-testid="stMain"] .stButton > button p,
    [data-testid="stSidebar"] .stButton > button[kind="primary"] p {{
        color: white !important;
        fill: white !important;
        font-weight: 700 !important;
    }}
    </style>
    <div class="vimos-fuji-wallpaper"></div>
    <div class="vimos-fuji-overlay"></div>
    """, unsafe_allow_html=True)

# Inject viewport meta tag for correct mobile scaling
st.markdown("""
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
<style>
/* Ensure Streamlit doesn't override viewport on embedded iframe */
iframe { max-width: 100% !important; }
/* Touch-friendly tap targets for mobile Checkin GPS */
@media screen and (max-width: 768px) {
    .stButton > button {
        min-height: 48px !important;
        font-size: 15px !important;
    }
    .stSelectbox, .stTextInput input, .stTextArea textarea {
        font-size: 16px !important; /* Prevents iOS auto-zoom on focus */
    }
    /* Checkin GPS map embed fits screen width */
    iframe[title*="map"], iframe[title*="Map"] {
        width: 100% !important;
        min-height: 260px !important;
    }
}
</style>
""", unsafe_allow_html=True)

st.markdown(f"""
<style>
:root {{
    --brand-primary:  {T['primary']};
    --brand-primary-hover: {T['primary_hover']};
    --brand-gradient: {T['primary_gradient']};
    --accent:         {T['accent']};
    --accent-gradient:{T['accent_gradient']};
    --bg-app:         {T['bg_app']};
    --bg-content:     {T['bg_content']};
    --bg-card:        {T['bg_card']};
    --bg-card-hover:  {T['bg_card_hover']};
    --text-primary:   {T['text_primary']};
    --text-secondary: {T['text_secondary']};
    --text-tertiary:  {T['text_tertiary']};
    --border:         {T['border']};
    --border-light:   {T['border_light']};
    --shadow:         {T['shadow']};
    --shadow-glow:    {T['shadow_glow']};
    --card-shadow:    {T['shadow']};
    --card-hover-shadow: {T['shadow_glow']};
    --ink-900: {T['text_primary']};
    --ink-700: {T['text_secondary']};
    --ink-500: {T['text_tertiary']};
    --line:    {T['border']};
}}

html, body, .stApp, [data-testid="stAppViewContainer"] {{
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif !important;
    color: {T['text_primary']};
    overflow-x: hidden !important;
    max-width: 100% !important;
    box-sizing: border-box !important;
}}

#MainMenu {{visibility: hidden;}}
footer {{visibility: hidden; display: none !important;}}
    [data-testid="stToolbar"] {{ background: transparent !important; }}

@media (max-width: 768px) {{
    .mobile-hide-footer {{ display: none !important; }}
}}

.block-container {{
    padding-top: 1.5rem !important;
    padding-left: 2rem !important;
    padding-right: 2rem !important;
    padding-bottom: 2rem !important;
    max-width: 98% !important;
    background: transparent !important;
}}

/* ===== App Header Banner ===== */
.app-header {{
    position: relative; overflow: hidden;
    background: {T['primary_gradient']};
    border-radius: 24px; padding: 36px 48px;
    display: flex; align-items: center; gap: 36px;
    margin-bottom: 32px;
    box-shadow: {T['shadow_glow']};
    color: white;
}}
.app-header::before {{
    content: ''; position: absolute; top: -50%; right: -5%;
    width: 350px; height: 350px;
    background: radial-gradient(circle, rgba(255,255,255,0.15) 0%, rgba(255,255,255,0) 70%);
    border-radius: 50%; pointer-events: none;
}}
.app-header::after {{
    content: ''; position: absolute; bottom: -40%; right: 15%;
    width: 250px; height: 250px;
    background: radial-gradient(circle, rgba(255,255,255,0.1) 0%, rgba(255,255,255,0) 70%);
    border-radius: 50%; pointer-events: none;
}}
.app-header-icon {{
    font-size: 32px; display: flex; align-items: center; justify-content: center;
    background: rgba(255,255,255,0.15); border-radius: 100px; padding: 12px 32px;
    box-shadow: 0 10px 30px rgba(0, 0, 0, 0.15); z-index: 1;
    transition: transform 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.275);
    backdrop-filter: blur(8px);
}}
.app-header-icon:hover {{ transform: scale(1.05) translateY(-2px); }}
.app-header-icon img {{ height: 64px; width: auto; display: block; mix-blend-mode: {'normal' if T['name'] == 'dark' else 'multiply'}; }}
.app-header-title {{ font-size: 30px; font-weight: 800; margin: 0; letter-spacing: -0.02em; z-index: 1; position: relative; color: white; }}
.app-header-sub {{ font-size: 16px; color: rgba(255,255,255,0.9); margin: 8px 0 0 0; z-index: 1; position: relative; font-weight: 500; }}
.app-header-badge {{
    margin-left: auto;
    background: rgba(255,255,255,0.1); backdrop-filter: blur(12px);
    border: 1px solid rgba(255,255,255,0.25); color: white;
    font-size: 14.5px; font-weight: 600; padding: 12px 24px;
    border-radius: 100px; box-shadow: 0 8px 24px rgba(0,0,0,0.1);
    z-index: 1; transition: all 0.3s ease;
}}
.app-header-badge:hover {{ background: rgba(255,255,255,0.2); transform: translateY(-2px); }}

/* ===== Cards ===== */
.card {{ background: {T['bg_card']}; border: 1px solid {T['border']}; border-radius: 16px; padding: 22px 26px; margin-bottom: 18px; box-shadow: {T['shadow']}; transition: transform 0.2s ease, box-shadow 0.2s ease; }}
.card:hover {{ transform: translateY(-3px); box-shadow: {T['shadow_glow']}; }}
.card-title {{ font-size: 15px; font-weight: 700; color: {T['text_primary']}; margin: 0 0 16px 0; display: flex; align-items: center; gap: 12px; letter-spacing: -0.01em; }}
.card-icon {{ display: inline-flex; align-items: center; justify-content: center; width: 32px; height: 32px; font-size: 16px; border-radius: 10px; background: {T['bg_card_hover']}; color: {T['primary']}; }}

/* Upload hint */
.upload-hint {{ background-color: {T['bg_card_hover']}; border-radius: 12px; padding: 12px 18px; font-size: 14px; color: {T['primary']}; font-weight: 500; margin-bottom: 16px; border: 1px solid {T['border']}; }}

/* Buttons */
div[data-testid="stButton"], div[data-testid="stDownloadButton"] {{
    display: flex;
    justify-content: center;
    width: 100%;
}}
div[data-testid="stButton"] > button, div[data-testid="stDownloadButton"] > button {{
    background: {T['primary_gradient']} !important;
    color: white !important; border: none !important; border-radius: 14px !important;
    padding: 0 28px !important; font-size: 14px !important; font-weight: 600 !important;
    box-shadow: {T['shadow_glow']} !important; transition: all 0.2s ease !important;
    width: auto !important;
    height: 40px !important;
    min-height: 40px !important;
    line-height: 40px !important;
    white-space: nowrap !important;
}}
div[data-testid="stButton"] > button:hover, div[data-testid="stDownloadButton"] > button:hover {{
    background: {T['accent_gradient']} !important;
    transform: translateY(-2px) !important; box-shadow: {T['shadow_glow']} !important; filter: brightness(1.05) !important;
}}
div[data-testid="stButton"] > button:active, div[data-testid="stDownloadButton"] > button:active {{ transform: translateY(0) !important; }}

/* Streamlit components */
[data-testid="stFileUploader"] {{ background: {T['bg_card']}; border-radius: 16px; padding: 12px; border: 1.5px dashed {T['primary']} !important; box-shadow: {T['shadow']}; transition: all 0.25s ease !important; }}
[data-testid="stFileUploader"]:hover {{ border-color: {T['accent']} !important; background: {T['bg_card_hover']} !important; box-shadow: {T['shadow_glow']} !important; transform: translateY(-2px); }}
[data-testid="stFileUploaderDropzone"] {{ padding: 20px !important; border-radius: 12px !important; }}
[data-testid="stDataFrame"] {{ background: {T['bg_card']} !important; border-radius: 16px; overflow: hidden; border: 1px solid {T['border']}; box-shadow: {T['shadow']}; transition: box-shadow 0.3s ease; }}
[data-testid="stDataFrame"]:hover {{ box-shadow: {T['shadow_glow']}; }}
[data-baseweb="input"] > div, [data-baseweb="textarea"] > div, [data-baseweb="select"] > div {{ border-radius: 10px !important; border: 1.5px solid {T['border']} !important; background-color: {T['bg_content']} !important; transition: all 0.25s ease !important; color: {T['text_primary']} !important; }}
[data-baseweb="input"] > div:hover, [data-baseweb="textarea"] > div:hover, [data-baseweb="select"] > div:hover {{ border-color: {T['primary']} !important; box-shadow: {T['shadow_glow']} !important; transform: translateY(-1px); }}
[data-baseweb="input"] > div:focus-within, [data-baseweb="textarea"] > div:focus-within, [data-baseweb="select"] > div:focus-within {{ border-color: {T['primary']} !important; box-shadow: {T['shadow_glow']} !important; transform: translateY(-1px); }}

[data-testid="stDateInput"] input, [data-testid="stTimeInput"] input, [data-testid="stNumberInput"] input, [data-testid="stTextInput"] input, [data-testid="stTextArea"] textarea {{ border: none !important; box-shadow: none !important; background: transparent !important; color: {T['text_primary']} !important; }}

[data-testid="stExpander"] {{ border: 1px solid {T['border']} !important; border-radius: 16px !important; background: {T['bg_card']} !important; box-shadow: {T['shadow']} !important; transition: all 0.25s ease !important; }}
[data-testid="stExpander"]:hover {{ box-shadow: {T['shadow_glow']} !important; border-color: {T['primary']} !important; transform: translateY(-2px); }}
/* Container box widgets */
[data-testid="stVerticalBlockBorderWrapper"] {{ border-radius: 16px !important; background: {T['bg_card']} !important; box-shadow: {T['shadow']} !important; transition: transform 0.2s ease, box-shadow 0.2s ease; border: 1px solid {T['border']} !important; padding: 20px !important;}}
[data-testid="stVerticalBlockBorderWrapper"]:hover {{ box-shadow: {T['shadow_glow']} !important; transform: translateY(-2px); }}

div[data-testid="stTabs"] button[data-baseweb="tab"][aria-selected="true"] {{ background: {T.get('primary_gradient', T['primary'])} !important; color: white !important; border-color: transparent !important; box-shadow: {T['shadow_glow']} !important; }}
div[data-testid="stTabs"] button[data-baseweb="tab"][aria-selected="true"] p {{ color: white !important; }}
div[data-testid="stTabs"] button[data-baseweb="tab"]:not([aria-selected="true"]):hover {{ border-color: {T['primary']} !important; color: {T['primary']} !important; box-shadow: {T['shadow_glow']} !important; transform: translateY(-2px); }}
/* Metrics */
[data-testid="stMetric"] {{
    background: {T['bg_card']};
    border: 1px solid {T['border']}; border-radius: 16px; padding: 18px 20px;
    box-shadow: {T['shadow']}; border-left: 4px solid {T['primary']};
    transition: transform 0.2s ease, box-shadow 0.2s ease;
}}
[data-testid="stMetric"]:hover {{ transform: translateY(-2px); box-shadow: {T['shadow_glow']}; }}
[data-testid="stMetricLabel"] {{ font-size: 13px !important; color: {T['text_tertiary']} !important; font-weight: 500 !important; }}
[data-testid="stMetricValue"] {{ font-family: 'Inter', 'Be Vietnam Pro', sans-serif !important; font-size: 26px !important; font-weight: 700 !important; color: {T['text_primary']} !important; }}

.stSpinner > div {{ border-top-color: {T['primary']} !important; }}
.stAlert {{ border-radius: 12px !important; border: none !important; box-shadow: {T['shadow']} !important; }}
[data-testid="stSelectbox"] > div {{ border-radius: 10px !important; }}
[data-testid="stMultiSelect"] > div > div {{ border-radius: 10px !important; }}
[data-baseweb="tag"] {{ background-color: {T['bg_card_hover']} !important; color: {T['primary']} !important; border-radius: 6px !important; }}
.stSubheader, h3 {{ font-size: 16px !important; font-weight: 700 !important; color: {T['text_primary']} !important; letter-spacing: -0.01em !important; }}

/* Sidebar */
[data-testid="stSidebar"] {{
    background: {'linear-gradient(180deg, #0B1120, #0F172A, #1E293B)' if T['name'] == 'dark' else T['bg_card']} !important;
    border-right: 1px solid {T['border']} !important;
}}
[data-testid="stSidebar"] *:not(.vmos-company-subtitle) {{ color: {T['text_primary']} !important; }}
html body [data-testid="stSidebar"] .vmos-company-subtitle,
html body [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] .vmos-company-subtitle {{
    color: #0077BE !important;
}}

/* ===== Ẩn chữ "Select all" mặc định của Streamlit trong multiselect ===== */
[data-baseweb="menu"] li:first-child > div[aria-selected] {{
    display: none !important;
}}
[data-baseweb="menu"] [data-testid="stMultiSelectOptionAll"] {{
    display: none !important;
}}
/* ===== Tùy chỉnh màu nền cho Menu thả xuống (Selectbox / Multiselect) ===== */
div[data-baseweb="popover"] > div,
div[data-baseweb="popover"] ul,
div[data-baseweb="menu"],
div[data-baseweb="menu"] ul,
ul[data-testid="stVirtualDropdown"],
div[data-testid="stVirtualDropdown"],
div[role="listbox"],
ul[role="listbox"] {{
    background-color: {T['bg_card']} !important;
}}

li[role="option"],
div[data-baseweb="menu"] li,
[data-testid="stVirtualDropdown"] li {{
    background-color: {T['bg_card']} !important;
    color: {T['text_primary']} !important;
}}

li[role="option"]:hover,
li[role="option"]:focus,
li[role="option"][aria-selected="true"],
div[data-baseweb="menu"] li:hover,
div[data-baseweb="menu"] li[aria-selected="true"] {{
    background-color: {T['bg_card_hover']} !important;
    color: {T['primary']} !important;
}}
/* Dropdown Menu & Top Right Nav */
.global-top-right-nav {{
    position: fixed;
    top: 15px;
    right: 25px;
    z-index: 999999;
    display: flex;
    align-items: center;
    gap: 24px;
    font-family: 'Inter', 'Be Vietnam Pro', sans-serif;
    color: {T['text_primary']};
}}
.menu-dropdown {{
    position: relative;
    display: inline-block;
    padding-bottom: 30px;
    margin-bottom: -30px;
}}
.dropdown-content {{
    visibility: hidden;
    opacity: 0;
    position: absolute;
    top: 100%;
    left: 50%;
    transform: translateX(-50%) translateY(10px);
    background-color: {T['bg_card']};
    min-width: 340px;
    box-shadow: {T['shadow_glow']};
    border-radius: 16px;
    padding: 16px;
    z-index: 1000000;
    display: flex;
    flex-direction: column;
    gap: 12px;
    transition: all 0.3s ease;
    cursor: default;
    border: 1px solid {T['border']};
}}
.menu-dropdown:hover .dropdown-content {{
    visibility: visible;
    opacity: 1;
    transform: translateX(-50%) translateY(0);
}}
.menu-dropdown.right-align .dropdown-content {{
    left: auto;
    right: 0;
    transform: translateY(10px);
}}
.menu-dropdown.right-align:hover .dropdown-content {{
    transform: translateY(0);
}}
.feature-card {{
    background: {T['bg_card_hover']};
    border: 1px solid {T['border']};
    border-radius: 12px;
    padding: 16px;
    text-align: left;
    box-shadow: {T['shadow']};
    display: grid;
    grid-template-columns: auto 1fr;
    column-gap: 16px;
    row-gap: 4px;
    align-items: start;
    cursor: pointer;
    transition: background 0.2s ease;
}}
.feature-card:hover {{
    background: {T['bg_content']};
    border-color: {T['primary']};
}}
.fc-icon {{
    font-size: 20px;
    background: {T['bg_content']};
    width: 36px; height: 36px;
    display: inline-flex;
    align-items: center; justify-content: center;
    border-radius: 8px;
    grid-row: span 2;
    margin-bottom: 0;
}}
.feature-card h3 {{
    font-size: 14px;
    font-weight: 700;
    color: {T['text_primary']};
    margin: 0;
}}
.feature-card p {{
    font-size: 13px;
    color: {T['text_secondary']};
    margin: 0;
    line-height: 1.4;
}}
.contact-box {{
    display: flex; align-items: center; gap: 14px; background: {T['bg_card_hover']}; padding: 14px;
    border-radius: 12px; border: 1px solid {T['border']}; transition: all 0.2s ease; cursor: default;
}}
.contact-box:hover {{
    border-color: {T['primary']}; background: {T['bg_content']}; transform: translateY(-2px); box-shadow: {T['shadow_glow']};
}}
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. HÀM TIỆN ÍCH
# ==========================================




# ==========================================
# 3. LOGIC TÍNH TOÁN
# ==========================================







# ==========================================
# 4. PARSER ĐỌC FILE
# ==========================================



# ==========================================
# 5. EXPORTER XUẤT FILE
# ==========================================

# ==========================================
# 6. GIAO DIỆN CHÍNH
# ==========================================
if "step" not in st.session_state: st.session_state.step = 1
if "df_raw" not in st.session_state: st.session_state.df_raw = None
if "show_history" not in st.session_state: st.session_state.show_history = False

# ----- QUẢN LÝ TRẠNG THÁI TRANG (ROUTING) -----
if "app_page" not in st.session_state:
    st.session_state.app_page = "overview"

# ----- QUẢN LÝ HIỂN THỊ SIDEBAR TOÀN CỤC -----
st.markdown("""
<style>
    /* Cho phép sidebar hiển thị và có thể kéo thay đổi kích thước, nhưng không block tính năng thu gọn mặc định */
    [data-testid="stSidebar"] { 
        background-color: white !important;
        box-shadow: 2px 0 10px rgba(0,0,0,0.06) !important;
        z-index: 999999 !important;
    }
    
    /* Cấu hình hiển thị sidebar đã được chuyển sang JS custom component */
</style>
""", unsafe_allow_html=True)

# ----- NGÔN NGỮ (LANGUAGE) -----
if "lang" not in st.session_state:
    st.session_state.lang = "vi"

def render_lang_toggle():
    """Render thanh điều hướng trên cùng bên phải: Ngôn ngữ, Dịu mắt, và các nút chức năng."""
    is_ja = (st.session_state.lang == "ja")
    lang_code = st.session_state.lang
    if 'theme_mode' not in st.session_state:
        st.session_state.theme_mode = 'light'
    T = get_theme(st.session_state.theme_mode)
    is_sepia = (st.session_state.theme_mode == 'sepia')
    is_dark = (st.session_state.theme_mode == 'dark')

    track_class = "ja-active" if is_ja else "vi-active"
    vi_class    = "inactive"  if is_ja else "active"
    ja_class    = "active"    if is_ja else "inactive"

    st.markdown(f"""
    <style>
    /* Removed conflicting stHeader styling to restore top navbar */

    /* ===== Language Toggle Button – pill có cờ + text ===== */
    .st-key-lang_switch_btn {{
        position: fixed !important;
        top: 13px !important;
        right: 600px !important;
        z-index: 999999 !important;
        width: 110px !important;
    }}
    .st-key-lang_switch_btn > div {{ width: 110px !important; }}
    .st-key-lang_switch_btn button {{
        background: {T['bg_card']}99 !important;
        backdrop-filter: blur(14px) !important;
        border: 1.5px solid {T['border']} !important;
        border-radius: 50px !important;
        width: 110px !important;
        height: 32px !important;
        min-height: 32px !important;
        padding: 0 12px !important;
        cursor: pointer !important;
        box-shadow: 0 3px 12px rgba(14,165,233,0.1) !important;
        transition: all 0.2s ease !important;
    }}
    .st-key-lang_switch_btn button:hover {{
        background: {T['primary_gradient']} !important;
        border-color: transparent !important;
        box-shadow: 0 5px 16px rgba(14,165,233,0.28) !important;
        transform: translateY(-1px) !important;
    }}
    .st-key-lang_switch_btn button p {{
        display: block !important;
        font-size: 13px !important;
        font-weight: 800 !important;
        color: {T['text_primary']} !important;
        white-space: nowrap !important;
        letter-spacing: 0.3px !important;
        margin: 0 !important;
        line-height: 1 !important;
    }}
    .st-key-lang_switch_btn button:hover p {{
        color: white !important;
    }}
    </style>
    """, unsafe_allow_html=True)

    def toggle_lang():
        st.session_state.lang = "ja" if st.session_state.lang == "vi" else "vi"
    _lang_btn_lbl = "🇻🇳 VI → JA" if st.session_state.lang == "vi" else "🇯🇵 JA → VI"
    st.button(_lang_btn_lbl, key="lang_switch_btn", help="Chuyển đổi ngôn ngữ / 言語切替", on_click=toggle_lang)

    # ----- CHẾ ĐỘ SÁNG / DỊU MẮT -----
    curr_t = st.session_state.get('theme_mode', 'light')
    if curr_t == 'sepia':
        lbl_eyecare = "☀️ Sáng" if lang_code == 'vi' else "☀️ ライト"
    else:
        lbl_eyecare = "☕ Dịu mắt" if lang_code == 'vi' else "☕ セピア"

    # ----- CÁC NÚT ĐIỀU HƯỚNG QUẢN TRỊ NỘI BỘ (THÔNG BÁO, TÀI LIỆU, HỖ TRỢ, TÀI KHOẢN) -----
    def open_nav_modal(tab_name):
        st.session_state.active_nav_modal = tab_name

    lbl_notif   = "🔔 Thông báo" if lang_code == 'vi' else "🔔 お知らせ"
    lbl_docs    = "📚 Tài liệu" if lang_code == 'vi' else "📚 ドキュメント"
    lbl_support = "🛠️ Hỗ trợ" if lang_code == 'vi' else "🛠️ サポート"
    lbl_profile = "👤 Tài khoản" if lang_code == 'vi' else "👤 マイページ"

    st.markdown(f"""
    <style>
    /* Add a solid background to the header so content scrolling under won't overlap with transparent buttons */
    .stAppHeader, header[data-testid="stHeader"] {{
        background-color: {T['bg_app']} !important;
        z-index: 999990 !important;
    }}
    .st-key-btn_top_eyecare_fixed, .st-key-nav_btn_profile, .st-key-nav_btn_support, .st-key-nav_btn_docs, .st-key-nav_btn_notif {{
        position: fixed !important;
        top: 14px !important;
        z-index: 999999 !important;
    }}
    div[data-testid="stElementContainer"]:has([class*="st-key-"]),
    div[data-testid="stElementContainer"]:has([class*="btn_top_"]),
    div[data-testid="stElementContainer"]:has([class*="lang_switch"]),
    div[data-testid="stElementContainer"]:has([class*="nav_btn_"]) {{
        z-index: 999999 !important;
    }}
    .st-key-btn_top_eyecare_fixed {{ right: 470px !important; width: 125px !important; top: 13px !important; }}
    .st-key-nav_btn_notif         {{ right: 360px !important; width: 105px !important; }}
    .st-key-nav_btn_docs          {{ right: 225px !important; width: 130px !important; }}
    .st-key-nav_btn_support       {{ right: 125px !important; width: 95px !important; }}
    .st-key-nav_btn_profile       {{ right: 20px !important;  width: 100px !important; }}

    .st-key-btn_top_eyecare_fixed > div, .st-key-nav_btn_profile > div, .st-key-nav_btn_support > div, .st-key-nav_btn_docs > div, .st-key-nav_btn_notif > div {{
        width: 100% !important;
    }}

    /* Nút Chế độ Sáng/Tối giữ nguyên dạng nút nhỏ gọn sang trọng */
    .st-key-btn_top_eyecare_fixed button {{
        background: {T['bg_card']} !important;
        backdrop-filter: blur(14px) !important;
        border: 1.5px solid {T['border']} !important;
        border-radius: 50px !important;
        height: 30px !important;
        min-height: 30px !important;
        padding: 0 10px !important;
        box-shadow: 0 3px 12px rgba(0,0,0,0.08) !important;
        cursor: pointer !important;
        transition: all 0.2s ease !important;
    }}
    .st-key-btn_top_eyecare_fixed button:hover {{
        background: linear-gradient(135deg, #0EA5E9, #0284C7) !important;
        border-color: transparent !important;
        transform: translateY(-1.5px) !important;
    }}
    .st-key-btn_top_eyecare_fixed button:hover p {{ color: white !important; }}
    .st-key-btn_top_eyecare_fixed button p {{
        color: {T['text_primary']} !important;
        font-weight: 700 !important;
        font-size: 12.5px !important;
        margin: 0 !important;
    }}
    .stChatFloatingInputContainer {{
        background: transparent !important;
    }}
    /* HIDE STREAMLIT RUNNING STATUS WIDGET (AUTO-REFRESH FLASH) */
    [data-testid="stStatusWidget"],
    [data-testid="stAppStatusWidget"] {{
        display: none !important;
        visibility: hidden !important;
        opacity: 0 !important;
    }}

    /* 4 nút quản trị nội bộ dạng Text links chữ rõ nét dễ đọc */
    .st-key-nav_btn_profile div[data-testid="stButton"] > button,
    .st-key-nav_btn_support div[data-testid="stButton"] > button,
    .st-key-nav_btn_docs div[data-testid="stButton"] > button,
    .st-key-nav_btn_notif div[data-testid="stButton"] > button {{
        background: transparent !important;
        background-color: transparent !important;
        background-image: none !important;
        border: none !important;
        box-shadow: none !important;
        backdrop-filter: none !important;
        height: 30px !important;
        min-height: 30px !important;
        padding: 0 !important;
        cursor: pointer !important;
    }}
    .st-key-nav_btn_profile div[data-testid="stButton"] > button:hover,
    .st-key-nav_btn_support div[data-testid="stButton"] > button:hover,
    .st-key-nav_btn_docs div[data-testid="stButton"] > button:hover,
    .st-key-nav_btn_notif div[data-testid="stButton"] > button:hover {{
        background: transparent !important;
        background-color: transparent !important;
        background-image: none !important;
        border: none !important;
        box-shadow: none !important;
        transform: translateY(-1.5px) !important;
        filter: none !important;
    }}
    .st-key-nav_btn_profile div[data-testid="stButton"] > button p,
    .st-key-nav_btn_support div[data-testid="stButton"] > button p,
    .st-key-nav_btn_docs div[data-testid="stButton"] > button p,
    .st-key-nav_btn_notif div[data-testid="stButton"] > button p {{
        color: {T['text_primary']} !important;
        font-weight: 800 !important;
        font-size: 14px !important;
        margin: 0 !important;
        text-shadow: none !important;
        transition: color 0.2s ease !important;
    }}
    .st-key-nav_btn_profile div[data-testid="stButton"] > button:hover p,
    .st-key-nav_btn_support div[data-testid="stButton"] > button:hover p,
    .st-key-nav_btn_docs div[data-testid="stButton"] > button:hover p,
    .st-key-nav_btn_notif div[data-testid="stButton"] > button:hover p {{
        color: {T['primary']} !important;
    }}
    </style>
    """, unsafe_allow_html=True)

    def toggle_theme_mode():
        curr = st.session_state.get('theme_mode', 'light')
        st.session_state.theme_mode = 'sepia' if curr != 'sepia' else 'light'
        st.session_state.theme_just_toggled = True
    st.button(lbl_eyecare, key="btn_top_eyecare_fixed", on_click=toggle_theme_mode, help="Chuyển đổi Sáng / Dịu mắt" if lang_code == 'vi' else "ライト/セピア切替")
    
    @st.dialog("V.MOS Enterprise", width="large")
    def vmos_dialog(m_type, lang_code):
        if m_type == 'notif':
            title_txt = "🔔 Thông báo hệ thống V.MOS" if lang_code == 'vi' else "🔔 システムお知らせ"
            pending_reqs = [p for p in st.session_state.get('pending_hr_approvals', []) if p['status'] == '⏳ Chờ duyệt']
            extra_notif = ""
            if pending_reqs:
                if lang_code == 'vi':
                    extra_notif = f"""<div style="padding: 12px 16px; background: #FEF3C7; border-left: 4px solid #F59E0B; border-radius: 8px; margin-bottom: 12px;">
<b style="color: #D97706;">🔔 [HR Yêu cầu phê duyệt] Có {len(pending_reqs)} đơn đăng ký Phép/OT mới từ nhân viên!</b><br>
Kỹ sư {pending_reqs[0]['emp']} vừa gửi đơn {pending_reqs[0]['type']}. <b style="color:#B45309;">Vào mục Đăng ký Phép/OT để duyệt ngay.</b>
</div>"""
                else:
                    extra_notif = f"""<div style="padding: 12px 16px; background: #FEF3C7; border-left: 4px solid #F59E0B; border-radius: 8px; margin-bottom: 12px;">
<b style="color: #D97706;">🔔 [人事承認待ち] 従業員から未承認の休暇・残業申請が {len(pending_reqs)} 件あります！</b><br>
{pending_reqs[0]['emp']} 様からの {pending_reqs[0]['type']} 申請等。
</div>"""

            import sqlite3
            from db import DB_FILE
            
            conn = sqlite3.connect(DB_FILE)
            c = conn.cursor()
            try:
                c.execute("SELECT date_str, title_vi, content_vi, title_ja, content_ja, type FROM system_notifications ORDER BY id DESC")
                notifs = c.fetchall()
            except sqlite3.OperationalError:
                notifs = []
            conn.close()
            
            notif_html = ""
            for n in notifs:
                date_str, title_vi, content_vi, title_ja, content_ja, n_type = n
                
                # Colors based on type
                if n_type == 'secondary':
                    bg_col = "#F8FAFC"
                    border_col = "#64748B"
                    title_col = "#475569"
                elif n_type == 'warning':
                    bg_col = "#FEF3C7"
                    border_col = "#F59E0B"
                    title_col = "#D97706"
                else: # info
                    bg_col = "#F0F9FF"
                    border_col = "#0EA5E9"
                    title_col = "#0284C7"
                    
                title = title_vi if lang_code == 'vi' else title_ja
                content = content_vi if lang_code == 'vi' else content_ja
                
                date_val = date_str
                try:
                    if "/" in date_val:
                        parts = date_val.split("/")
                        if len(parts) == 3:
                            if len(parts[0]) == 4: # YYYY/MM/DD
                                d_obj = datetime.datetime(int(parts[0]), int(parts[1]), int(parts[2]))
                            else: # DD/MM/YYYY
                                d_obj = datetime.datetime(int(parts[2]), int(parts[1]), int(parts[0]))
                            date_val = d_obj.strftime("%d/%m/%Y") if lang_code == 'vi' else d_obj.strftime("%Y/%m/%d")
                except Exception:
                    pass
                
                notif_html += f"""<div style="padding: 12px 16px; background: {bg_col}; border-left: 4px solid {border_col}; border-radius: 8px; margin-bottom: 12px;">
<b style="color: {title_col};">[{date_val}] {title}</b><br>
{content}
</div>"""

            if not notifs:
                notif_html = "<div style='color: #64748B; font-style: italic;'>Không có thông báo mới / 新しいお知らせはありません</div>"

            html_cnt = extra_notif + f"""<div style="line-height: 1.8; color: #334155; font-size: 14px;">
{notif_html}
</div>"""
        elif m_type == 'docs':
            title_txt = "📚 Tài liệu & Hướng dẫn nội bộ" if lang_code == 'vi' else "📚 ドキュメント＆利用マニュアル"
            html_cnt = """<div style="line-height: 1.85; color: #334155; font-size: 14px;">
<p>Hệ thống quản trị tài liệu chuẩn kỹ thuật Nhật Bản của VIET.MOS COMPANY LIMITED:</p>
<ul style="margin-left: 20px; margin-top: 8px;">
<li>📄 <b>Quy trình chấm công & Tự động hóa:</b> <a href="data:text/plain;charset=utf-8;base64,SMaw4bubbmcgZOG6q24gVi5NT1MgRW50ZXJwcmlzZSB2My4wCgoxLiBRdXkgdOG6r2MgdMOtbmgga+G7syBjw7RuZzogVOG7qyAyMS90aMOhbmcgdHLGsOG7m2MgxJHhur9uIDIwL3Row6FuZyBuw6B5LgoyLiBHaeG7nSB0acOqdSBjaHXhuqluOiBUcuG7qyBjw6FjIG5nw6B5IFQ3LCBDTiB2w6AgTmdo4buJIGzhu4UsIGPhu5luZyB0aMOqbSAxIG5nw6B5IFQ3IGN14buRaSB0aMOhbmcuCjMuIELhuqNuZyBLUEkgc+G6vSB04buxIMSR4buZbmcgxJHhur9tIHPhu5EgbmfGsOG7nWkgdsOgIHPhu5EgZ2nhu50u" download="HuongDanV3.txt" style="color:#0EA5E9; text-decoration:none;">Tải hướng dẫn (.TXT)</a></li>
<li>📊 <b>Quy định kê khai công số dự án MOS:</b> <a href="data:text/plain;charset=utf-8;base64,UXV5IMSR4buLbmgga8OqIGtoYWkgY8O0bmcgc+G7kSBk4buxIMOhbiBNT1M6Ci0gxJBp4buBbiDEkeG6p3kgxJHhu6cgVMOqbiBuaMOibiB2acOqbiB2w6AgTmfDoHkgdGjDoW5nLgotIEtoYWkgYsOhbyBz4buRIGdp4budIGzDoG0gdmnhu4djIHRo4buxYyB04bq/IHbDoCBPVC4KLSBMxrB1IGZpbGUg4bufIMSR4buLbmggZOG6oW5nIEV4Y2VsIHRyxrDhu5tjIGtoaSB04bqjaSBsw6puIGjhu4cgdGjhu5FuZy4=" download="QuyDinhMOS.txt" style="color:#0EA5E9; text-decoration:none;">Xem chi tiết (.TXT)</a></li>
<li>🌴 <b>Quy chế đăng ký Nghỉ phép & Tăng ca (OT):</b> <a href="data:text/plain;charset=utf-8;base64,Qmnhu4N1IG3huqt1IMSQxINuZyBrw70gTmdo4buJIHBow6lwICYgVMSDbmcgY2EgKE9UKQoKVMOqbiBuaMOibiB2acOqbjogLi4uLi4uLi4uLi4uLi4uLi4uLi4KTmfDoHkgxJHEg25nIGvDvTogLi4uLi4uLi4uLi4uLi4uLi4uLi4KTG/huqFpIMSRxINuZyBrw70gKE5naOG7iSBwaMOpcCAvIE9UKTogLi4uLi4uLi4uLi4uLi4uLi4uLi4KVGjhu51pIGdpYW46IFThu6sgLi4uLi4uLiDEkeG6v24gLi4uLi4uLgpMw70gZG86IC4uLi4uLi4uLi4uLi4uLi4uLi4uLi4uLi4uLi4uLi4uLi4uLi4uLi4=" download="BieuMauNghiPhepOT.txt" style="color:#0EA5E9; text-decoration:none;">Tải biểu mẫu chuẩn (.TXT)</a></li>
</ul>
</div>""" if lang_code == 'vi' else """<div style="line-height: 1.85; color: #334155; font-size: 14px;">
<p>VIET.MOS COMPANY LIMITED 社内標準エンジニアリングドキュメント：</p>
<ul style="margin-left: 20px; margin-top: 8px;">
<li>📄 <b>勤怠管理＆自動化マニュアル:</b> <a href="data:text/plain;charset=utf-8;base64,Vi5NT1MgRW50ZXJwcmlzZSB2My4wIOWIqeeUqOODnuODi+ODpeOCouODqwoKMS4g5pyf6ZaT6KiI566XOiDliY3mnIgyMeaXpeOBi+OCieW9k+aciDIw5pel44G+44Gn44CCCjIuIOaomea6lueovOWDjeaXpTog5Zyf5pel44O756Wd5pel44KS6Zmk44GN44CB5pyI5pyr44Gu5Zyf5puc5pelMeaXpeOCkui/veWKoOOAggozLiBLUEnooajjga/nt6jpm4blhoXlrrnjgavpgKPli5XjgZfjgaboh6rli5XoqIjnrpfjgZXjgozjgb7jgZnjgII=" download="ManualV3.txt" style="color:#0EA5E9; text-decoration:none;">ダウンロード (.TXT)</a></li>
<li>📊 <b>MOSプロジェクト工数入力規定:</b> <a href="data:text/plain;charset=utf-8;base64,TU9T44OX44Ot44K444Kn44Kv44OI5bel5pWw5YWl5Yqb6KaP5a6aOgotIOawj+WQjeOBqOaXpeS7mOOCkuW/heOBmuiomOWFpeOBl+OBpuOBj+OBoOOBleOBhOOAggotIOWun+WDjeaZgumWk+OBqE9U44KS5q2j44GX44GP55Sz5ZGK44GX44Gm44GP44Gg44GV44GE44CCCi0g44Ki44OD44OX44Ot44O844OJ5YmN44GrRXhjZWzlvaLlvI/jgafkv53lrZjjgZfjgabjgY/jgaDjgZXjgYTjgII=" download="RuleMOS.txt" style="color:#0EA5E9; text-decoration:none;">詳細を見る (.TXT)</a></li>
<li>🌴 <b>休暇＆残業(OT)申請ガイドライン:</b> <a href="data:text/plain;charset=utf-8;base64,5LyR5pqH44O75q6L5qWtKE9UKeeUs+iri+abuAoK5rCP5ZCNOiAuLi4uLi4uLi4uLi4uLi4uLi4uLgrnlLPoq4vml6U6IC4uLi4uLi4uLi4uLi4uLi4uLi4uCueUs+iri+eoruWIpSAo5LyR5pqHIC8gT1QpOiAuLi4uLi4uLi4uLi4uLi4uLi4uLgrmnJ/plpM6IC4uLi4uLi4uLi4uLi4uLi4uLi4uCueQhueUsTogLi4uLi4uLi4uLi4uLi4uLi4uLi4uLi4uLi4uLi4uLi4uLi4uLi4uLg==" download="FormOT.txt" style="color:#0EA5E9; text-decoration:none;">標準フォーマット取得 (.TXT)</a></li>
</ul>
</div>"""
        elif m_type == 'support':
            title_txt = "🛠️ Hỗ trợ kỹ thuật" if lang_code == 'vi' else "🛠️ テクニカルサポート"
            html_cnt = """<div style="line-height: 1.8; color: #334155; font-size: 14.5px;">
<p>Bộ phận Hỗ trợ kỹ thuật luôn sẵn sàng giải đáp sự cố hệ thống:</p>
<div style="background: #F8FAFC; padding: 16px; border-radius: 12px; border: 1px solid #E2E8F0; margin-top: 10px;">
💬 <b>Hotline Kỹ thuật / Zalo:</b> <span style="color:#0284C7; font-weight:700;">...</span><br>
📧 <b>Email Hỗ trợ:</b> <span style="color:#0284C7; font-weight:700;">...</span><br>
🏢 <b>Văn phòng làm việc:</b> Hỗ trợ trực tuyến 08:00 - 17:30 (T2 - T6)
</div>
</div>""" if lang_code == 'vi' else """<div style="line-height: 1.8; color: #334155; font-size: 14.5px;">
<p>テクニカルサポートチームがシステムのトラブルシューティングをサポートします：</p>
<div style="background: #F8FAFC; padding: 16px; border-radius: 12px; border: 1px solid #E2E8F0; margin-top: 10px;">
💬 <b>技術ホットライン / Zalo:</b> <span style="color:#0284C7; font-weight:700;">...</span><br>
📧 <b>サポートメール:</b> <span style="color:#0284C7; font-weight:700;">...</span><br>
🏢 <b>オフィス:</b> 対応時間 08:00 - 17:30 (月〜金)
</div>
</div>"""
        elif m_type == 'profile':
            title_txt = "👤 Hồ sơ cá nhân nhân viên" if lang_code == 'vi' else "👤 従業員プロフィール"
            html_cnt = """<div style="line-height: 1.8; color: #334155; font-size: 14.5px;">
<div style="display:flex; align-items:center; gap: 16px; margin-bottom: 16px; padding-bottom: 16px; border-bottom: 1px solid #E2E8F0;">
<div style="width: 60px; height: 60px; border-radius: 50%; background: linear-gradient(135deg, #0EA5E9, #0284C7); color: white; display:flex; align-items:center; justify-content:center; font-size: 26px; font-weight: 800;">NV</div>
<div>
    <b style="font-size: 18px; color: #0F172A;">Nhân Viên</b><br>
    <span style="color: #64748B; font-size: 13.5px;">Chức vụ: ... • ID: ...</span>
</div>
</div>
<div style="display:grid; grid-template-columns: 1fr 1fr; gap: 10px; font-size: 14px;">
<div>🏢 <b>Đơn vị:</b> ...</div>
<div>🛡️ <b>Quyền hạn:</b> ...</div>
<div>📧 <b>Email:</b> ...</div>
<div>📍 <b>Trạng thái:</b> ...</div>
</div>
</div>""" if lang_code == 'vi' else """<div style="line-height: 1.8; color: #334155; font-size: 14.5px;">
<div style="display:flex; align-items:center; gap: 16px; margin-bottom: 16px; padding-bottom: 16px; border-bottom: 1px solid #E2E8F0;">
<div style="width: 60px; height: 60px; border-radius: 50%; background: linear-gradient(135deg, #0EA5E9, #0284C7); color: white; display:flex; align-items:center; justify-content:center; font-size: 26px; font-weight: 800;">NV</div>
<div>
    <b style="font-size: 18px; color: #0F172A;">従業員</b><br>
    <span style="color: #64748B; font-size: 13.5px;">役職: ... • ID: ...</span>
</div>
</div>
<div style="display:grid; grid-template-columns: 1fr 1fr; gap: 10px; font-size: 14px;">
<div>🏢 <b>部署:</b> ...</div>
<div>🛡️ <b>権限:</b> ...</div>
<div>📧 <b>メール:</b> ...</div>
<div>📍 <b>ステータス:</b> ...</div>
</div>
</div>"""

        st.markdown(f"""
<div style="background: linear-gradient(135deg, #0EA5E9, #0284C7); color: white; padding: 18px 24px; font-weight: 800; font-size: 18px; border-radius: 8px; margin-bottom: 16px;">
{title_txt}
</div>
<div style="padding: 0 8px;">
{html_cnt}
</div>
""", unsafe_allow_html=True)

    if st.button(lbl_profile, key="nav_btn_profile"):
        vmos_dialog('profile', lang_code)
    if st.button(lbl_support, key="nav_btn_support"):
        vmos_dialog('support', lang_code)
    if st.button(lbl_docs,    key="nav_btn_docs"):
        vmos_dialog('docs', lang_code)
    if st.button(lbl_notif,   key="nav_btn_notif"):
        vmos_dialog('notif', lang_code)

    if is_sepia:
        st.markdown("""
        <style>
        /* WARM PARCHMENT SEPIA EYE-CARE MODE (LỌC ÁNH SÁNG XANH) */
        .vimos-fuji-overlay {
            background: linear-gradient(135deg, rgba(254, 252, 232, 0.88) 0%, rgba(253, 246, 178, 0.9) 100%) !important;
        }
        .app-header {
            background: linear-gradient(135deg, #78350F 0%, #92400E 50%, #B45309 100%) !important;
        }
        div[data-testid="stExpander"],
        div[data-testid="stDataFrame"],
        section[data-testid="stSidebar"],
        .st-key-nav_modal_popup,
        div[data-testid="stTabs"] button[data-baseweb="tab"],
        div[data-testid="stMetric"],
        .card,
        [data-testid="stVerticalBlockBorderWrapper"],
        [data-testid="stFileUploader"],
        [data-testid="stFileUploader"] section,
        [data-testid="stFileUploaderDropzone"] {
            background: rgba(254, 252, 232, 0.96) !important; /* Kem giấy sáp ấm áp #FEFCE8 */
            border-color: rgba(217, 119, 6, 0.45) !important;
            box-shadow: 0 10px 30px rgba(180, 83, 9, 0.14) !important;
        }
        div[data-testid="stExpander"] details > summary {
            background: linear-gradient(90deg, #FEFCE8 0%, #FEF3C7 100%) !important;
            color: #78350F !important;
            border-bottom-color: rgba(217, 119, 6, 0.3) !important;
        }
        div[data-testid="stExpander"] details[open] > summary {
            background: linear-gradient(135deg, #78350F 0%, #92400E 100%) !important;
            color: #FFFFFF !important;
            border-bottom: 2.5px solid #F59E0B !important;
        }
        input[type="text"], input[type="time"], input[type="date"], input[type="number"], div[data-baseweb="select"] > div, textarea, div[data-baseweb="textarea"] > div, div[data-baseweb="input"] > div, div[data-baseweb="base-input"],
        div[data-testid="stTimeInput"] > div > div > div, div[data-testid="stTimeInput"] div[data-baseweb="select"] > div, div[data-testid="stTimeInput"] div[data-baseweb="input"] > div {
            background-color: #FEFCE8 !important;
            border-color: #D97706 !important;
            color: #78350F !important;
        }
        h1, h2, h3, h4, p, label, .upload-hint {
            color: #451A03 !important;
        }
        /* Bộ lọc cho các hộp thông báo (st.info, st.success, v.v.) */
        div[data-testid="stAlert"] {
            filter: sepia(0.4) hue-rotate(-10deg) contrast(0.95);
        }
        /* Bắt toàn bộ các thẻ div có mã màu sáng/trắng cứng bằng CSS attribute selector để ép sang Sepia */
        .stMarkdown div[style*="rgba(255,255,255"],
        .stMarkdown div[style*="rgba(255, 255, 255"],
        .stMarkdown div[style*="rgba(240,249,255"],
        .stMarkdown div[style*="#F8FAFC"],
        .stMarkdown div[style*="#F0F9FF"],
        .stMarkdown div[style*="#FFFFFF"] {
            background: rgba(254, 252, 232, 0.96) !important;
            border-color: rgba(217, 119, 6, 0.45) !important;
        }
        /* Bộ lọc cho bảng dữ liệu Data Editor (canvas) để trông dịu mắt hơn */
        div[data-testid="stDataFrame"] {
            filter: sepia(0.65) hue-rotate(-15deg) contrast(0.9) brightness(0.95);
        }
        [data-testid="stFileUploader"] button {
            background-color: #FEFCE8 !important;
            border: 1px solid #D97706 !important;
            color: #92400E !important;
        }
        [data-testid="stFileUploader"] button:hover {
            background-color: #FEF3C7 !important;
        }
        </style>
        """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <style>
        /* Placeholder for non-sepia mode to preserve Streamlit element indices and prevent DOM shifting bugs */
        </style>
        """, unsafe_allow_html=True)






# ==========================================
# CÁC HÀM XỬ LÝ DỮ LIỆU MOS
# ==========================================
import re

import httpx
import json








def render_holiday_makeup_sidebar():
    
    if "custom_holidays" not in st.session_state:
        st.session_state.custom_holidays = set()
    if "custom_workdays" not in st.session_state:
        st.session_state.custom_workdays = set()

    with st.sidebar.expander(t("exp_holidays"), expanded=False):
        st.markdown(f"<small style='color:green;'>{t('holidays_note')}</small>", unsafe_allow_html=True)

        st.markdown(f"<small>{t('holiday_choose')}</small>", unsafe_allow_html=True)
        selected_date = st.date_input(t("holiday_select"), value=datetime.date.today(), label_visibility="collapsed", key="date_in_holiday")
        if st.button(t("btn_add_custom_holiday"), key="btn_add_custom_holiday", use_container_width=True):
            st.session_state.custom_holidays.add(selected_date)
            st.rerun()
        if st.session_state.custom_holidays:
            st.markdown(t("custom_holiday_count", count=len(st.session_state.custom_holidays)), unsafe_allow_html=True)
            holiday_list = [d.strftime('%d/%m/%Y') for d in sorted(st.session_state.custom_holidays)]
            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                selected_to_remove = st.multiselect(t("holiday_del_sel"), options=holiday_list, placeholder=t("holiday_del_placeholder"), label_visibility="collapsed", key="multi_sel_del_holiday")
            with col_btn2:
                if st.button(t("btn_del_selected"), key="btn_del_sel_holiday", use_container_width=True) and selected_to_remove:
                    for d_str in selected_to_remove:
                        try:
                            st.session_state.custom_holidays.remove(datetime.datetime.strptime(d_str, '%d/%m/%Y').date())
                        except Exception as e: logger.warning(f"Lỗi: {e}", exc_info=True); pass
                    st.rerun()
            with col_btn2:
                if st.button(t("btn_del_all"), key="btn_del_all_holiday", use_container_width=True):
                    st.session_state.custom_holidays = set()
                    st.rerun()

    with st.sidebar.expander(t("exp_makeup"), expanded=False):
        st.markdown(f"<small>{t('makeup_note')}</small>", unsafe_allow_html=True)
        selected_makeup = st.date_input(t("makeup_choose"), value=datetime.date.today(), label_visibility="collapsed", key="date_makeup_input")
        if st.button(t("btn_add_makeup"), key="btn_add_makeup", use_container_width=True):
            st.session_state.custom_workdays.add(selected_makeup)
            st.rerun()
        if st.session_state.custom_workdays:
            st.markdown(f"<small><b>{t('makeup_count').format(count=len(st.session_state.custom_workdays))}</b></small>", unsafe_allow_html=True)
            makeup_list = [d.strftime('%d/%m/%Y') for d in sorted(st.session_state.custom_workdays)]
            selected_makeup_remove = st.multiselect(t("makeup_remove_prompt"), options=makeup_list, placeholder="...", label_visibility="collapsed", key="multi_sel_del_makeup")
            col_btn3, col_btn4 = st.columns(2)
            with col_btn3:
                if st.button(t("btn_del_selected"), key="btn_del_sel_makeup", use_container_width=True) and selected_makeup_remove:
                    for d_str in selected_makeup_remove:
                        try: st.session_state.custom_workdays.remove(datetime.datetime.strptime(d_str, '%d/%m/%Y').date())
                        except Exception as e: logger.warning(f"Lỗi: {e}", exc_info=True); pass
                    st.rerun()
            with col_btn4:
                if st.button(t("btn_del_all"), key="btn_del_all_makeup", use_container_width=True):
                    st.session_state.custom_workdays = set()
                    st.rerun()

def render_leave_ot_sidebar():
    with st.sidebar.expander("📂 Khớp nối Nghỉ phép & OT" if st.session_state.lang == 'vi' else "📂 休暇＆残業の登録", expanded=False):
        if "manual_leave" not in st.session_state: st.session_state.manual_leave = {}
        if "manual_ot_reason" not in st.session_state: st.session_state.manual_ot_reason = {}
        
        emp_options = get_company_emp_options(st.session_state.lang)
        
        tab_leave, tab_ot = st.tabs(["🌴 Nghỉ phép" if st.session_state.lang=='vi' else "🌴 休暇登録", "⏰ OT" if st.session_state.lang=='vi' else "⏰ 残業登録"])
        with tab_leave:
            sel_e_leave = st.selectbox("Nhân viên" if st.session_state.lang=='vi' else "社員", emp_options, key="sb_sel_e_leave")
            sel_d_leave = st.date_input("Ngày nghỉ" if st.session_state.lang=='vi' else "休暇日", value=datetime.date.today(), key="sb_sel_d_leave")
            if st.button("➕ Thêm lịch nghỉ" if st.session_state.lang=='vi' else "➕ 休暇を追加", type="secondary", use_container_width=True, key="btn_add_l"):
                if sel_e_leave:
                    ma_l = sel_e_leave.split(" - ")[0].strip().upper()
                    d_str = sel_d_leave.strftime("%d/%m/%Y")
                    st.session_state.manual_leave[(ma_l, d_str)] = True
                    st.success(f"✅ Đã thêm: {ma_l} ({d_str})")
            if st.session_state.manual_leave:
                st.caption(f"📌 Đã đăng ký: {len(st.session_state.manual_leave)} ngày nghỉ")
                for (ma_item, d_item) in list(st.session_state.manual_leave.keys()):
                    st.markdown(f"<div style='font-size:12px; color:#1E293B; padding:2px 0;'>▫️ <b>{ma_item}</b> : {d_item}</div>", unsafe_allow_html=True)
                if st.button("🗑️ Xóa danh sách nghỉ" if st.session_state.lang=='vi' else "🗑️ 休暇リストを削除", key="btn_clr_l", use_container_width=True):
                    st.session_state.confirm_clr_l_side = True
                if st.session_state.get('confirm_clr_l_side', False):
                    st.warning("Xác nhận xóa?" if st.session_state.lang=='vi' else "削除を確認?")
                    cy, cn = st.columns(2)
                    if cy.button("✔️ Có" if st.session_state.lang=='vi' else "✔️ はい", key="btn_clr_l_side_y"):
                        st.session_state.manual_leave = {}
                        st.session_state.confirm_clr_l_side = False
                        st.rerun()
                    if cn.button("❌ Hủy" if st.session_state.lang=='vi' else "❌ キャンセル", key="btn_clr_l_side_n"):
                        st.session_state.confirm_clr_l_side = False
                        st.rerun()

        with tab_ot:
            sel_e_ot = st.selectbox("Nhân viên" if st.session_state.lang=='vi' else "社員", emp_options, key="sb_sel_e_ot")
            sel_d_ot = st.date_input("Ngày OT" if st.session_state.lang=='vi' else "残業日", value=datetime.date.today(), key="sb_sel_d_ot")
            ot_reasons_list = ["Xử lý sự cố khẩn cấp", "Bảo trì công trình", "Chạy thử máy mới", "Họp MOS định kỳ", "Khác"] if st.session_state.lang=='vi' else ["緊急トラブル対応", "設備定期メンテナンス", "新機種テスト運転", "定例MOSミーティング", "その他"]
            sel_r_ot = st.selectbox("Lý do OT" if st.session_state.lang=='vi' else "残業理由", ot_reasons_list, key="sb_sel_r_ot")
            if st.button("➕ Thêm lịch OT" if st.session_state.lang=='vi' else "➕ 残業を追加", type="secondary", use_container_width=True, key="btn_add_ot"):
                if sel_e_ot:
                    ma_o = sel_e_ot.split(" - ")[0].strip().upper()
                    d_str = sel_d_ot.strftime("%d/%m/%Y")
                    st.session_state.manual_ot_reason[(ma_o, d_str)] = sel_r_ot
                    st.success(f"✅ Đã thêm: {ma_o} ({d_str})")
            if st.session_state.manual_ot_reason:
                st.caption(f"📌 Đã đăng ký: {len(st.session_state.manual_ot_reason)} lịch OT")
                for (ma_item, d_item), r_item in list(st.session_state.manual_ot_reason.items()):
                    st.markdown(f"<div style='font-size:12px; color:#1E293B; padding:2px 0;'>▫️ <b>{ma_item}</b> ({d_item})<br><span style='color:#0284C7; font-size:11px;'>└ {r_item}</span></div>", unsafe_allow_html=True)
                if st.button("🗑️ Xóa danh sách OT" if st.session_state.lang=='vi' else "🗑️ 残業リストを削除", key="btn_clr_ot", use_container_width=True):
                    st.session_state.confirm_clr_ot_side = True
                if st.session_state.get('confirm_clr_ot_side', False):
                    st.warning("Xác nhận xóa?" if st.session_state.lang=='vi' else "削除を確認?")
                    cy, cn = st.columns(2)
                    if cy.button("✔️ Có" if st.session_state.lang=='vi' else "✔️ はい", key="btn_clr_ot_side_y"):
                        st.session_state.manual_ot_reason = {}
                        st.session_state.confirm_clr_ot_side = False
                        st.rerun()
                    if cn.button("❌ Hủy" if st.session_state.lang=='vi' else "❌ キャンセル", key="btn_clr_ot_side_n"):
                        st.session_state.confirm_clr_ot_side = False
                        st.rerun()

def render_email_sending_sidebar():
    with st.sidebar.expander("📧 Gửi Phiếu Xác Nhận Chấm Công" if st.session_state.lang == 'vi' else "📧 給与・勤怠明細の送信", expanded=False):
        st.caption("Gửi email thông báo tự động tới từng kỹ sư." if st.session_state.lang == 'vi' else "エンジニアへの明細自動メール送信。")
        mode_opts = ["🧪 Mô phỏng nhanh", "📨 Gửi thực SMTP"] if st.session_state.lang == 'vi' else ["🧪 デモシミュレーション", "📨 SMTP実送信"]
        mode_mail = st.radio("Chế độ gửi" if st.session_state.lang == 'vi' else "送信モード", mode_opts, horizontal=True, key="sb_mode_mail_global")
        if "SMTP" in mode_mail:
            st.text_input("SMTP Server" if st.session_state.lang == 'vi' else "SMTPサーバー", value="smtp.gmail.com", key="sb_smtp_srv_global")
            st.text_input("Sender Email" if st.session_state.lang == 'vi' else "送信元メール", placeholder="hr@vietmos.com", key="sb_smtp_mail_global")
            st.text_input("App Password" if st.session_state.lang == 'vi' else "アプリパスワード", type="password", key="sb_smtp_pwd_global")
        if st.button("🚀 Khởi chạy Phát Hành Email" if st.session_state.lang == 'vi' else "🚀 メール送信実行", type="primary", use_container_width=True, key="sb_btn_send_mail_global"):
            import time
            msg_busy = "⏳ Đang tổng hợp và phát hành phiếu chấm công..." if st.session_state.lang == 'vi' else "⏳ 明細データを集計して送信中..."
            with st.status(msg_busy):
                time.sleep(0.5)
                if 'df_raw' in st.session_state and st.session_state.df_raw is not None:
                    count_s = len(st.session_state.df_raw)
                    st.write(f"📨 Đã tạo và phát hành phiếu cho `{count_s}` bản ghi chấm công -> *Thành công*")
                else:
                    st.write("📨 Đã khởi chạy mô phỏng phát hành email -> *Thành công*")
            st.success("✅ Đã phát hành phiếu xác nhận tới email các kỹ sư!" if st.session_state.lang == 'vi' else "✅ 全エンジニアへのメール送信が完了しました！")


# ==========================================
# GIAO DIỆN CÁC CHỨC NĂNG (DASHBOARD)
# ==========================================

# Render nút chuyển đổi ngôn ngữ toàn cục trên mọi trang
render_lang_toggle()
t = get_t(st.session_state.get('lang', 'vi'))
st.session_state['cached_t'] = t

def render_global_sidebar_menu():
    is_vi = (st.session_state.get('lang', 'vi') == 'vi')
    with st.sidebar:
        st.markdown("""
        <style>
        html body [data-testid="stSidebar"] .vmos-company-subtitle,
        html body [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] .vmos-company-subtitle {
            color: #0077BE !important;
        }
        /* Modern SaaS Sidebar Menu Styling - Tối giản tinh tế phong cách Nhật Bản */
        section[data-testid="stSidebar"] div[data-testid="stButton"] {
            width: 100% !important;
            display: block !important;
            margin-bottom: 4px !important;
        }
        /* Nút menu mặc định (Inactive / Secondary) - Thiết kế Card hiện đại, sắc nét */
        section[data-testid="stSidebar"] div[data-testid="stButton"] > button[kind="secondary"] {
            width: 100% !important;
            min-width: 100% !important;
            display: flex !important;
            justify-content: flex-start !important;
            align-items: center !important;
            text-align: left !important;
            padding: 10px 16px !important;
            border-radius: 12px !important;
            border: 1.5px solid var(--border) !important;
            background: var(--bg-card) !important;
            color: var(--text-primary) !important;
            font-weight: 600 !important;
            font-size: 14px !important;
            box-shadow: 0 2px 6px rgba(0, 0, 0, 0.03) !important;
            transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1) !important;
        }
        section[data-testid="stSidebar"] div[data-testid="stButton"] > button[kind="secondary"]:hover {
            background: var(--bg-card-hover) !important;
            color: var(--brand-primary) !important;
            transform: translateX(4px) !important;
            border-color: var(--brand-primary) !important;
            box-shadow: 0 6px 16px rgba(14, 165, 233, 0.18) !important;
        }
        section[data-testid="stSidebar"] div[data-testid="stButton"] > button[kind="secondary"] p {
            color: var(--text-primary) !important;
            font-weight: 600 !important;
            font-size: 14px !important;
            text-align: left !important;
            margin: 0 !important;
            transition: color 0.25s ease !important;
        }
        section[data-testid="stSidebar"] div[data-testid="stButton"] > button[kind="secondary"]:hover p {
            color: var(--brand-primary) !important;
        }

        /* Nút menu đang chọn (Active / Primary) - Siêu đẹp & Sang trọng phong cách Enterprise */
        section[data-testid="stSidebar"] div[data-testid="stButton"] > button[kind="primary"] {
            width: 100% !important;
            min-width: 100% !important;
            display: flex !important;
            justify-content: flex-start !important;
            align-items: center !important;
            text-align: left !important;
            padding: 11px 16px !important;
            border-radius: 12px !important;
            border: 1.5px solid rgba(255, 255, 255, 0.25) !important;
            background: var(--brand-gradient) !important;
            color: white !important;
            font-weight: 800 !important;
            font-size: 14.5px !important;
            box-shadow: var(--shadow-glow), inset 0 1px 2px rgba(255, 255, 255, 0.2) !important;
            border-left: 5px solid var(--accent) !important;
            transition: all 0.3s cubic-bezier(0.16, 1, 0.3, 1) !important;
        }
        section[data-testid="stSidebar"] div[data-testid="stButton"] > button[kind="primary"]:hover {
            transform: translateY(-2px) !important;
            box-shadow: var(--shadow-glow), inset 0 1px 3px rgba(255, 255, 255, 0.3) !important;
            background: var(--accent-gradient) !important;
        }
        section[data-testid="stSidebar"] div[data-testid="stButton"] > button[kind="primary"] p {
            color: white !important;
            font-weight: 800 !important;
            font-size: 14.5px !important;
            text-align: left !important;
            margin: 0 !important;
            text-shadow: 0 1px 2px rgba(0, 0, 0, 0.25) !important;
            letter-spacing: 0.3px !important;
        }

        /* Custom sidebar styles removed to restore native Streamlit behavior */
        </style>
        """, unsafe_allow_html=True)

        logo_b64 = LOGO_HEADER_B64
        if logo_b64:
            st.markdown(f'''
            <div style="display: flex; flex-direction: column; align-items: center; justify-content: center; width: 100%; padding: 14px 0 10px 0; text-align: center;">
                <img src="data:image/png;base64,{logo_b64}" style="height:78px; object-fit:contain; margin-bottom: 6px;">
                <div class="vmos-company-subtitle" style="color:#0077BE !important; font-size:14px !important; font-weight:800 !important; letter-spacing: 0.5px !important; text-align: center !important; width: 100%;">VIET.MOS COMPANY LIMITED</div>
            </div>
            ''', unsafe_allow_html=True)
        else:
            st.markdown('<h2 style="text-align:center; color:#0052CC; margin-bottom:6px;">VIET.MOS</h2>', unsafe_allow_html=True)
            st.markdown('<div class="vmos-company-subtitle" style="text-align:center !important; color:#0077BE !important; font-size:14px !important; font-weight:800 !important; margin-bottom:12px !important; letter-spacing: 0.5px !important;">VIET.MOS COMPANY LIMITED</div>', unsafe_allow_html=True)
        st.divider()

        curr_page = st.session_state.get("app_page", "overview")
        
        def change_page(new_page):
            st.session_state.app_page = new_page

        # --- NHÓM 1: NGHIỆP VỤ HÀNG NGÀY ---
        st.markdown(t("auto_text_app_1"), unsafe_allow_html=True)
        menu_daily = [
            ("overview", t("auto_text_app_2")),
            ("attendance_sheet", t("auto_text_app_4")),
            ("checkin", t("auto_text_app_5")),
            ("leave_ot", t("auto_text_app_6")),
            ("kpi_schedule", t("auto_text_app_7")),
            ("org_chart", t("auto_text_app_8")),
        ]
        for p_key, p_label in menu_daily:
            is_active = (curr_page == p_key)
            btn_type = "primary" if is_active else "secondary"
            st.button(p_label, key=f"menu_btn_{p_key}", type=btn_type, use_container_width=True, on_click=change_page, args=(p_key,) if not is_active else (curr_page,))
        
        # --- NHÓM 2: QUẢN LÝ & BÁO CÁO ---
        st.markdown(t("auto_text_app_9"), unsafe_allow_html=True)
        menu_admin = [
            ("chamcong", t("auto_text_app_10")),
            ("mos", t("auto_text_app_11")),
        ]
        for p_key, p_label in menu_admin:
            is_active = (curr_page == p_key)
            btn_type = "primary" if is_active else "secondary"
            st.button(p_label, key=f"menu_btn_{p_key}", type=btn_type, use_container_width=True, on_click=change_page, args=(p_key,) if not is_active else (curr_page,))
        
        # --- NHÓM 3: VĂN HÓA & PHÁT TRIỂN ---
        st.markdown(t("auto_text_app_81"), unsafe_allow_html=True)
        menu_culture = [
            ("social", t("auto_text_app_82")),
            ("learning", t("auto_text_app_83")),
        ]
        for p_key, p_label in menu_culture:
            is_active = (curr_page == p_key)
            btn_type = "primary" if is_active else "secondary"
            st.button(p_label, key=f"menu_btn_{p_key}", type=btn_type, use_container_width=True, on_click=change_page, args=(p_key,) if not is_active else (curr_page,))

        st.divider()

        import os
        if not os.path.exists("templates"):
            os.makedirs("templates")
        
        template_files = os.listdir("templates")

        @st.cache_data
        def get_default_template_excel():
            df_tpl = pd.DataFrame({
                "Mã NV": ["NV001", "NV002"],
                "Họ tên": ["Nguyễn Văn A", "Trần Thị B"],
                "Phòng ban": ["Kỹ thuật", "Hành chính"],
                "Chức vụ": ["Kỹ sư", "Nhân sự"],
                "Ngày": ["01/06/2026", "01/06/2026"],
                "Giờ vào": ["08:00", "08:15"],
                "Giờ ra": ["17:00", "17:00"],
                "Dự án": ["Bảo trì nhà máy", ""],
                "Nội dung công việc": ["Kiểm tra định kỳ", ""],
                "Ghi chú": ["", "Đi trễ"]
            })
            import io
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_tpl.to_excel(writer, index=False, sheet_name='Data')
            return output.getvalue()
        
        st.markdown("<div style='height: 10px'></div>", unsafe_allow_html=True)
        
        with st.expander(t("auto_text_app_12")):
            if not template_files:
                st.download_button(
                    label=t("auto_text_app_13"),
                    data=get_default_template_excel(),
                    file_name="Form_Mau_Cham_Cong_Default.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="secondary"
                )
            else:
                selected_template = st.selectbox(t("auto_text_app_14"), template_files, label_visibility="collapsed")
                if selected_template:
                    with open(os.path.join("templates", selected_template), "rb") as f:
                        template_data = f.read()
                    st.download_button(
                        label=t("auto_text_app_15"),
                        data=template_data,
                        file_name=selected_template,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        type="secondary"
                    )

        # ----- NÚT CÀI ĐẶT CỐ ĐỊNH Ở ĐÁY SIDEBAR -----
        st.divider()
        curr_page_now = st.session_state.get("app_page", "overview")
        is_settings_active = (curr_page_now == "history")
        st.button(
            t("auto_text_app_16"),
            key="btn_sb_settings_bottom",
            type="primary" if is_settings_active else "secondary",
            use_container_width=True,
            on_click=change_page,
            args=("history",)
        )

def render_pomo_timer():
    import time
    import streamlit.components.v1 as components
    if "pomo_start_time" not in st.session_state:
        st.session_state.pomo_start_time = time.time()
        
    if st.button("🔄", key="btn_pomo_reset_top", help="Đặt lại đồng hồ / 時計リセット"):
        st.session_state.pomo_start_time = time.time()
        st.rerun()
        
    st.markdown("""
    <style>
    .st-key-btn_pomo_reset_top {
        position: fixed !important;
        top: 12px !important;
        right: 825px !important;
        z-index: 999999 !important;
        width: 32px !important;
    }
    .st-key-btn_pomo_reset_top button {
        background: transparent !important;
        backdrop-filter: none !important;
        border: 1.5px solid rgba(14,165,233,0.35) !important;
        border-radius: 50px !important;
        height: 32px !important;
        min-height: 32px !important;
        padding: 0 !important;
        box-shadow: none !important;
        cursor: pointer !important;
        transition: all 0.3s ease !important;
    }
    .st-key-btn_pomo_reset_top button:hover {
        background: rgba(14,165,233,0.1) !important;
    }
    </style>
    """, unsafe_allow_html=True)
    
    is_vi = (st.session_state.get('lang', 'vi') == 'vi')
    pomo_start_ts = st.session_state.pomo_start_time
    label_work = t("auto_text_app_17")
    label_soon = t("auto_text_app_18")
    label_rest = t("auto_text_app_19")
    label_title = t("auto_text_app_20")
    
    timer_html = f"""
<!DOCTYPE html>
<html>
<head>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ background: transparent; font-family: 'Outfit', 'Be Vietnam Pro', sans-serif; overflow: hidden; }}
  #pill {{
    display: flex; align-items: center; justify-content: center; gap: 8px;
    height: 32px;
    border-radius: 50px;
    background: transparent;
    backdrop-filter: none;
    border: 1.5px solid rgba(14,165,233,0.35);
    box-shadow: none;
    cursor: default;
    transition: all 0.2s ease;
    padding: 0 14px;
    width: 100%;
    color: #0284C7;
  }}
  #time-label {{
    font-size: 14px; font-weight: 800; font-family: 'Outfit', monospace; letter-spacing: 0.5px;
  }}
  #status-icon {{ font-size: 15px; line-height: 1; }}
  
  /* State classes */
  .work {{ color: #0284C7; border-color: rgba(14,165,233,0.35); }}
  .soon {{ color: #D97706; border-color: rgba(217,119,6,0.35); background: rgba(254, 243, 199, 0.95) !important; }}
  .rest {{ color: #EF4444; border-color: rgba(239,68,68,0.35); background: rgba(254, 226, 226, 0.95) !important; }}
</style>
</head>
<body>
<div id="pill" class="work" title="{label_work}">
  <span id="status-icon">⏰</span>
  <div id="time-label">20:00</div>
</div>
<script>
const START_TS = {pomo_start_ts};
const CYCLE = 1200; // 20 * 60

// Tự động nổi widget lên góc trên cùng bên phải màn hình
try {{
    const parentIframe = window.frameElement;
    if (parentIframe) {{
        let wrapper = parentIframe.parentElement;
        while(wrapper && !wrapper.classList.contains("element-container")) {{
            wrapper = wrapper.parentElement;
        }}
        if(wrapper) {{
            wrapper.style.position = "fixed";
            wrapper.style.top = "12px";
            wrapper.style.right = "715px"; /* Đặt trước nút chuyển ngôn ngữ */
            wrapper.style.width = "105px";
            wrapper.style.height = "32px";
            wrapper.style.zIndex = "9999999";
        }}
    }}
}} catch (e) {{}}

function tick() {{
  const now = Date.now() / 1000;
  const elapsed = now - START_TS;
  const rem = Math.max(0, CYCLE - (elapsed % CYCLE));
  const mins = Math.floor(rem / 60);
  const secs = Math.floor(rem % 60);
  
  document.getElementById('time-label').textContent =
    String(mins).padStart(2,'0') + ':' + String(secs).padStart(2,'0');
  
  const pill = document.getElementById('pill');
  if (rem <= 60) {{
    pill.className = 'rest';
    pill.title = "{label_rest}";
  }} else if (rem <= 300) {{
    pill.className = 'soon';
    pill.title = "{label_soon}";
  }} else {{
    pill.className = 'work';
    pill.title = "{label_work}";
  }}
}}
tick();
setInterval(tick, 1000);
</script>
</body>
</html>
"""
    components.html(timer_html, height=32, width=105, scrolling=False)

render_global_sidebar_menu()
render_pomo_timer()

# Render chatbot HERE so that all CSS has already been injected before any blocking API calls
ai_chat.render_chatbot()

if st.session_state.app_page == "mos":
    st.markdown("""
    <style>
    div[data-testid="stButton"] {
        display: flex !important;
        
    }
    div[data-testid="stButton"] > button {
        height: 40px !important;
        min-height: 40px !important;
        line-height: 40px !important;
        padding: 0 28px !important;
        border-radius: 14px !important;
        background: linear-gradient(135deg, #0EA5E9 0%, #0284C7 100%) !important;
        box-shadow: 0 4px 14px rgba(14, 165, 233, 0.3) !important;
        transition: transform 0.2s ease, box-shadow 0.2s ease !important;
        color: white !important;
        border: none !important;
        width: auto !important; /* Force width to fit text */
        white-space: nowrap !important;
    }
    div[data-testid="stButton"] > button:hover {
        background: linear-gradient(135deg, #0284C7 0%, #0369A1 100%) !important;
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 18px rgba(14, 165, 233, 0.45) !important;
    }
    .stButton > button p {
        font-size: 16px !important;
        font-weight: 600 !important;
    }
    </style>
    

    """, unsafe_allow_html=True)





    render_mos_page()
    
    st.stop()



from page_history import render_history_page
from page_overview import render_enterprise_dashboard, render_overview_page
from page_leave_ot import render_leave_ot_page
from page_kpi_schedule import render_kpi_schedule_page
if st.session_state.app_page == "overview":
    render_overview_page()
    st.stop()

# =========================================================================
# V.MOS COPILOT (TRỢ LÝ ẢO AI)
# =========================================================================

# =========================================================================
# ATTENDANCE SHEET (BẢNG CÔNG CÁ NHÂN)
# =========================================================================
if st.session_state.app_page == "attendance_sheet":
    try:
        from page_attendance_sheet import render_attendance_sheet_page
        render_attendance_sheet_page()
    except Exception as e:
        st.error(f"Lỗi tải trang Bảng công cá nhân: {str(e)}")
    st.stop()

if st.session_state.app_page == "leave_ot":
    render_leave_ot_page()
    st.stop()

if st.session_state.app_page == "social":
    from page_social import render_social_page
    render_social_page()
    st.stop()

if st.session_state.app_page == "learning":
    from page_learning_feedback import render_learning_feedback_page
    render_learning_feedback_page()
    st.stop()

if st.session_state.app_page == "org_chart":
    is_vi_nf = (st.session_state.get('lang', 'vi') == 'vi')
    st.markdown("<h2 style='color:#0F172A; font-size:28px; font-family:Plus Jakarta Sans, Inter, sans-serif; font-weight:800; margin-bottom: 8px;'>🏢 Sơ đồ Tổ chức & Danh bạ</h2>" if is_vi_nf else "<h2 style='color:#0F172A; font-size:28px; font-family:Plus Jakarta Sans, Inter, sans-serif; font-weight:800; margin-bottom: 8px;'>🏢 組織図と社員名簿</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color:#64748B; font-size: 15px; margin-bottom: 30px;'>Sơ đồ cấu trúc công ty và danh sách toàn bộ nhân sự hiện tại trong hệ thống.</p>" if is_vi_nf else "<p style='color:#64748B; font-size: 15px; margin-bottom: 30px;'>会社の組織図とシステム内の全社員リスト。</p>", unsafe_allow_html=True)

    @st.cache_data(ttl=30, show_spinner=False)
    def _load_org_emps_from_db():
        result = {}
        try:
            import sqlite3 as _sq
            conn = _sq.connect(DB_FILE)
            c = conn.cursor()
            c.execute("SELECT ma_nv, ten_nv, chuc_vu, phong_ban FROM employees")
            for row in c.fetchall():
                result[row[0]] = {"ma": row[0], "ten": row[1], "cv": row[2] or "", "pb": row[3] or ""}
            conn.close()
        except Exception as _e:
            logger.warning(f"Lỗi load org emps: {_e}")
        return result

    all_emps = dict(_load_org_emps_from_db())
    
    for e in st.session_state.get('manual_emps', []):
        all_emps[e['ma']] = e

    sys_emps = get_company_emp_dict(st.session_state.get('lang', 'vi'))
    for ma, ten in sys_emps.items():
        if ma not in all_emps and ma not in st.session_state.get('deleted_emps', []):
            all_emps[ma] = {"ma": ma, "ten": ten, "cv": "", "pb": ""}
    
    specific_people = [
        {"ma": "GD01", "ten": "Otaki Masahide" if is_vi_nf else "大滝 正英", "cv": "Tổng giám đốc" if is_vi_nf else "最高経営責任者 (CEO)", "pb": "Tổng giám đốc"},
        {"ma": "HC01", "ten": "Lê Thanh Phương" if is_vi_nf else "レ・タイン・フォン", "cv": "Hành chính - Kế toán" if is_vi_nf else "総務・会計", "pb": "Bộ phận hành chính - kế toán"},
        {"ma": "CK01", "ten": "Lê Văn Long" if is_vi_nf else "レ・ヴァン・ロン", "cv": "Kỹ sư cơ khí" if is_vi_nf else "機械設計エンジニア", "pb": "Thiết kế cơ khí"}
    ]
    for sp in specific_people:
        found = False
        for m, e in all_emps.items():
            if str(sp['ma']).strip().upper() == str(e.get('ma', m)).strip().upper() or sp['ten'].split()[-1].lower() in e['ten'].lower():
                found = True
                if not is_vi_nf:
                    e['ten'] = sp['ten']
                    e['cv'] = sp['cv']
                break
        if not found:
            all_emps[sp['ma']] = sp

    del_emps = st.session_state.get('deleted_emps', [])
    for d in del_emps:
        if d in all_emps:
            del all_emps[d]

    to_remove_codes = [k for k, v in all_emps.items() if str(k).strip().upper() in ['VM037', 'VM038'] or str(v.get('ma', '')).strip().upper() in ['VM037', 'VM038']]
    for k in to_remove_codes:
        all_emps.pop(k, None)

    depts = {}
    for m, e in all_emps.items():
        name_lower = e.get('ten', '').lower().strip()
        ma_upper = str(e.get('ma', m)).strip().upper()
        
        if not is_vi_nf and e.get('ten'):
            e['ten'] = translate_name(e['ten'], 'ja')
            
        if ma_upper == 'GD01' or 'otaki' in name_lower or 'masahide' in name_lower or '大滝' in name_lower or '正英' in name_lower:
            e['pb'] = 'Tổng giám đốc'
            e['cv'] = 'Tổng giám đốc' if is_vi_nf else '最高経営責任者 (CEO)'
        elif ma_upper in ['HC01', 'VM012'] or 'phương' in name_lower or 'phuong' in name_lower or 'フォン' in name_lower:
            e['pb'] = 'Bộ phận hành chính - kế toán'
            if not e.get('cv') or not is_vi_nf: e['cv'] = 'Hành chính - Kế toán' if is_vi_nf else '総務・会計'
        elif ma_upper in ['CK01', 'VM011'] or 'long' in name_lower or 'ロン' in name_lower:
            e['pb'] = 'Thiết kế cơ khí'
            if not e.get('cv') or not is_vi_nf: e['cv'] = 'Kỹ sư cơ khí' if is_vi_nf else '機械設計エンジニア'
        elif ma_upper == 'VM028' or 'đạo' in name_lower or 'dao' in name_lower or 'ダオ' in name_lower:
            e['pb'] = 'Thiết kế điện'
            if not e.get('cv') or not is_vi_nf:
                e['cv'] = 'Kỹ sư thiết kế điện & Lập trình điều khiển' if is_vi_nf else '電気設計・制御プログラミングエンジニア'
        elif ma_upper in ['VM022', 'VM024'] or 'hưng' in name_lower or 'hung' in name_lower or 'quân' in name_lower or 'quan' in name_lower or 'フン' in name_lower or 'クアン' in name_lower:
            e['pb'] = 'Mô phỏng 3D'
            if not e.get('cv') or not is_vi_nf:
                e['cv'] = 'Kỹ sư Mô phỏng 3D' if is_vi_nf else '3Dシミュレーションエンジニア'
        else:
            if e.get('pb', '') not in ['Tổng giám đốc', 'Bộ phận hành chính - kế toán', 'Thiết kế cơ khí', 'Thiết kế điện', 'Lập trình điều khiển', 'Mô phỏng 3D']:
                e['pb'] = 'Mô phỏng 3D'

        if not is_vi_nf and e.get('cv'):
            cv_l = e['cv'].lower()
            if 'tổng giám đốc' in cv_l: e['cv'] = '最高経営責任者 (CEO)'
            elif 'hành chính' in cv_l or 'kế toán' in cv_l: e['cv'] = '総務・会計'
            elif 'cơ khí' in cv_l: e['cv'] = '機械設計エンジニア'
            elif 'điện' in cv_l or 'lập trình' in cv_l: e['cv'] = '電気設計・制御プログラミングエンジニア'
            elif 'mô phỏng' in cv_l or '3d' in cv_l: e['cv'] = '3Dシミュレーションエンジニア'

        pb = e.get('pb', '').strip()
        if not pb:
            pb = "Chưa phân bổ phòng ban" if is_vi_nf else "部署未割り当て"
        if pb not in depts:
            depts[pb] = []
        depts[pb].append(e)

    dao_list = [e for e in all_emps.values() if str(e.get('ma', '')).strip().upper() == 'VM028' or 'đạo' in e.get('ten', '').lower() or 'dao' in e.get('ten', '').lower() or 'ダオ' in e.get('ten', '').lower()]
    depts['Thiết kế điện'] = list(dao_list)
    depts['Lập trình điều khiển'] = list(dao_list)
    depts['Thiết kế điện điều khiển'] = list(dao_list)

    @st.dialog("🏢 Thông tin Bộ phận / 部署情報", width="large")
    def show_dept_modal_dialog(dept_name, members, is_vi):
        display_dept_name = dept_name
        if not is_vi:
            dept_ja_map = {
                'Tổng giám đốc': '最高経営責任者 (CEO)',
                'Bộ phận kỹ thuật': '技術部',
                'Thiết kế cơ khí': '機械設計部',
                'Thiết kế điện điều khiển': '制御・電気設計部',
                'Thiết kế điện': '電気設計',
                'Lập trình điều khiển': '制御プログラミング',
                'Mô phỏng 3D': '3Dシミュレーション',
                'Bộ phận hành chính - kế toán': '総務・経理部'
            }
            display_dept_name = dept_ja_map.get(dept_name, dept_name)
        st.markdown(f"""
        <div style="background: {T['primary_gradient']}; padding: 22px; border-radius: 16px; color: white; margin-bottom: 20px; box-shadow: 0 8px 25px rgba(0,0,0,0.15);">
            <h2 style="margin: 0; font-size: 26px; font-weight: 800; color: white;">&#128193; {display_dept_name}</h2>
            <p style="margin: 6px 0 0 0; opacity: 0.95; font-size: 14.5px;">{t("auto_text_app_21")}</p>
        </div>
        """, unsafe_allow_html=True)

        is_ceo_dept = ("tổng giám đốc" in dept_name.lower() or "giám đốc" in dept_name.lower() or "otaki" in dept_name.lower())

        if is_ceo_dept:
            img_tag = f'<img src="data:image/png;base64,{OTAKI_B64}" style="width: 145px; height: 145px; border-radius: 50%; object-fit: cover; border: 4px solid #F59E0B; box-shadow: 0 6px 18px rgba(245,158,11,0.25); flex-shrink: 0;">' if OTAKI_B64 else '<span style="font-size: 50px;">👑</span>'
            if is_vi:
                vi_html = """
                <div style="padding: 24px; background: linear-gradient(to right, #FFFBEB, #FEF3C7); border: 1.5px solid #FCD34D; border-radius: 18px; margin-bottom: 24px; box-shadow: 0 4px 15px rgba(245,158,11,0.12);">
                    <div style="display: flex; align-items: center; gap: 24px; margin-bottom: 18px; flex-wrap: wrap;">
                        """ + img_tag + """
                        <div style="flex: 1; min-width: 250px;">
                            <div style="display: flex; align-items: center; gap: 10px; margin-bottom: 10px;">
                                <span style="font-size: 26px;">&#128081;</span>
                                <h3 style="margin: 0; color: #92400E; font-size: 19px; font-weight: 800;">Giám đốc \u5927\u6edd \u6b63\u82f1 (Otaki Masahide) & Triết lý Lãnh đạo</h3>
                            </div>
                            <p style="color: #78350F; font-size: 14.5px; line-height: 1.65; font-weight: 600; margin: 0;">
                                Với hơn 30 năm kinh nghiệm trong ngành sản xuất, thiết kế và chế tạo máy, Giám đốc Otaki chính là nguồn cảm hứng và nguồn tri thức sâu rộng cho thế hệ nhân sự trẻ trong công ty.
                            </p>
                        </div>
                    </div>
                    <div style="padding: 16px 20px; background: white; border-left: 4.5px solid #F59E0B; border-radius: 10px; font-style: italic; color: #334155; font-size: 14.5px; line-height: 1.75; box-shadow: 0 2px 5px rgba(0,0,0,0.03);">
                        &quot;Công việc là một trong những công cụ không thể thiếu để hoạch định cuộc đời. Nhưng để thực sự hạnh phúc, chúng ta cần cả một công việc tốt lẫn những người cộng sự tâm huyết. Bởi lẽ, sự nghiệp vững vàng và những mối quan hệ chất lượng chính là chìa khóa mở ra một cuộc sống tốt đẹp.<br><br>
                        Công việc tốt là công việc tạo ra hứng thú, khiến bạn cảm thấy mình xứng đáng với đồng tiền kiếm được. Vì vậy, phải chăm chỉ làm việc để nâng cao kỹ năng và rèn luyện bản thân. Tại VIET.MOS, sự phát triển và hạnh phúc của đội ngũ nhân sự luôn là ưu tiên hàng đầu. Chúng phát huy tối đa tiềm năng để cùng kiến tạo một cuộc sống tốt đẹp hơn. Tôi xây dựng một môi trường làm việc tích cực và tự do sáng tạo, nơi mỗi cá nhân đều có cơ hội phát triển.&quot;
                    </div>
                </div>
                """
                st.markdown(vi_html, unsafe_allow_html=True)
            else:
                ja_html = """
                <div style="padding: 24px; background: linear-gradient(to right, #FFFBEB, #FEF3C7); border: 1.5px solid #FCD34D; border-radius: 18px; margin-bottom: 24px; box-shadow: 0 4px 15px rgba(245,158,11,0.12);">
                    <div style="display: flex; align-items: center; gap: 24px; margin-bottom: 18px; flex-wrap: wrap;">
                        """ + img_tag + """
                        <div style="flex: 1; min-width: 250px;">
                            <div style="display: flex; align-items: center; gap: 10px; margin-bottom: 10px;">
                                <span style="font-size: 26px;">&#128081;</span>
                                <h3 style="margin: 0; color: #92400E; font-size: 19px; font-weight: 800;">\u5927\u6edd\u793e\u9577\u306e\u3054\u7d39\u4ecb\u3068\u7d4c\u55b6\u7406\u5ff5</h3>
                            </div>
                            <p style="color: #78350F; font-size: 14.5px; line-height: 1.65; font-weight: 600; margin: 0;">
                                \u88fd\u9020\u30fb\u6a5f\u68b0\u8a2d\u8a08\u30fb\u3082\u306e\u3065\u304f\u308a\u696d\u754c\u306b\u304a\u3044\u306630\u5e74\u4ee5\u4e0a\u306e\u8c4a\u5bcc\u306a\u7d4c\u9a13\u3092\u6301\u3064\u5927\u6edd\u793e\u9577\u306f\u3001\u793e\u5185\u306e\u82e5\u3044\u4eba\u6750\u306b\u3068\u3063\u3066\u6df1\u3044\u77e5\u898b\u3068\u30a4\u30f3\u30b9\u30d4\u30ec\u30fc\u30b7\u30e7\u30f3\u306e\u6e90\u3067\u3059\u3002
                            </p>
                        </div>
                    </div>
                    <div style="padding: 16px 20px; background: {T['bg_card']}; border-left: 4.5px solid #F59E0B; border-radius: 10px; font-style: italic; color: {T['text_secondary']}; font-size: 14.5px; line-height: 1.75; box-shadow: 0 2px 5px rgba(0,0,0,0.03);">
                        &quot;\u4ed5\u4e8b\u306f\u4eba\u751f\u3092\u8a2d\u8a08\u3059\u308b\u305f\u3081\u306e\u4e0d\u53ef\u6b20\u306a\u30c4\u30fc\u30eb\u306e\u3072\u3068\u3064\u3067\u3059\u3002\u3057\u304b\u3057\u3001\u771f\u306b\u5e78\u305b\u306b\u306a\u308b\u305f\u3081\u306b\u306f\u3001\u300e\u826f\u3044\u4ed5\u4e8b\u300f\u3068\u300e\u60c5\u71b1\u7684\u306a\u4ef2\u9593\u300f\u306e\u4e21\u65b9\u304c\u5fc5\u8981\u3067\u3059\u3002\u78ba\u56fa\u305f\u308b\u30ad\u30e3\u30ea\u30a2\u3068\u8cea\u306e\u9ad8\u3044\u4eba\u9593\u95a2\u4fc2\u3053\u305d\u304c\u3001\u7d20\u6674\u3089\u3057\u3044\u4eba\u751f\u3092\u5207\u308a\u62d3\u304f\u9375\u3060\u304b\u3089\u3067\u3059\u3002<br><br>
                        \u826f\u3044\u4ed5\u4e8b\u3068\u306f\u3001\u30ef\u30af\u30ef\u30af\u3059\u308b\u8208\u5473\u3092\u751f\u307f\u51fa\u3057\u3001\u7a3c\u3044\u3060\u5bfe\u4fa1\u306b\u3075\u3055\u308f\u3057\u3044\u81ea\u5206\u3060\u3068\u5b9f\u611f\u3067\u304d\u308b\u4ed5\u4e8b\u3067\u3059\u3002\u3060\u304b\u3089\u3053\u305d\u3001\u61f8\u547d\u306b\u50cd\u304d\u3001\u30b9\u30ad\u30eb\u3092\u78e8\u304d\u3001\u81ea\u5206\u81ea\u8eab\u3092\u935b\u3048\u306a\u3051\u308c\u3070\u306a\u308a\u307e\u305b\u3093\u3002VIET.MOS\u3067\u306f\u3001\u793e\u54e1\u30c1\u30fc\u30e0\u306e\u6210\u9577\u3068\u5e78\u798f\u304c\u5e38\u306b\u6700\u512a\u5148\u4e8b\u9805\u3067\u3059\u3002\u79c1\u305f\u3061\u306f\u793e\u54e1\u4e00\u4eba\u3072\u3068\u308a\u304c\u6700\u5927\u9650\u306b\u30dd\u30c6\u30f3\u30b7\u30e3\u30eb\u3092\u767a\u63ee\u3057\u3001\u5171\u306b\u7d20\u6674\u3089\u3057\u3044\u4eba\u751f\u3092\u7bc9\u3051\u308b\u3088\u3046\u5168\u529b\u3067\u30b5\u30dd\u30fc\u30c8\u3057\u307e\u3059\u3002\u79c1\u304c\u76ee\u6307\u3059\u306e\u306f\u3001\u30dd\u30b8\u30c6\u30a3\u30d6\u3067\u81ea\u7531\u306a\u5275\u9020\u6027\u306b\u3042\u3075\u308c\u3001\u8ab0\u3082\u304c\u6210\u9577\u3067\u304d\u308b\u8077\u5834\u74b0\u5883\u3092\u7bc9\u304f\u3053\u3068\u3067\u3059\u3002&quot;
                    </div>
                </div>
                """
                st.markdown(ja_html, unsafe_allow_html=True)
        else:
            is_mech_dept = ("cơ khí" in dept_name.lower())
            if is_mech_dept:
                if is_vi:
                    st.markdown(f"""
                    <div style="padding: 18px 22px; background: {T['bg_card_hover']}; border: 1.5px solid {T['border']}; border-radius: 14px; margin-bottom: 20px;">
                        <div style="font-size: 18px; font-weight: 800; color: {T['text_primary']}; margin-bottom: 6px;">⚙️ Thiết kế Cơ khí</div>
                        <div style="font-size: 14px; color: {T['text_primary']}; line-height: 1.6;">Quy trình thiết kế chuẩn hóa và toàn diện đáp ứng các tiêu chuẩn kỹ thuật khắt khe nhất trong ngành chế tạo máy.</div>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    t1, t2, t3 = st.tabs(["01 | Xuất bản vẽ 2D từ 3D", "02 | Xuất mô hình 3D từ 2D", "03 | Thiết kế cơ khí mô hình 3D"])
                    
                    with t1:
                        st.markdown(f"""
                        <p style="font-style: italic; color: {T['text_secondary']}; font-size: 14px; line-height: 1.6; margin: 10px 0 16px 0; padding: 12px; background: {T['bg_card_hover']}; border-left: 4px solid {T['primary']}; border-radius: 6px;">
                            Bản vẽ 2D đóng vai trò then chốt trong giai đoạn sản xuất và lắp đặt. Đội ngũ kỹ sư của chúng tôi thực hiện xuất bản vẽ 2D chi tiết cho từng bộ phận, làm cơ sở để lắp ráp thành một hệ thống hoàn chỉnh. Quy trình này cho phép phân tích và đánh giá sản phẩm một cách chi tiết nhất dựa trên các tiêu chuẩn nghiêm ngặt:
                        </p>
                        """, unsafe_allow_html=True)
                        c1, c2 = st.columns(2)
                        with c1:
                            st.markdown(f"""
                            <div style="padding: 14px; background: {T['bg_card']}; border: 1px solid #E2E8F0; border-radius: 10px; margin-bottom: 12px;">
                                <div style="font-weight: 700; color: {T['primary']}; font-size: 14.5px; margin-bottom: 6px;">🔹 Kỹ Thuật và Thiết Kế</div>
                                <div style="font-size: 13.5px; color: {T['text_secondary']}; line-height: 1.5;">Xác định chính xác kích thước, tỷ lệ và mối quan hệ hình học giữa các thành phần trong sản phẩm.</div>
                            </div>
                            <div style="padding: 14px; background: {T['bg_card']}; border: 1px solid #E2E8F0; border-radius: 10px; margin-bottom: 12px;">
                                <div style="font-weight: 700; color: {T['primary']}; font-size: 14.5px; margin-bottom: 6px;">🔹 Sản Xuất và Gia Công</div>
                                <div style="font-size: 13.5px; color: {T['text_secondary']}; line-height: 1.5;">Bản vẽ 2D là cơ sở để lập trình gia công CNC (CAM) và hướng dẫn chi tiết quy trình chế tạo.</div>
                            </div>
                            <div style="padding: 14px; background: {T['bg_card']}; border: 1px solid #E2E8F0; border-radius: 10px; margin-bottom: 12px;">
                                <div style="font-weight: 700; color: {T['primary']}; font-size: 14.5px; margin-bottom: 6px;">🔹 Tạo Bản Vẽ Kỹ Thuật</div>
                                <div style="font-size: 13.5px; color: {T['text_secondary']}; line-height: 1.5;">Sử dụng bản vẽ 2D để tạo bản vẽ kỹ thuật chi tiết, bao gồm các kích thước dung sai, thông số vật liệu và chỉ dẫn cụ thể về cách sản phẩm nên được sản xuất.</div>
                            </div>
                            """, unsafe_allow_html=True)
                        with c2:
                            st.markdown(f"""
                            <div style="padding: 14px; background: {T['bg_card']}; border: 1px solid #E2E8F0; border-radius: 10px; margin-bottom: 12px;">
                                <div style="font-weight: 700; color: {T['primary']}; font-size: 14.5px; margin-bottom: 6px;">🔹 Kiểm Tra và Phân Tích</div>
                                <div style="font-size: 13.5px; color: {T['text_secondary']}; line-height: 1.5;">Đánh giá các yếu tố kỹ thuật trọng yếu trước khi chuyển đổi sang mô hình 3D hoặc đưa vào quy trình sản xuất.</div>
                            </div>
                            <div style="padding: 14px; background: {T['bg_card']}; border: 1px solid #E2E8F0; border-radius: 10px; margin-bottom: 12px;">
                                <div style="font-weight: 700; color: {T['primary']}; font-size: 14.5px; margin-bottom: 6px;">🔹 Chia sẻ và Giao Tiếp</div>
                                <div style="font-size: 13.5px; color: {T['text_secondary']}; line-height: 1.5;">Cung cấp thông tin cho các bộ phận liên quan và khách hàng. Bản vẽ giúp hiện thực hóa ý tưởng, giải pháp kỹ thuật và các chi tiết đặc thù của sản phẩm.</div>
                            </div>
                            """, unsafe_allow_html=True)
                    
                    with t2:
                        st.markdown(f"""
                        <p style="font-style: italic; color: {T['text_secondary']}; font-size: 14px; line-height: 1.6; margin: 10px 0 16px 0; padding: 12px; background: {T['bg_card_hover']}; border-left: 4px solid #0EA5E9; border-radius: 6px;">
                            Dựa trên nền tảng các bản vẽ 2D, các kỹ sư sẽ tiến hành dựng mô hình 3D cho từng chi tiết và toàn bộ cụm sản phẩm. Việc này giúp khách hàng có cái nhìn trực quan và chính xác nhất về thiết kế trước khi tiến hành chế tạo thực tế:
                        </p>
                        """, unsafe_allow_html=True)
                        c1, c2 = st.columns(2)
                        with c1:
                            st.markdown(f"""
                            <div style="padding: 14px; background: {T['bg_card']}; border: 1px solid {T['border']}; border-radius: 10px; margin-bottom: 12px;">
                                <div style="font-weight: 700; color: {T['primary']}; font-size: 14.5px; margin-bottom: 6px;">📦 Hiển thị sản phẩm</div>
                                <div style="font-size: 13.5px; color: {T['text_secondary']}; line-height: 1.5;">Mô hình 3D cho phép quan sát sản phẩm từ mọi góc độ, giúp người dùng nắm bắt rõ ràng cấu trúc hình học và diện mạo tổng thể của sản phẩm.</div>
                            </div>
                            <div style="padding: 14px; background: {T['bg_card']}; border: 1px solid {T['border']}; border-radius: 10px; margin-bottom: 12px;">
                                <div style="font-weight: 700; color: {T['primary']}; font-size: 14.5px; margin-bottom: 6px;">🧩 Kiểm tra tính tương tác</div>
                                <div style="font-size: 13.5px; color: {T['text_secondary']}; line-height: 1.5;">Hỗ trợ kiểm tra sự tương tác cơ học giữa các linh kiện trong cụm lắp ráp. Điều này đảm bảo tính tương thích tuyệt đối và khả năng vận hành đồng bộ của toàn hệ thống.</div>
                            </div>
                            <div style="padding: 14px; background: {T['bg_card']}; border: 1px solid {T['border']}; border-radius: 10px; margin-bottom: 12px;">
                                <div style="font-weight: 700; color: {T['primary']}; font-size: 14.5px; margin-bottom: 6px;">☁️ Chia sẻ dữ liệu kỹ thuật</div>
                                <div style="font-size: 13.5px; color: {T['text_secondary']}; line-height: 1.5;">Mô hình 3D là phương thức tối ưu để trao đổi thông tin kỹ thuật giữa các đối tác, đồng nghiệp và khách hàng đảm bảo tính thống nhất và hạn chế sai sót trong quá trình triển khai dự án.</div>
                            </div>
                            """, unsafe_allow_html=True)
                        with c2:
                            st.markdown(f"""
                            <div style="padding: 14px; background: {T['primary']}; color: white; border-radius: 10px; margin-bottom: 12px;">
                                <div style="font-weight: 700; color: {T['bg_card']}; font-size: 14.5px; margin-bottom: 6px;">📊 Phân tích và mô phỏng</div>
                                <div style="font-size: 13.5px; color: {T['bg_card']}; line-height: 1.5;">Cung cấp nền tảng cho các công cụ phân tích kỹ thuật và mô phỏng. Qua đó có thể đánh giá chính xác hiệu suất vận hành và các tính năng của sản phẩm ngay từ giai đoạn thiết kế.</div>
                            </div>
                            <div style="padding: 14px; background: {T['primary']}; color: white; border-radius: 10px; margin-bottom: 12px;">
                                <div style="font-weight: 700; color: {T['bg_card']}; font-size: 14.5px; margin-bottom: 6px;">🎬 Kết xuất hình ảnh & video (Simulator)</div>
                                <div style="font-size: 13.5px; color: {T['bg_card']}; line-height: 1.5;">Tạo ra các hình ảnh phối cảnh chất lượng cao và video mô phỏng vận hành thực tế. Đây là tư liệu quý giá phục vụ cho công tác quảng bá, đào tạo kỹ thuật và giới thiệu sản phẩm tới khách hàng.</div>
                            </div>
                            <div style="padding: 14px; background: {T['primary']}; color: white; border-radius: 10px; margin-bottom: 12px;">
                                <div style="font-weight: 700; color: {T['bg_card']}; font-size: 14.5px; margin-bottom: 6px;">🦾 Sản xuất và gia công</div>
                                <div style="font-size: 13.5px; color: {T['bg_card']}; line-height: 1.5;">Mô hình 3D là cơ sở dữ liệu quan trọng để lập trình gia công trên máy công cụ (CAM) và tối ưu hóa quy trình sản xuất tự động, giúp rút ngắn thời gian và nâng cao độ chính xác của thành phẩm.</div>
                            </div>
                            """, unsafe_allow_html=True)
                    
                    with t3:
                        st.markdown(f"""
                        <p style="font-style: italic; color: {T['text_secondary']}; font-size: 14px; line-height: 1.6; margin: 10px 0 16px 0; padding: 12px; background: {T['bg_card_hover']}; border-left: 4px solid {T['primary']}; border-radius: 6px;">
                            Tại VIET.MOS, chúng tôi tiên phong cung cấp các giải pháp thiết kế cơ khí 3D tối ưu, hiện thực hóa mọi ý tưởng từ bản vẽ đến thực tiễn một cách sáng tạo và chính xác. Trong suốt quá trình phát triển dự án, VIET.MOS luôn đồng hành cùng khách hàng để đưa ra những giải pháp kỹ thuật ưu việt nhất. Việc ứng dụng mô hình 3D giúp nâng cao khả năng tương tác, đảm bảo mọi yêu cầu khắt khe nhất của khách hàng đều được đáp ứng trọn vẹn.
                        </p>
                        """, unsafe_allow_html=True)
                        c1, c2 = st.columns(2)
                        with c1:
                            st.markdown(f"""
                            <div style="padding: 14px; background: {T['bg_card']}; border: 1.5px solid {T['border']}; border-radius: 10px; margin-bottom: 12px;">
                                <div style="font-weight: 700; color: {T['text_primary']}; font-size: 14.5px; margin-bottom: 6px;">💡 Trực quan hóa ý tưởng và thiết kế</div>
                                <div style="font-size: 13.5px; color: {T['text_secondary']}; line-height: 1.5;">Mô hình 3D là công cụ mạnh mẽ giúp khách hàng hình dung rõ nét về sản phẩm tương lai ngay từ giai đoạn sơ khai. Điều này đảm bảo sự thống nhất tuyệt đối giữa ý tưởng thiết kế và mong đợi thực tế của khách hàng.</div>
                            </div>
                            <div style="padding: 14px; background: {T['bg_card']}; border: 1.5px solid {T['border']}; border-radius: 10px; margin-bottom: 12px;">
                                <div style="font-weight: 700; color: {T['text_primary']}; font-size: 14.5px; margin-bottom: 6px;">🏭 Tối ưu hóa hiệu suất sản xuất</div>
                                <div style="font-size: 13.5px; color: {T['text_secondary']}; line-height: 1.5;">Thiết kế 3D là nền tảng cốt lõi để chuẩn hóa quy trình sản xuất và gia công. Giải pháp này giúp tăng cường hiệu quả vận hành, rút ngắn thời gian triển khai và tiết kiệm tối đa chi phí đầu tư.</div>
                            </div>
                            <div style="padding: 14px; background: {T['bg_card']}; border: 1.5px solid {T['border']}; border-radius: 10px; margin-bottom: 12px;">
                                <div style="font-weight: 700; color: {T['text_primary']}; font-size: 14.5px; margin-bottom: 6px;">🔗 Chia sẻ dữ liệu kỹ thuật</div>
                                <div style="font-size: 13.5px; color: {T['text_secondary']}; line-height: 1.5;">Cung cấp nguồn dữ liệu kỹ thuật chính xác, dễ dàng chia sẻ và kết nối giữa các đối tác, nhà sản xuất và các bên liên quan. Sự minh bạch về thông tin giúp toàn bộ dự án được vận hành trơn tru và nhất quán.</div>
                            </div>
                            """, unsafe_allow_html=True)
                        with c2:
                            st.markdown(f"""
                            <div style="padding: 14px; background: {T['bg_card']}; border: 1.5px solid {T['border']}; border-radius: 10px; margin-bottom: 12px;">
                                <div style="font-weight: 700; color: {T['primary']}; font-size: 14.5px; margin-bottom: 6px;">📈 Đánh giá tính năng và hiệu suất</div>
                                <div style="font-size: 13.5px; color: {T['text_secondary']}; line-height: 1.5;">Cho phép phân tích chuyên sâu về công năng và hiệu suất vận hành trước khi tiến hành chế tạo. Khách hàng có thể đánh giá chi tiết về độ bền, khả năng hoạt động và các thông số kỹ thuật then chốt của sản phẩm.</div>
                            </div>
                            <div style="padding: 14px; background: {T['bg_card']}; border: 1.5px solid {T['border']}; border-radius: 10px; margin-bottom: 12px;">
                                <div style="font-weight: 700; color: {T['primary']}; font-size: 14.5px; margin-bottom: 6px;">🛡️ Kiểm soát và Giảm thiểu rủi ro</div>
                                <div style="font-size: 13.5px; color: {T['text_secondary']}; line-height: 1.5;">Thông qua mô hình 3D, các sai sót hoặc xung đột thiết kế tiềm ẩn sẽ được phát hiện và xử lý sớm. Điều này giúp loại bỏ rủi ro trong quá trình sản xuất, tránh lãng phí nguồn lực và những chi phí phát sinh không đáng có.</div>
                            </div>
                            <div style="padding: 14px; background: {T['bg_card']}; border: 1.5px solid {T['border']}; border-radius: 10px; margin-bottom: 12px;">
                                <div style="font-weight: 700; color: {T['primary']}; font-size: 14.5px; margin-bottom: 6px;">🌐 Nâng tầm trải nghiệm Khách hàng</div>
                                <div style="font-size: 13.5px; color: {T['text_secondary']}; line-height: 1.5;">Chúng tôi mang đến những trải nghiệm thực tế ảo (VR) hoặc mô hình tương tác 3D sinh động, giúp khách hàng có cái nhìn đa chiều và cảm nhận chân thực nhất về sản phẩm trước khi hoàn thiện.</div>
                            </div>
                            """, unsafe_allow_html=True)
                else:
                    st.markdown(f"""
                    <div style="padding: 18px 22px; background: {T['bg_card_hover']}; border: 1.5px solid {T['border']}; border-radius: 14px; margin-bottom: 20px;">
                        <div style="font-size: 18px; font-weight: 800; color: {T['text_primary']}; margin-bottom: 6px;">⚙️ \u6a5f\u68b0\u8a2d\u8a08\u30bd\u30ea\u30e5\u30fc\u30b7\u30e7\u30f3</div>
                        <div style="font-size: 14px; color: {T['text_primary']}; line-height: 1.6;">\u3082\u306e\u3065\u304f\u308a\u304a\u3088\u3073\u6a5f\u68b0\u8a2d\u8a08\u30fb\u88fd\u9020\u696d\u754c\u306e\u6700\u3082\u53b3\u683c\u306a\u57fa\u6e96\u306b\u5bfe\u5fdc\u3059\u308b\u6a19\u6e96\u5316\u304a\u3088\u3073\u5305\u62ec\u7684\u306a\u30a8\u30f3\u30b8\u30cb\u30a2\u30ea\u30f3\u30b0\u8a2d\u8a08\u30d7\u30ed\u30bb\u30b9\u3002</div>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    t1, t2, t3 = st.tabs(["01 | 3D\u304b\u30892D\u56f3\u9762\u4f5c\u6210", "02 | 2D\u304b\u30893D\u30e2\u30c7\u30ea\u30f3\u30b0", "03 | 3D\u6a5f\u68b0\u8a2d\u8a08\u30bd\u30ea\u30e5\u30fc\u30b7\u30e7\u30f3"])
                    
                    with t1:
                        st.markdown(f"""
                        <p style="font-style: italic; color: {T['text_secondary']}; font-size: 14px; line-height: 1.6; margin: 10px 0 16px 0; padding: 12px; background: {T['bg_card_hover']}; border-left: 4px solid {T['primary']}; border-radius: 6px;">
                            2D\u56f3\u9762\u306f\u88fd\u9020\u304a\u3088\u3073\u7d44\u7acb\u3066\u5de5\u7a0b\u306b\u304a\u3044\u3066\u6975\u3081\u3066\u91cd\u8981\u306a\u5f79\u5272\u3092\u679c\u305f\u3057\u307e\u3059\u3002\u5f53\u793e\u306e\u30a8\u30f3\u30b8\u30cb\u30a2\u30ea\u30f3\u30b0\u30c1\u30fc\u30e0\u306f\u3001\u5404\u90e8\u54c1\u306e\u8a73\u7d30\u306a2D\u56f3\u9762\u3092\u4f5c\u6210\u3057\u3001\u5b8c\u5168\u306a\u30b7\u30b9\u30c6\u30e0\u3092\u7d44\u307f\u7acb\u3066\u308b\u305f\u3081\u306e\u57fa\u790e\u3092\u7bc9\u304d\u307e\u3059\u3002\u3053\u306e\u30d7\u30ed\u30bb\u30b9\u306b\u3088\u308a\u3001\u53b3\u683c\u306a\u57fa\u6e96\u306b\u57fa\u3065\u304f\u8a73\u7d30\u306a\u88fd\u54c1\u5206\u6790\u3068\u8a55\u4fa1\u304c\u53ef\u80fd\u306b\u306a\u308a\u307e\u3059\u3002
                        </p>
                        """, unsafe_allow_html=True)
                        c1, c2 = st.columns(2)
                        with c1:
                            st.markdown(f"""
                            <div style="padding: 14px; background: {T['bg_card']}; border: 1px solid #E2E8F0; border-radius: 10px; margin-bottom: 12px;">
                                <div style="font-weight: 700; color: {T['primary']}; font-size: 14.5px; margin-bottom: 6px;">🔹 \u6280\u8853\u3068\u8a2d\u8a08</div>
                                <div style="font-size: 13.5px; color: {T['text_secondary']}; line-height: 1.5;">\u88fd\u54c1\u5185\u306e\u5404\u30b3\u30f3\u30dd\u30fc\u30cd\u30f3\u30c8\u9593\u306e\u6b63\u78ba\u306a\u5bf8\u6cd5\u3001\u7e2e\u5c3a\u3001\u5e7e\u4f55\u5b66\u7684\u95a2\u4fc2\u3092\u5b9a\u7fa9\u3057\u307e\u3059\u3002</div>
                            </div>
                            <div style="padding: 14px; background: {T['bg_card']}; border: 1px solid #E2E8F0; border-radius: 10px; margin-bottom: 12px;">
                                <div style="font-weight: 700; color: {T['primary']}; font-size: 14.5px; margin-bottom: 6px;">🔹 \u88fd\u9020\u3068\u52a0\u5de5</div>
                                <div style="font-size: 13.5px; color: {T['text_secondary']}; line-height: 1.5;">2D\u56f3\u9762\u306fCNC\u52a0\u5de5\u30d7\u30ed\u30b0\u30e9\u30e0\uff08CAM\uff09\u306e\u4f5c\u6210\u3084\u3001\u8a73\u7d30\u306a\u88fd\u9020\u624b\u9806\u306e\u30ac\u30a4\u30c9\u3068\u306a\u308a\u307e\u3059\u3002</div>
                            </div>
                            <div style="padding: 14px; background: {T['bg_card']}; border: 1px solid #E2E8F0; border-radius: 10px; margin-bottom: 12px;">
                                <div style="font-weight: 700; color: {T['primary']}; font-size: 14.5px; margin-bottom: 6px;">🔹 \u6280\u8853\u56f3\u9762\u306e\u4f5c\u6210</div>
                                <div style="font-size: 13.5px; color: {T['text_secondary']}; line-height: 1.5;">\u516c\u5dee\u3001\u6750\u6599\u4ed5\u69d8\u3001\u88fd\u9020\u65b9\u6cd5\u306b\u95a2\u3059\u308b\u5177\u4f53\u7684\u306a\u6307\u793a\u3092\u542b\u3080\u8a73\u7d30\u306a\u6280\u8853\u56f3\u9762\u3092\u4f5c\u6210\u3057\u307e\u3059\u3002</div>
                            </div>
                            """, unsafe_allow_html=True)
                        with c2:
                            st.markdown(f"""
                            <div style="padding: 14px; background: {T['bg_card']}; border: 1px solid #E2E8F0; border-radius: 10px; margin-bottom: 12px;">
                                <div style="font-weight: 700; color: {T['primary']}; font-size: 14.5px; margin-bottom: 6px;">🔹 \u691c\u67fb\u3068\u5206\u6790</div>
                                <div style="font-size: 13.5px; color: {T['text_secondary']}; line-height: 1.5;">3D\u30e2\u30c7\u30eb\u3078\u306e\u79fb\u884c\u3084\u88fd\u9020\u30d7\u30ed\u30bb\u30b9\u306b\u5165\u308b\u524d\u306b\u3001\u91cd\u8981\u306a\u6280\u8853\u8981\u7d20\u3092\u8a55\u4fa1\u3057\u307e\u3059\u3002</div>
                            </div>
                            <div style="padding: 14px; background: {T['bg_card']}; border: 1px solid #E2E8F0; border-radius: 10px; margin-bottom: 12px;">
                                <div style="font-weight: 700; color: {T['primary']}; font-size: 14.5px; margin-bottom: 6px;">🔹 \u5171\u6709\u3068\u30b3\u30df\u30e5\u30cb\u30b1\u30fc\u30b7\u30e7\u30f3</div>
                                <div style="font-size: 13.5px; color: {T['text_secondary']}; line-height: 1.5;">\u95a2\u9023\u90e8\u7f72\u3084\u304a\u5ba2\u69d8\u306b\u6b63\u78ba\u306a\u60c5\u5831\u3092\u63d0\u4f9b\u3057\u307e\u3059\u3002\u56f3\u9762\u306f\u30a2\u30a4\u30c7\u30a2\u3001\u6280\u8853\u30bd\u30ea\u30e5\u30fc\u30b7\u30e7\u30f3\u3001\u304a\u3088\u3073\u88fd\u54c1\u7279\u6709\u306e\u8a73\u7d30\u3092\u5177\u73fe\u5316\u3057\u307e\u3059\u3002</div>
                            </div>
                            """, unsafe_allow_html=True)
                    
                    with t2:
                        st.markdown(f"""
                        <p style="font-style: italic; color: {T['text_secondary']}; font-size: 14px; line-height: 1.6; margin: 10px 0 16px 0; padding: 12px; background: {T['bg_card_hover']}; border-left: 4px solid #0EA5E9; border-radius: 6px;">
                            2D\u56f3\u9762\u3092\u30d9\u30fc\u30b9\u306b\u3001\u30a8\u30f3\u30b8\u30cb\u30a2\u306f\u5404\u90e8\u54c1\u304a\u3088\u3073\u30a2\u30bb\u30f3\u30d6\u30ea\u5168\u4f53\u306e3D\u30e2\u30c7\u30eb\u3092\u69cb\u7bc9\u3057\u307e\u3059\u3002\u3053\u308c\u306b\u3088\u308a\u3001\u5b9f\u969b\u306e\u88fd\u9020\u524d\u306b\u8a2d\u8a08\u306b\u95a2\u3059\u308b\u76f4\u611f\u7684\u304b\u3064\u6b63\u78ba\u306a\u8996\u70b9\u3092\u304a\u5ba2\u69d8\u306b\u63d0\u4f9b\u3057\u307e\u3059\u3002
                        </p>
                        """, unsafe_allow_html=True)
                        c1, c2 = st.columns(2)
                        with c1:
                            st.markdown(f"""
                            <div style="padding: 14px; background: {T['bg_card']}; border: 1px solid {T['border']}; border-radius: 10px; margin-bottom: 12px;">
                                <div style="font-weight: 700; color: {T['primary']}; font-size: 14.5px; margin-bottom: 6px;">📦 \u88fd\u54c1\u306e\u53ef\u8996\u5316</div>
                                <div style="font-size: 13.5px; color: {T['text_secondary']}; line-height: 1.5;">\u3042\u3089\u3086\u308b\u89d2\u5ea6\u304b\u3089\u88fd\u54c1\u3092\u89b3\u5bdf\u3067\u304d\u3001\u5e7e\u4f55\u5b66\u7684\u306a\u69cb\u9020\u3068\u5168\u4f53\u7684\u306a\u5916\u89b3\u3092\u660e\u78ba\u306b\u628a\u63e1\u3067\u304d\u307e\u3059\u3002</div>
                            </div>
                            <div style="padding: 14px; background: {T['bg_card']}; border: 1px solid {T['border']}; border-radius: 10px; margin-bottom: 12px;">
                                <div style="font-weight: 700; color: {T['primary']}; font-size: 14.5px; margin-bottom: 6px;">🧩 \u5e72\u6e09\u30fb\u6574\u5408\u6027\u30c1\u30a7\u30c3\u30af</div>
                                <div style="font-size: 13.5px; color: {T['text_secondary']}; line-height: 1.5;">\u30a2\u30bb\u30f3\u30d6\u30ea\u5185\u306e\u90e8\u54c1\u9593\u306e\u6a5f\u68b0\u7684\u306a\u5e72\u6e09\u3084\u76f8\u4e92\u4f5c\u7528\u3092\u30c1\u30a7\u30c3\u30af\u3057\u3001\u5b8c\u5168\u306a\u4e92\u63db\u6027\u3068\u30b7\u30b9\u30c6\u30e0\u5168\u4f53\u306e\u540c\u671f\u904b\u8ee2\u3092\u4fdd\u8a3c\u3057\u307e\u3059\u3002</div>
                            </div>
                            <div style="padding: 14px; background: {T['bg_card']}; border: 1px solid {T['border']}; border-radius: 10px; margin-bottom: 12px;">
                                <div style="font-weight: 700; color: {T['primary']}; font-size: 14.5px; margin-bottom: 6px;">☁️ \u6280\u8853\u30c7\u30fc\u30bf\u306e\u5171\u6709</div>
                                <div style="font-size: 13.5px; color: {T['text_secondary']}; line-height: 1.5;">\u30d1\u30fc\u30c8\u30ca\u30fc\u3001\u540c\u50da\u3001\u304a\u5ba2\u69d8\u9593\u3067\u6280\u8853\u60c5\u5831\u3092\u5171\u6709\u3059\u308b\u305f\u3081\u306e\u6700\u9069\u306a\u624b\u6bb5\u3067\u3042\u308a\u3001\u30d7\u30ed\u30b8\u30a7\u30af\u30c8\u5c55\u958b\u4e2d\u306e\u30a8\u30e9\u30fc\u3092\u6700\u5c0f\u9650\u306b\u6291\u3048\u307e\u3059\u3002</div>
                            </div>
                            """, unsafe_allow_html=True)
                        with c2:
                            st.markdown(f"""
                            <div style="padding: 14px; background: {T['primary']}; color: white; border-radius: 10px; margin-bottom: 12px;">
                                <div style="font-weight: 700; color: {T['bg_card']}; font-size: 14.5px; margin-bottom: 6px;">📊 \u5206\u6790\u3068\u30b7\u30df\u30e5\u30ec\u30fc\u30b7\u30e7\u30f3</div>
                                <div style="font-size: 13.5px; color: {T['bg_card']}; line-height: 1.5;">\u6280\u8853\u5206\u6790\u304a\u3088\u3073\u30b7\u30df\u30e5\u30ec\u30fc\u30b7\u30e7\u30f3\u30c4\u30fc\u30eb\u306e\u57fa\u76e4\u3092\u63d0\u4f9b\u3057\u3001\u8a2d\u8a08\u6bb5\u968e\u304b\u3089\u52d5\u4f5c\u6027\u80fd\u3084\u6a5f\u80fd\u3092\u6b63\u78ba\u306b\u8a55\u4fa1\u3067\u304d\u307e\u3059\u3002</div>
                            </div>
                            <div style="padding: 14px; background: {T['primary']}; color: white; border-radius: 10px; margin-bottom: 12px;">
                                <div style="font-weight: 700; color: {T['bg_card']}; font-size: 14.5px; margin-bottom: 6px;">🎬 \u753b\u50cf\uff06\u52d5\u753b\u30ec\u30f3\u30c0\u30ea\u30f3\u30b0 (Simulator)</div>
                                <div style="font-size: 13.5px; color: {T['bg_card']}; line-height: 1.5;">\u9ad8\u54c1\u8cea\u306a\u30d1\u30fc\u30b9\u753b\u50cf\u3084\u52d5\u4f5c\u30b7\u30df\u30e5\u30ec\u30fc\u30b7\u30e7\u30f3\u52d5\u753b\u3092\u4f5c\u6210\u3057\u307e\u3059\u3002\u3053\u308c\u3089\u306f\u30d7\u30ed\u30e2\u30fc\u30b7\u30e7\u30f3\u3001\u6280\u8853\u6559\u80b2\u3001\u88fd\u54c1\u7d39\u4ecb\u306b\u8cb4\u91cd\u306a\u8cc7\u6599\u3068\u306a\u308a\u307e\u3059\u3002</div>
                            </div>
                            <div style="padding: 14px; background: {T['primary']}; color: white; border-radius: 10px; margin-bottom: 12px;">
                                <div style="font-weight: 700; color: {T['bg_card']}; font-size: 14.5px; margin-bottom: 6px;">🦾 製造と加工</div>
                                <div style="font-size: 13.5px; color: {T['bg_card']}; line-height: 1.5;">工作機械（CAM）での加工プログラミングや自動製造プロセスの最適化のための重要なデータベースであり、時間の短縮と精度の向上に貢献します。</div>
                            </div>
                            """, unsafe_allow_html=True)
                    
                    with t3:
                        st.markdown(f"""
                        <p style="font-style: italic; color: {T['text_secondary']}; font-size: 14px; line-height: 1.6; margin: 10px 0 16px 0; padding: 12px; background: {T['bg_card_hover']}; border-left: 4px solid {T['primary']}; border-radius: 6px;">
                            VIET.MOS\u3067\u306f\u3001\u6700\u9069\u306a3D\u6a5f\u68b0\u8a2d\u8a08\u30bd\u30ea\u30e5\u30fc\u30b7\u30e7\u30f3\u306e\u63d0\u4f9b\u3092\u30ea\u30fc\u30c9\u3057\u3001\u56f3\u9762\u304b\u3089\u5b9f\u8df5\u306b\u81f3\u308b\u3042\u3089\u3086\u308b\u30a2\u30a4\u30c7\u30a2\u3092\u5275\u9020\u7684\u304b\u3064\u6b63\u78ba\u306b\u5b9f\u73fe\u3057\u307e\u3059\u3002\u30d7\u30ed\u30b8\u30a7\u30af\u30c8\u958b\u767a\u306e\u5168\u30d7\u30ed\u30bb\u30b9\u306b\u304a\u3044\u3066\u3001\u5e38\u306b\u304a\u5ba2\u69d8\u3068\u4f34\u8d70\u3057\u3001\u512a\u308c\u305f\u6280\u8853\u30bd\u30ea\u30e5\u30fc\u30b7\u30e7\u30f3\u3092\u63d0\u6848\u3057\u307e\u3059\u30023D\u30e2\u30c7\u30eb\u306e\u6d3b\u7528\u306b\u3088\u308a\u30a4\u30f3\u30bf\u30e9\u30af\u30b7\u30e7\u30f3\u3092\u9ad8\u3081\u3001\u6700\u3082\u53b3\u3057\u3044\u3054\u8981\u671b\u306b\u3082\u5b8c\u5168\u306b\u304a\u5fdc\u3048\u3057\u307e\u3059\u3002
                        </p>
                        """, unsafe_allow_html=True)
                        c1, c2 = st.columns(2)
                        with c1:
                            st.markdown(f"""
                            <div style="padding: 14px; background: {T['bg_card']}; border: 1.5px solid {T['border']}; border-radius: 10px; margin-bottom: 12px;">
                                <div style="font-weight: 700; color: {T['text_primary']}; font-size: 14.5px; margin-bottom: 6px;">💡 \u30a2\u30a4\u30c7\u30a2\u3068\u8a2d\u8a08\u306e\u53ef\u8996\u5316</div>
                                <div style="font-size: 13.5px; color: {T['text_secondary']}; line-height: 1.5;">\u521d\u671f\u6bb5\u968e\u304b\u3089\u5c06\u6765\u306e\u88fd\u54c1\u3092\u9bae\u660e\u306b\u30a4\u30e1\u30fc\u30b8\u3067\u304d\u308b\u5f37\u529b\u306a\u30c4\u30fc\u30eb\u3067\u3042\u308a\u3001\u8a2d\u8a08\u610f\u56f3\u3068\u304a\u5ba2\u69d8\u306e\u671f\u5f85\u3068\u306e\u5b8c\u5168\u306a\u4e00\u81f4\u3092\u4fdd\u8a3c\u3057\u307e\u3059\u3002</div>
                            </div>
                            <div style="padding: 14px; background: {T['bg_card']}; border: 1.5px solid {T['border']}; border-radius: 10px; margin-bottom: 12px;">
                                <div style="font-weight: 700; color: {T['text_primary']}; font-size: 14.5px; margin-bottom: 6px;">🏭 \u88fd\u9020\u52b9\u7387\u306e\u6700\u9069\u5316</div>
                                <div style="font-size: 13.5px; color: {T['text_secondary']}; line-height: 1.5;">\u88fd\u9020\u304a\u3088\u3073\u52a0\u5de5\u30d7\u30ed\u30bb\u30b9\u3092\u6a19\u6e96\u5316\u3059\u308b\u305f\u3081\u306e\u30b3\u30a2\u57fa\u76e4\u3067\u3042\u308a\u3001\u904b\u7528\u52b9\u7387\u306e\u5411\u4e0a\u3068\u6295\u8cc7\u30b3\u30b9\u30c8\u306e\u524a\u6e1b\u3092\u5b9f\u73fe\u3057\u307e\u3059\u3002</div>
                            </div>
                            <div style="padding: 14px; background: {T['bg_card']}; border: 1.5px solid {T['border']}; border-radius: 10px; margin-bottom: 12px;">
                                <div style="font-weight: 700; color: {T['text_primary']}; font-size: 14.5px; margin-bottom: 6px;">🔗 \u6280\u8853\u30c7\u30fc\u30bf\u306e\u5171\u6709</div>
                                <div style="font-size: 13.5px; color: {T['text_secondary']}; line-height: 1.5;">\u6b63\u78ba\u306a\u6280\u8853\u30c7\u30fc\u30bf\u3092\u63d0\u4f9b\u3057\u3001\u95a2\u4fc2\u8005\u9593\u306e\u30b7\u30fc\u30e0\u30ec\u30b9\u306a\u9023\u643a\u3068\u900f\u660e\u6027\u306e\u9ad8\u3044\u30d7\u30ed\u30b8\u30a7\u30af\u30c8\u904b\u55b6\u3092\u5b9f\u73fe\u3057\u307e\u3059\u3002</div>
                            </div>
                            """, unsafe_allow_html=True)
                        with c2:
                            st.markdown(f"""
                            <div style="padding: 14px; background: {T['bg_card']}; border: 1.5px solid {T['border']}; border-radius: 10px; margin-bottom: 12px;">
                                <div style="font-weight: 700; color: {T['primary']}; font-size: 14.5px; margin-bottom: 6px;">📈 \u6a5f\u80fd\u3068\u6027\u80fd\u306e\u8a55\u4fa1</div>
                                <div style="font-size: 13.5px; color: {T['text_secondary']}; line-height: 1.5;">\u88fd\u9020\u524d\u306b\u6a5f\u80fd\u3068\u52d5\u4f5c\u6027\u80fd\u306e\u8a73\u7d30\u306a\u5206\u6790\u3092\u53ef\u80fd\u306b\u3057\u3001\u8010\u4e45\u6027\u3084\u4e3b\u8981\u6280\u8853\u30d1\u30e9\u30e1\u30fc\u30bf\u30fc\u3092\u8a55\u4fa1\u3067\u304d\u307e\u3059\u3002</div>
                            </div>
                            <div style="padding: 14px; background: {T['bg_card']}; border: 1.5px solid {T['border']}; border-radius: 10px; margin-bottom: 12px;">
                                <div style="font-weight: 700; color: {T['primary']}; font-size: 14.5px; margin-bottom: 6px;">🛡️ \u30ea\u30b9\u30af\u7ba1\u7406\u3068\u524a\u6e1b</div>
                                <div style="font-size: 13.5px; color: {T['text_secondary']}; line-height: 1.5;">\u6f5c\u5728\u7684\u306a\u8a2d\u8a08\u4e0a\u306e\u30df\u30b9\u3084\u5e72\u6e09\u3092\u65e9\u671f\u306b\u767a\u898b\u30fb\u5bfe\u51e6\u3057\u3001\u624b\u623b\u308a\u3084\u7121\u99c4\u306a\u30b3\u30b9\u30c8\u3092\u6392\u9664\u3057\u307e\u3059\u3002</div>
                            </div>
                            <div style="padding: 14px; background: {T['bg_card']}; border: 1.5px solid {T['border']}; border-radius: 10px; margin-bottom: 12px;">
                                <div style="font-weight: 700; color: {T['primary']}; font-size: 14.5px; margin-bottom: 6px;">🌐 \u9867\u5ba2\u4f53\u9a13\u306e\u5411\u4e0a</div>
                                <div style="font-size: 13.5px; color: {T['text_secondary']}; line-height: 1.5;">VR\u3084\u30a4\u30f3\u30bf\u30e9\u30af\u30c6\u30a3\u30d6\u306a3D\u30e2\u30c7\u30eb\u3092\u901a\u3058\u3066\u3001\u5b8c\u6210\u524d\u306b\u88fd\u54c1\u3092\u591a\u89d2\u7684\u306b\u4f53\u611f\u3067\u304d\u308b\u30ea\u30a2\u30eb\u306a\u4f53\u9a13\u3092\u63d0\u4f9b\u3057\u307e\u3059\u3002</div>
                            </div>
                            """, unsafe_allow_html=True)
            elif "điện" in dept_name.lower() or "lập trình điều khiển" in dept_name.lower():
                if is_vi:
                    st.markdown(f"""
                    <div style="padding: 18px 22px; background: {T['bg_card_hover']}; border: 1.5px solid {T['border']}; border-radius: 14px; margin-bottom: 20px;">
                        <div style="font-size: 18px; font-weight: 800; color: {T['text_primary']}; margin-bottom: 6px;">⚡ Thiết kế Lập trình điều khiển</div>
                        <div style="font-size: 14px; color: {T['text_primary']}; line-height: 1.6;">Giải pháp toàn diện về lập trình điều khiển PLC, giao diện HMI/SCADA và tích hợp Robot & Camera Vision cho hệ thống tự động hóa công nghiệp.</div>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    t1, t2, t3 = st.tabs(["01 | Lập trình PLC bằng ngôn ngữ LADDER", "02 | Thiết kế giao diện HMI và SCADA", "03 | Sử dụng ROBOT và CAMERA VISION"])
                    
                    with t1:
                        st.markdown(f"""
                        <p style="font-style: italic; color: {T['text_secondary']}; font-size: 14px; line-height: 1.6; margin: 10px 0 16px 0; padding: 12px; background: {T['bg_card_hover']}; border-left: 4px solid {T['primary']}; border-radius: 6px;">
                            Tại VIET.MOS, chúng tôi không chỉ dừng lại ở lập trình điều khiển PLC mà còn cung cấp giải pháp toàn diện về giao diện HMI (Human-Machine Interface) và hệ thống giám sát, thu thập dữ liệu SCADA (Supervisory Control and Data Acquisition). Quy trình thiết kế của chúng tôi tập trung vào các giá trị cốt lõi:
                        </p>
                        """, unsafe_allow_html=True)
                        st.markdown(f"""
                        <div style="padding: 16px; background: {T['bg_card']}; border: 1px solid #E2E8F0; border-radius: 12px; margin-bottom: 12px; box-shadow: 0 2px 5px rgba(0,0,0,0.02);">
                            <div style="font-weight: 800; color: {T['primary']}; font-size: 15px; margin-bottom: 6px;">🔷 Độ ổn định cao</div>
                            <div style="font-size: 14px; color: {T['text_secondary']}; line-height: 1.6;">Việc ưu tiên sử dụng ngôn ngữ Ladder giúp hệ thống đạt được tính ổn định và độ tin cậy tối đa. Với ưu điểm trực quan, dễ hiểu, giải pháp này giúp tối thiểu hóa rủi ro phát sinh lỗi vận hành và nâng cao khả năng bảo trì định kỳ cho hệ thống.</div>
                        </div>
                        <div style="padding: 16px; background: {T['bg_card']}; border: 1px solid #E2E8F0; border-radius: 12px; margin-bottom: 12px; box-shadow: 0 2px 5px rgba(0,0,0,0.02);">
                            <div style="font-weight: 800; color: {T['primary']}; font-size: 15px; margin-bottom: 6px;">🔷 Hiệu suất tối ưu</div>
                            <div style="font-size: 14px; color: {T['text_secondary']}; line-height: 1.6;">Nhờ am hiểu sâu sắc về phần cứng của nhiều hãng PLC lớn trên thế giới, đội ngũ VIET.MOS có khả năng tối ưu hóa mã nguồn, giúp tăng tốc độ xử lý và hiệu suất điều khiển. Chúng tôi linh hoạt trong việc lựa chọn và tích hợp dòng PLC phù hợp nhất với đặc thù và yêu cầu kỹ thuật của từng dự án.</div>
                        </div>
                        <div style="padding: 16px; background: {T['bg_card']}; border: 1px solid #E2E8F0; border-radius: 12px; margin-bottom: 12px; box-shadow: 0 2px 5px rgba(0,0,0,0.02);">
                            <div style="font-weight: 800; color: {T['primary']}; font-size: 15px; margin-bottom: 6px;">🔷 Linh hoạt chỉnh sửa và mở rộng</div>
                            <div style="font-size: 14px; color: {T['text_secondary']}; line-height: 1.6;">Cấu trúc chương trình mạch lạc không chỉ giúp đơn giản hóa các thuật toán phức tạp mà còn tạo điều kiện thuận lợi cho việc hiệu chỉnh. Điều này cho phép khách hàng dễ dàng nâng cấp hoặc mở rộng quy mô hệ thống tự động hóa theo nhu cầu sản xuất thực tế trong tương lai.</div>
                        </div>
                        """, unsafe_allow_html=True)
                        
                    with t2:
                        st.markdown(f"""
                        <p style="font-style: italic; color: {T['text_secondary']}; font-size: 14px; line-height: 1.6; margin: 10px 0 16px 0; padding: 12px; background: {T['bg_card_hover']}; border-left: 4px solid #0EA5E9; border-radius: 6px;">
                            Đội ngũ kỹ sư tại VIET.MOS sở hữu bề dày kinh nghiệm trong lĩnh vực lập trình PLC, chuyên sử dụng ngôn ngữ Ladder để xây dựng các chương trình điều khiển tối ưu cho thiết bị tự động hóa. Chúng tôi cam kết mang đến những giải pháp kỹ thuật vượt trội, đáp ứng khắt khe các tiêu chuẩn vận hành của khách hàng:
                        </p>
                        """, unsafe_allow_html=True)
                        c1, c2, c3 = st.columns(3)
                        with c1:
                            st.markdown(f"""
                            <div style="padding: 16px; background: {T['bg_card']}; border: 1px solid {T['border']}; border-radius: 12px; height: 100%; box-shadow: 0 2px 6px rgba(0,0,0,0.03);">
                                <div style="font-weight: 800; color: {T['primary']}; font-size: 15px; margin-bottom: 8px;">🖥️ Sử dụng phần mềm chuyên dụng:</div>
                                <div style="font-size: 13.5px; color: {T['text_secondary']}; line-height: 1.55; margin-bottom: 12px;">Chúng tôi khai thác tối đa sức mạnh từ các nền tảng phần mềm chuyên dụng của những thương hiệu uy tín toàn cầu như Mitsubishi, Keyence, Omron, Siemens... Điều này đảm bảo tính tương thích tuyệt đối và tạo ra những giao diện vận hành đạt tiêu chuẩn chất lượng cao nhất.</div>
                                <div style="font-weight: 800; color: {T['primary']}; font-size: 15px; margin-bottom: 8px;">✨ Thiết kế giao diện chuyên nghiệp:</div>
                                <div style="font-size: 13.5px; color: {T['text_secondary']}; line-height: 1.55;">VIET.MOS chú trọng thiết kế giao diện theo hướng chuyên nghiệp và thực tiễn. Hệ thống các nút bấm, biểu đồ và dữ liệu hiển thị được sắp xếp khoa học, trực quan, giúp người vận hành dễ dàng tương tác, giám sát và kiểm soát quy trình sản xuất một cách hiệu quả, chính xác.</div>
                            </div>
                            """, unsafe_allow_html=True)
                        with c2:
                            st.markdown(f"""
                            <div style="padding: 16px; background: {T['bg_card']}; border: 1px solid {T['border']}; border-radius: 12px; height: 100%; box-shadow: 0 2px 6px rgba(0,0,0,0.03);">
                                <div style="font-weight: 800; color: {T['primary']}; font-size: 15px; margin-bottom: 8px;">💡 Ngôn ngữ lập trình đa dạng:</div>
                                <div style="font-size: 13.5px; color: {T['text_secondary']}; line-height: 1.55;">Đội ngũ kỹ sư của chúng tôi còn làm chủ nhiều ngôn ngữ lập trình hiện đại như Python, C#, C++... Sự đa dạng này cho phép tùy biến linh hoạt các tính năng đặc thù, đảm bảo giao diện đáp ứng trọn vẹn mọi yêu cầu kỹ thuật khắt khe và đặc thù của từng hệ thống.</div>
                            </div>
                            """, unsafe_allow_html=True)
                        with c3:
                            st.markdown(f"""
                            <div style="padding: 16px; background: #EFF6FF; border: 1px solid #BFDBFE; border-radius: 12px; height: 100%; box-shadow: 0 2px 6px rgba(0,0,0,0.03);">
                                <div style="font-weight: 800; color: {T['text_primary']}; font-size: 15px; margin-bottom: 8px;">📊 Tích hợp SCADA cho quản lý toàn diện:</div>
                                <div style="font-size: 13.5px; color: {T['text_secondary']}; line-height: 1.55;">Giải pháp SCADA của chúng tôi được thiết kế để quản lý hệ thống một cách tổng thể. Từ việc thu thập, giám sát đến phân tích dữ liệu thời gian thực từ nhiều thiết bị khác nhau, SCADA tạo ra một nền tảng dữ liệu tập trung, cung cấp báo cáo chi tiết về hiệu suất vận hành, giúp doanh nghiệp tối ưu hóa quản lý và ra quyết định chính xác.</div>
                            </div>
                            """, unsafe_allow_html=True)
                            
                    with t3:
                        st.markdown(f"""
                        <p style="font-style: italic; color: {T['text_secondary']}; font-size: 14px; line-height: 1.6; margin: 10px 0 16px 0; padding: 12px; background: {T['bg_card_hover']}; border-left: 4px solid {T['primary']}; border-radius: 6px;">
                            Trong kỷ nguyên công nghiệp 4.0, sự hiện diện của Robot đã thay đổi hoàn toàn phương thức sản xuất truyền thống. Tại VIET.MOS, chúng tôi không chỉ làm chủ kỹ thuật lập trình Robot đơn thuần, mà còn tiên phong trong việc tích hợp hệ thống Camera Vision thông minh nhằm tối ưu hóa hiệu suất và tự động hóa quy trình sản xuất một cách toàn diện.
                        </p>
                        """, unsafe_allow_html=True)
                        st.markdown(f"""
                        <div style="padding: 16px; background: {T['bg_card']}; border: 1px solid #E2E8F0; border-radius: 12px; margin-bottom: 12px; box-shadow: 0 2px 5px rgba(0,0,0,0.02);">
                            <div style="font-weight: 800; color: {T['text_primary']}; font-size: 15px; margin-bottom: 6px;">🤖 Đối tác chiến lược của các hãng Robot hàng đầu</div>
                            <div style="font-size: 14px; color: {T['text_secondary']}; line-height: 1.6;">VIET.MOS tự hào sở hữu bề dày kinh nghiệm trong việc lập trình và triển khai giải pháp cho các dòng Robot từ những thương hiệu uy tín toàn cầu như Yaskawa, Nachi, Fanuc... Sự am hiểu sâu sắc về đặc tính kỹ thuật của từng dòng máy giúp chúng tôi linh hoạt tùy biến và tích hợp Robot vào dây chuyền sản xuất của khách hàng một cách tối ưu nhất.</div>
                        </div>
                        <div style="padding: 16px; background: {T['bg_card']}; border: 1px solid #E2E8F0; border-radius: 12px; margin-bottom: 12px; box-shadow: 0 2px 5px rgba(0,0,0,0.02);">
                            <div style="font-weight: 800; color: {T['text_primary']}; font-size: 15px; margin-bottom: 10px;">👁️ Giải pháp tích hợp Robot & Camera Vision</div>
                            <div style="font-size: 14px; color: {T['text_secondary']}; line-height: 1.6; margin-bottom: 10px;">Việc kết hợp Robot với hệ thống thị giác máy tính (Camera 2D & 3D) là "chìa khóa" để giải quyết các bài toán sản xuất phức tạp. Giải pháp của chúng tôi tập trung vào 3 giá trị cốt lõi:</div>
                            <ul style="margin: 0; padding-left: 20px; color: #334155; font-size: 14px; line-height: 1.7;">
                                <li><b>Kiểm tra chất lượng (Quality Control):</b> Hệ thống Camera tự động nhận diện và loại bỏ sản phẩm lỗi với độ chính xác tuyệt đối và tốc độ xử lý nhanh chóng, giúp doanh nghiệp giảm thiểu rủi ro sai sót.</li>
                                <li><b>Gia công chính xác (Precision Machining):</b> Dựa trên dữ liệu hình ảnh thời gian thực từ Camera, Robot được lập trình để thực hiện các công đoạn gia công với độ chính xác cực cao, đảm bảo tính đồng nhất của sản phẩm.</li>
                                <li><b>Gắp và đặt sản phẩm (Pick & Place):</b> Hệ thống Camera hỗ trợ Robot định vị vị trí và hướng của sản phẩm một cách linh hoạt, giúp quy trình phân loại và đóng gói trở nên tự động hóa hoàn toàn, ngay cả với các vật thể có hình dạng phức tạp.</li>
                            </ul>
                        </div>
                        """, unsafe_allow_html=True)
                else:
                    st.markdown(f"""
                    <div style="padding: 18px 22px; background: {T['bg_card_hover']}; border: 1.5px solid {T['border']}; border-radius: 14px; margin-bottom: 20px;">
                        <div style="font-size: 18px; font-weight: 800; color: {T['text_primary']}; margin-bottom: 6px;">⚡ \u5236\u5fa1\u8a2d\u8a08\u30fb\u30d7\u30ed\u30b0\u30e9\u30df\u30f3\u30b0\u30bd\u30ea\u30e5\u30fc\u30b7\u30e7\u30f3</div>
                        <div style="font-size: 14px; color: {T['text_primary']}; line-height: 1.6;">PLC\u30d7\u30ed\u30b0\u30e9\u30df\u30f3\u30b0\u3001HMI/SCADA\u753b\u9762\u69cb\u7bc9\u3001\u30a4\u30f3\u30c0\u30b9\u30c8\u30ea\u30fc4.0\u81ea\u52d5\u5316\u30b7\u30b9\u30c6\u30e0\u5411\u3051\u30ed\u30dc\u30c3\u30c8\uff06\u30ab\u30e1\u30e9\u30d3\u30b8\u30e7\u30f3\u7d71\u5408\u306e\u5305\u62ec\u7684\u306a\u30bd\u30ea\u30e5\u30fc\u30b7\u30e7\u30f3\u3002</div>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    t1, t2, t3 = st.tabs(["01 | \u30e9\u30c0\u30fc\u8a00\u8a9e\u306b\u3088\u308bPLC\u30d7\u30ed\u30b0\u30e9\u30df\u30f3\u30b0", "02 | HMI\uff06SCADA\u30a4\u30f3\u30bf\u30fc\u30d5\u30a7\u30fc\u30b9\u8a2d\u8a08", "03 | \u7523\u696d\u7528\u30ed\u30dc\u30c3\u30c8\uff06\u30ab\u30e1\u30e9\u30d3\u30b8\u30e7\u30f3\u6d3b\u7528"])
                    
                    with t1:
                        st.markdown(f"""
                        <p style="font-style: italic; color: {T['text_secondary']}; font-size: 14px; line-height: 1.6; margin: 10px 0 16px 0; padding: 12px; background: {T['bg_card_hover']}; border-left: 4px solid {T['primary']}; border-radius: 6px;">
                            VIET.MOS\u3067\u306f\u3001\u5358\u306a\u308bPLC\u5236\u5fa1\u30d7\u30ed\u30b0\u30e9\u30df\u30f3\u30b0\u306b\u3068\u3069\u307e\u3089\u305a\u3001HMI\uff08\u30d2\u30e5\u30fc\u30de\u30f3\u30de\u30b7\u30f3\u30a4\u30f3\u30bf\u30fc\u30d5\u30a7\u30fc\u30b9\uff09\u753b\u9762\u3084\u3001SCADA\uff08\u76e3\u8996\u30fb\u5236\u5fa1\u30fb\u30c7\u30fc\u30bf\u53ce\u96c6\uff09\u30b7\u30b9\u30c6\u30e0\u306b\u95a2\u3059\u308b\u5305\u62ec\u7684\u306a\u30bd\u30ea\u30e5\u30fc\u30b7\u30e7\u30f3\u3092\u63d0\u4f9b\u3057\u307e\u3059\u3002\u5f53\u793e\u306e\u8a2d\u8a08\u30d7\u30ed\u30bb\u30b9\u306f\u4ee5\u4e0b\u306e\u30b3\u30a2\u30d0\u30ea\u30e5\u30fc\u306b\u7126\u70b9\u3092\u5f53\u3066\u3066\u3044\u307e\u3059\uff1a
                        </p>
                        """, unsafe_allow_html=True)
                        st.markdown(f"""
                        <div style="padding: 16px; background: {T['bg_card']}; border: 1px solid #E2E8F0; border-radius: 12px; margin-bottom: 12px; box-shadow: 0 2px 5px rgba(0,0,0,0.02);">
                            <div style="font-weight: 800; color: {T['primary']}; font-size: 15px; margin-bottom: 6px;">🔷 \u9ad8\u3044\u5b89\u5b9a\u6027</div>
                            <div style="font-size: 14px; color: {T['text_secondary']}; line-height: 1.6;">\u30e9\u30c0\u30fc\u8a00\u8a9e\u3092\u512a\u5148\u3057\u3066\u4f7f\u7528\u3059\u308b\u3053\u3068\u3067\u3001\u30b7\u30b9\u30c6\u30e0\u306e\u6700\u5927\u9650\u306e\u5b89\u5b9a\u6027\u3068\u4fe1\u983c\u6027\u3092\u5b9f\u73fe\u3057\u307e\u3059\u3002\u76f4\u611f\u7684\u3067\u7406\u89e3\u3057\u3084\u3059\u3044\u5229\u70b9\u306b\u3088\u308a\u3001\u904b\u7528\u30a8\u30e9\u30fc\u306e\u767a\u751f\u30ea\u30b9\u30af\u3092\u6700\u5c0f\u9650\u306b\u6291\u3048\u3001\u30b7\u30b9\u30c6\u30e0\u306e\u5b9a\u671f\u30e1\u30f3\u30c6\u30ca\u30f3\u30b9\u6027\u3092\u5411\u4e0a\u3055\u305b\u307e\u3059\u3002</div>
                        </div>
                        <div style="padding: 16px; background: {T['bg_card']}; border: 1px solid #E2E8F0; border-radius: 12px; margin-bottom: 12px; box-shadow: 0 2px 5px rgba(0,0,0,0.02);">
                            <div style="font-weight: 800; color: {T['primary']}; font-size: 15px; margin-bottom: 6px;">🔷 \u6700\u9069\u5316\u3055\u308c\u305f\u30d1\u30d5\u30a9\u30fc\u30de\u30f3\u30b9</div>
                            <div style="font-size: 14px; color: {T['text_secondary']}; line-height: 1.6;">\u4e16\u754c\u306e\u5927\u624bPLC\u30e1\u30fc\u30ab\u30fc\u306e\u30cf\u30fc\u30c9\u30a6\u30a7\u30a2\u306b\u95a2\u3059\u308b\u6df1\u3044\u77e5\u898b\u306b\u3088\u308a\u3001VIET.MOS\u30c1\u30fc\u30e0\u306f\u30bd\u30fc\u30b9\u30b3\u30fc\u30c9\u3092\u6700\u9069\u5316\u3057\u3001\u51e6\u7406\u901f\u5ea6\u3068\u5236\u5fa1\u52b9\u7387\u3092\u5411\u4e0a\u3055\u305b\u307e\u3059\u3002\u5404\u30d7\u30ed\u30b8\u30a7\u30af\u30c8\u306e\u7279\u6027\u3084\u6280\u8853\u8981\u6c42\u306b\u6700\u3082\u9069\u3057\u305fPLC\u30e2\u30c7\u30eb\u3092\u67d4\u8edf\u306b\u9078\u629e\u30fb\u7d71\u5408\u3057\u307e\u3059\u3002</div>
                        </div>
                        <div style="padding: 16px; background: {T['bg_card']}; border: 1px solid #E2E8F0; border-radius: 12px; margin-bottom: 12px; box-shadow: 0 2px 5px rgba(0,0,0,0.02);">
                            <div style="font-weight: 800; color: {T['primary']}; font-size: 15px; margin-bottom: 6px;">🔷 \u67d4\u8edf\u306a\u4fee\u6b63\u3068\u62e1\u5f35\u6027</div>
                            <div style="font-size: 14px; color: {T['text_secondary']}; line-height: 1.6;">\u660e\u78ba\u3067\u8ad6\u7406\u7684\u306a\u30d7\u30ed\u30b0\u30e9\u30e0\u69cb\u9020\u306f\u3001\u8907\u96d1\u306a\u30a2\u30eb\u30b4\u30ea\u30ba\u30e0\u3092\u7c21\u7d20\u5316\u3059\u308b\u3060\u3051\u3067\u306a\u304f\u3001\u8abf\u6574\u3084\u4fee\u6b63\u3092\u5bb9\u6613\u306b\u3057\u307e\u3059\u3002\u3053\u308c\u306b\u3088\u308a\u3001\u5c06\u6765\u306e\u5b9f\u969b\u306e\u751f\u7523\u30cb\u30fc\u30ba\u306b\u5408\u308f\u305b\u3066\u3001\u81ea\u52d5\u5316\u30b7\u30b9\u30c6\u30e0\u306e\u30a2\u30c3\u30d7\u30b0\u30ec\u30fc\u30c9\u3084\u898f\u6a21\u62e1\u5f35\u3092\u7c21\u5358\u306b\u884c\u3046\u3053\u3068\u304c\u3067\u304d\u307e\u3059\u3002</div>
                        </div>
                        """, unsafe_allow_html=True)
                        
                    with t2:
                        st.markdown(f"""
                        <p style="font-style: italic; color: {T['text_secondary']}; font-size: 14px; line-height: 1.6; margin: 10px 0 16px 0; padding: 12px; background: {T['bg_card_hover']}; border-left: 4px solid #0EA5E9; border-radius: 6px;">
                            VIET.MOS\u306e\u30a8\u30f3\u30b8\u30cb\u30a2\u30ea\u30f3\u30b0\u30c1\u30fc\u30e0\u306f\u3001PLC\u30d7\u30ed\u30b0\u30e9\u30df\u30f3\u30b0\u5206\u91ce\u306b\u304a\u3051\u308b\u8c4a\u5bcc\u306a\u7d4c\u9a13\u3092\u6709\u3057\u3001\u30e9\u30c0\u30fc\u8a00\u8a9e\u3092\u99c6\u4f7f\u3057\u3066\u81ea\u52d5\u5316\u6a5f\u5668\u306b\u6700\u9069\u5316\u3055\u308c\u305f\u5236\u5fa1\u30d7\u30ed\u30b0\u30e9\u30e0\u3092\u69cb\u7bc9\u3057\u307e\u3059\u3002\u304a\u5ba2\u69d8\u306e\u53b3\u3057\u3044\u904b\u7528\u57fa\u6e96\u3092\u6e80\u305f\u3059\u5353\u8d8a\u3057\u305f\u6280\u8853\u30bd\u30ea\u30e5\u30fc\u30b7\u30e7\u30f3\u3092\u304a\u5c4a\u3051\u3059\u308b\u3053\u3068\u3092\u7d04\u675f\u3057\u307e\u3059\uff1a
                        </p>
                        """, unsafe_allow_html=True)
                        c1, c2, c3 = st.columns(3)
                        with c1:
                            st.markdown(f"""
                            <div style="padding: 16px; background: {T['bg_card']}; border: 1px solid {T['border']}; border-radius: 12px; height: 100%; box-shadow: 0 2px 6px rgba(0,0,0,0.03);">
                                <div style="font-weight: 800; color: {T['primary']}; font-size: 15px; margin-bottom: 8px;">🖥️ \u5c02\u7528\u30bd\u30d5\u30c8\u30a6\u30a7\u30a2\u306e\u6d3b\u7528\uff1a</div>
                                <div style="font-size: 13.5px; color: {T['text_secondary']}; line-height: 1.55; margin-bottom: 12px;">\u4e09\u83f1\u96fb\u6a5f\u3001\u30ad\u30fc\u30a8\u30f3\u30b9\u3001\u30aa\u30e0\u30ed\u30f3\u3001\u30b7\u30fc\u30e1\u30f3\u30b9\u306a\u3069\u3001\u4e16\u754c\u7684\u306b\u4fe1\u983c\u3055\u308c\u308b\u30d6\u30e9\u30f3\u30c9\u306e\u5c02\u7528\u30bd\u30d5\u30c8\u30a6\u30a7\u30a2\u30d7\u30e9\u30c3\u30c8\u30d5\u30a9\u30fc\u30e0\u306e\u80fd\u529b\u3092\u6700\u5927\u9650\u306b\u6d3b\u7528\u3057\u307e\u3059\u3002\u7d76\u5bfe\u7684\u306a\u4e92\u63db\u6027\u3092\u78ba\u4fdd\u3057\u3001\u6700\u9ad8\u54c1\u8cea\u57fa\u6e96\u3092\u6e80\u305f\u3059\u64cd\u4f5c\u753b\u9762\u3092\u4f5c\u6210\u3057\u307e\u3059\u3002</div>
                                <div style="font-weight: 800; color: {T['primary']}; font-size: 15px; margin-bottom: 8px;">✨ \u30d7\u30ed\u30d5\u30a7\u30c3\u30b7\u30e7\u30ca\u30eb\u306a\u753b\u9762\u8a2d\u8a08\uff1a</div>
                                <div style="font-size: 13.5px; color: {T['text_secondary']}; line-height: 1.55;">VIET.MOS\u306f\u30d7\u30ed\u30d5\u30a7\u30c3\u30b7\u30e7\u30ca\u30eb\u3067\u5b9f\u7528\u7684\u306a\u753b\u9762\u8a2d\u8a08\u3092\u91cd\u8996\u3057\u3066\u3044\u307e\u3059\u3002\u30dc\u30bf\u30f3\u3001\u30b0\u30e9\u30d5\u3001\u8868\u793a\u30c7\u30fc\u30bf\u304c\u79d1\u5b66\u7684\u304b\u3064\u76f4\u611f\u7684\u306b\u914d\u7f6e\u3055\u308c\u3001\u30aa\u30da\u30ec\u30fc\u30bf\u30fc\u304c\u52b9\u679c\u7684\u304b\u3064\u6b63\u78ba\u306b\u751f\u7523\u30d7\u30ed\u30bb\u30b9\u3092\u64cd\u4f5c\u3001\u76e3\u8996\u3001\u5236\u5fa1\u3067\u304d\u308b\u3088\u3046\u306b\u3057\u307e\u3059\u3002</div>
                            </div>
                            """, unsafe_allow_html=True)
                        with c2:
                            st.markdown(f"""
                            <div style="padding: 16px; background: {T['bg_card']}; border: 1px solid {T['border']}; border-radius: 12px; height: 100%; box-shadow: 0 2px 6px rgba(0,0,0,0.03);">
                                <div style="font-weight: 800; color: {T['primary']}; font-size: 15px; margin-bottom: 8px;">💡 \u591a\u69d8\u306a\u30d7\u30ed\u30b0\u30e9\u30df\u30f3\u30b0\u8a00\u8a9e\uff1a</div>
                                <div style="font-size: 13.5px; color: {T['text_secondary']}; line-height: 1.55;">\u5f53\u793e\u306e\u30a8\u30f3\u30b8\u30cb\u30a2\u306f\u3001Python\u3001C#\u3001C++\u306a\u3069\u306e\u591a\u69d8\u306a\u73fe\u4ee3\u306e\u30d7\u30ed\u30b0\u30e9\u30df\u30f3\u30b0\u8a00\u8a9e\u306b\u3082\u7cbe\u901a\u3057\u3066\u3044\u307e\u3059\u3002\u3053\u306e\u591a\u69d8\u6027\u306b\u3088\u308a\u3001\u7279\u6709\u306e\u6a5f\u80fd\u3092\u67d4\u8edf\u306b\u30ab\u30b9\u30bf\u30de\u30a4\u30ba\u3067\u304d\u3001\u5404\u30b7\u30b9\u30c6\u30e0\u306e\u53b3\u683c\u3067\u72ec\u81ea\u306e\u6280\u8853\u8981\u6c42\u306b\u5b8c\u5168\u306b\u5bfe\u5fdc\u3057\u305f\u753b\u9762\u3092\u63d0\u4f9b\u3057\u307e\u3059\u3002</div>
                            </div>
                            """, unsafe_allow_html=True)
                        with c3:
                            st.markdown(f"""
                            <div style="padding: 16px; background: #EFF6FF; border: 1px solid #BFDBFE; border-radius: 12px; height: 100%; box-shadow: 0 2px 6px rgba(0,0,0,0.03);">
                                <div style="font-weight: 800; color: {T['text_primary']}; font-size: 15px; margin-bottom: 8px;">📊 \u7d71\u5408SCADA\u306b\u3088\u308b\u7dcf\u5408\u7ba1\u7406\uff1a</div>
                                <div style="font-size: 13.5px; color: {T['text_secondary']}; line-height: 1.55;">\u5f53\u793e\u306eSCADA\u30bd\u30ea\u30e5\u30fc\u30b7\u30e7\u30f3\u306f\u3001\u30b7\u30b9\u30c6\u30e0\u3092\u7dcf\u5408\u7684\u306b\u7ba1\u7406\u3059\u308b\u305f\u3081\u306b\u8a2d\u8a08\u3055\u308c\u3066\u3044\u307e\u3059\u3002\u8907\u6570\u306e\u7570\u306a\u308b\u6a5f\u5668\u304b\u3089\u30ea\u30a2\u30eb\u30bf\u30a4\u30e0\u30c7\u30fc\u30bf\u306e\u53ce\u96c6\u3001\u76e3\u8996\u3001\u5206\u6790\u3092\u884c\u3044\u3001\u96c6\u4e2d\u30c7\u30fc\u30bf\u30d7\u30e9\u30c3\u30c8\u30d5\u30a9\u30fc\u30e0\u3092\u69cb\u7bc9\u3057\u307e\u3059\u3002\u7a3c\u50cd\u30d1\u30d5\u30a9\u30fc\u30de\u30f3\u30b9\u306b\u95a2\u3059\u308b\u8a73\u7d30\u306a\u30ec\u30dd\u30fc\u30c8\u3092\u63d0\u4f9b\u3057\u3001\u4f01\u696d\u306e\u7ba1\u7406\u6700\u9069\u5316\u3068\u6b63\u78ba\u306a\u610f\u601d\u6c7a\u5b9a\u3092\u652f\u63f4\u3057\u307e\u3059\u3002</div>
                            </div>
                            """, unsafe_allow_html=True)
                            
                    with t3:
                        st.markdown(f"""
                        <p style="font-style: italic; color: {T['text_secondary']}; font-size: 14px; line-height: 1.6; margin: 10px 0 16px 0; padding: 12px; background: {T['bg_card_hover']}; border-left: 4px solid {T['primary']}; border-radius: 6px;">
                            \u30a4\u30f3\u30c0\u30b9\u30c8\u30ea\u30fc4.0\u306e\u6642\u4ee3\u306b\u304a\u3044\u3066\u3001\u30ed\u30dc\u30c3\u30c8\u306e\u5b58\u5728\u306f\u5f93\u6765\u306e\u751f\u7523\u65b9\u5f0f\u3092\u5b8c\u5168\u306b\u5909\u9769\u3057\u307e\u3057\u305f\u3002VIET.MOS\u3067\u306f\u3001\u5358\u306a\u308b\u30ed\u30dc\u30c3\u30c8\u30d7\u30ed\u30b0\u30e9\u30df\u30f3\u30b0\u6280\u8853\u306e\u7fd2\u5f97\u306b\u3068\u3069\u307e\u3089\u305a\u3001\u30b9\u30de\u30fc\u30c8\u306a\u30ab\u30e1\u30e9\u30d3\u30b8\u30e7\u30f3\u30b7\u30b9\u30c6\u30e0\u306e\u7d71\u5408\u3092\u5148\u5c0e\u3057\u3001\u751f\u7523\u30d7\u30ed\u30bb\u30b9\u306e\u30d1\u30d5\u30a9\u30fc\u30de\u30f3\u30b9\u6700\u9069\u5316\u3068\u5305\u62ec\u7684\u306a\u81ea\u52d5\u5316\u3092\u5b9f\u73fe\u3057\u307e\u3059\u3002
                        </p>
                        """, unsafe_allow_html=True)
                        st.markdown(f"""
                        <div style="padding: 16px; background: {T['bg_card']}; border: 1px solid #E2E8F0; border-radius: 12px; margin-bottom: 12px; box-shadow: 0 2px 5px rgba(0,0,0,0.02);">
                            <div style="font-weight: 800; color: {T['text_primary']}; font-size: 15px; margin-bottom: 6px;">🤖 \u4e3b\u8981\u30ed\u30dc\u30c3\u30c8\u30e1\u30fc\u30ab\u30fc\u3068\u306e\u6226\u7565\u7684\u30d1\u30fc\u30c8\u30ca\u30fc\u30b7\u30c3\u30d7</div>
                            <div style="font-size: 14px; color: {T['text_secondary']}; line-height: 1.6;">VIET.MOS\u306f\u3001\u5b89\u5ddd\u96fb\u6a5f\u3001NACHI\uff08\u4e0d\u4e8c\u8d8a\uff09\u3001\u30d5\u30a1\u30ca\u30c3\u30af\u306a\u3069\u3001\u4e16\u754c\u7684\u306a\u4fe1\u983c\u3092\u8a87\u308b\u30d6\u30e9\u30f3\u30c9\u306e\u30ed\u30dc\u30c3\u30c8\u306e\u30d7\u30ed\u30b0\u30e9\u30df\u30f3\u30b0\u304a\u3088\u3073\u5c0e\u5165\u306b\u95a2\u3059\u308b\u8c4a\u5bcc\u306a\u7d4c\u9a13\u3092\u6301\u3063\u3066\u3044\u307e\u3059\u3002\u5404\u6a5f\u7a2e\u306e\u6280\u8853\u7279\u6027\u306b\u5bfe\u3059\u308b\u6df1\u3044\u7406\u89e3\u306b\u3088\u308a\u3001\u304a\u5ba2\u69d8\u306e\u751f\u7523\u30e9\u30a4\u30f3\u306b\u30ed\u30dc\u30c3\u30c8\u3092\u6700\u9069\u304b\u3064\u67d4\u8edf\u306b\u30ab\u30b9\u30bf\u30de\u30a4\u30ba\u30fb\u7d71\u5408\u3057\u307e\u3059\u3002</div>
                        </div>
                        <div style="padding: 16px; background: {T['bg_card']}; border: 1px solid #E2E8F0; border-radius: 12px; margin-bottom: 12px; box-shadow: 0 2px 5px rgba(0,0,0,0.02);">
                            <div style="font-weight: 800; color: {T['text_primary']}; font-size: 15px; margin-bottom: 10px;">👁️ \u30ed\u30dc\u30c3\u30c8\uff06\u30ab\u30e1\u30e9\u30d3\u30b8\u30e7\u30f3\u7d71\u5408\u30bd\u30ea\u30e5\u30fc\u30b7\u30e7\u30f3</div>
                            <div style="font-size: 14px; color: {T['text_secondary']}; line-height: 1.6; margin-bottom: 10px;">\u30ed\u30dc\u30c3\u30c8\u3068\u30de\u30b7\u30f3\u30d3\u30b8\u30e7\u30f3\u30b7\u30b9\u30c6\u30e0\uff082D\uff063D\u30ab\u30e1\u30e9\uff09\u306e\u878d\u5408\u306f\u3001\u8907\u96d1\u306a\u751f\u7523\u8ab2\u984c\u3092\u89e3\u6c7a\u3059\u308b\u305f\u3081\u306e\u300c\u9375\u300d\u3067\u3059\u3002\u5f53\u793e\u306e\u30bd\u30ea\u30e5\u30fc\u30b7\u30e7\u30f3\u306f\u4ee5\u4e0b\u306e3\u3064\u306e\u30b3\u30a2\u30d0\u30ea\u30e5\u30fc\u306b\u7126\u70b9\u3092\u5f53\u3066\u3066\u3044\u307e\u3059\uff1a</div>
                            <ul style="margin: 0; padding-left: 20px; color: #334155; font-size: 14px; line-height: 1.7;">
                                <li><b>\u54c1\u8cea\u691c\u67fb\uff08Quality Control\uff09\uff1a</b>\u30ab\u30e1\u30e9\u30b7\u30b9\u30c6\u30e0\u304c\u7d76\u5bfe\u7684\u306a\u7cbe\u5ea6\u3068\u9ad8\u901f\u51e6\u7406\u3067\u4e0d\u826f\u54c1\u3092\u81ea\u52d5\u691c\u51fa\u30fb\u6392\u9664\u3057\u3001\u4f01\u696d\u306e\u30d2\u30e5\u30fc\u30de\u30f3\u30a8\u30e9\u30fc\u3084\u88fd\u9020\u30ea\u30b9\u30af\u3092\u4f4e\u6e1b\u3057\u307e\u3059\u3002</li>
                                <li><b>\u7cbe\u5bc6\u52a0\u5de5\uff08Precision Machining\uff09\uff1a</b>\u30ab\u30e1\u30e9\u304b\u3089\u306e\u30ea\u30a2\u30eb\u30bf\u30a4\u30e0\u753b\u50cf\u30c7\u30fc\u30bf\u306b\u57fa\u3065\u304d\u3001\u30ed\u30dc\u30c3\u30c8\u3092\u30d7\u30ed\u30b0\u30e9\u30df\u30f3\u30b0\u3057\u3066\u6975\u3081\u3066\u9ad8\u3044\u7cbe\u5ea6\u3067\u52a0\u5de5\u5de5\u7a0b\u3092\u5b9f\u884c\u3057\u3001\u88fd\u54c1\u306e\u5747\u4e00\u6027\u3092\u4fdd\u8a3c\u3057\u307e\u3059\u3002</li>
                                <li><b>\u30d4\u30c3\u30ad\u30f3\u30b0\uff06\u30d7\u30ec\u30fc\u30b9\uff08Pick & Place\uff09\uff1a</b>\u30ab\u30e1\u30e9\u30b7\u30b9\u30c6\u30e0\u304c\u88fd\u54c1\u306e\u4f4d\u7f6e\u3068\u65b9\u5411\u306e\u67d4\u8edf\u306a\u8a8d\u8b58\u3092\u30b5\u30dd\u30fc\u30c8\u3057\u3001\u8907\u96d1\u306a\u5f62\u72b6\u306e\u5bfe\u8c61\u7269\u3067\u3042\u3063\u3066\u3082\u3001\u4ed5\u5206\u3051\u304a\u3088\u3073\u68b1\u5305\u30d7\u30ed\u30bb\u30b9\u306e\u5b8c\u5168\u81ea\u52d5\u5316\u3092\u5b9f\u73fe\u3057\u307e\u3059\u3002</li>
                            </ul>
                        </div>
                        """, unsafe_allow_html=True)
            elif "mô phỏng" in dept_name.lower() or "3d" in dept_name.lower():
                if is_vi:
                    st.markdown(f"""
                    <div style="padding: 18px 22px; background: {T['bg_card_hover']}; border: 1.5px solid {T['border']}; border-radius: 14px; margin-bottom: 20px;">
                        <div style="font-size: 18px; font-weight: 800; color: {T['text_primary']}; margin-bottom: 6px;">🧊 Thiết kế Mô phỏng</div>
                        <div style="font-size: 14px; color: {T['text_primary']}; line-height: 1.6;">Ứng dụng phần mềm chuyên dụng Visual Components để mô phỏng chuyển động kỹ thuật số (Animation) và mô phỏng lập trình PLC (PLS) trong môi trường thực tế ảo.</div>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    t1, t2 = st.tabs(["01 | ANIMATION – Mô phỏng chuyển động", "02 | PERFECT LADDER SIMULATION – PLS"])
                    
                    with t1:
                        st.markdown(f"""
                        <p style="font-style: italic; color: {T['text_secondary']}; font-size: 14px; line-height: 1.6; margin: 10px 0 16px 0; padding: 12px; background: {T['bg_card_hover']}; border-left: 4px solid {T['primary']}; border-radius: 6px;">
                            Tại <b>VIET.MOS</b>, chúng tôi ứng dụng phần mềm chuyên dụng <b>Visual Components</b> để thực hiện mô phỏng kỹ thuật số trước khi tiến hành chế tạo thực tế. Giải pháp này mang đến cho khách hàng cái nhìn trực quan sinh động và chuẩn xác nhất về hệ thống trước khi quyết định đầu tư sản xuất.
                        </p>
                        """, unsafe_allow_html=True)
                        st.markdown(f"""
                        <div style="padding: 18px; background: {T['bg_card']}; border: 1.5px dashed {T['primary']}; border-radius: 14px; margin-bottom: 14px; box-shadow: 0 2px 6px rgba(0,0,0,0.02);">
                            <div style="margin-bottom: 12px; display: flex; align-items: flex-start; gap: 10px;">
                                <span style="color: {T['primary']}; font-size: 18px; font-weight: bold;">✔</span>
                                <div><b style="color: {T['text_primary']}; font-size: 14.5px;">Trực quan hóa bản vẽ:</b> <span style="color: {T['text_secondary']}; font-size: 14px;">Cung cấp hình ảnh chuyển động thực tế của toàn bộ hệ thống.</span></div>
                            </div>
                            <div style="margin-bottom: 12px; display: flex; align-items: flex-start; gap: 10px;">
                                <span style="color: {T['primary']}; font-size: 18px; font-weight: bold;">✔</span>
                                <div><b style="color: {T['text_primary']}; font-size: 14.5px;">Tối ưu hóa thiết kế:</b> <span style="color: {T['text_secondary']}; font-size: 14px;">Xác nhận va chạm và kiểm tra chu kỳ sản xuất (Cycle Time) ngay từ giai đoạn thiết kế.</span></div>
                            </div>
                            <div style="margin-bottom: 12px; display: flex; align-items: flex-start; gap: 10px;">
                                <span style="color: {T['primary']}; font-size: 18px; font-weight: bold;">✔</span>
                                <div><b style="color: {T['text_primary']}; font-size: 14.5px;">Quản trị hiệu suất:</b> <span style="color: {T['text_secondary']}; font-size: 14px;">Dự đoán sản lượng thực tế và kiểm tra tính liên kết giữa các công đoạn sản xuất.</span></div>
                            </div>
                            <div style="display: flex; align-items: flex-start; gap: 10px;">
                                <span style="color: {T['primary']}; font-size: 18px; font-weight: bold;">✔</span>
                                <div><b style="color: {T['text_primary']}; font-size: 14.5px;">Trải nghiệm thực tế ảo (VR):</b> <span style="color: {T['text_secondary']}; font-size: 14px;">Ứng dụng công nghệ VR giúp khách hàng tương tác trực tiếp với máy móc trong môi trường ảo, đảm bảo tính khả thi tuyệt đối trước khi triển khai.</span></div>
                            </div>
                        </div>
                        """, unsafe_allow_html=True)
                        
                    with t2:
                        st.markdown(f"""
                        <p style="font-style: italic; color: {T['text_secondary']}; font-size: 14px; line-height: 1.6; margin: 10px 0 16px 0; padding: 12px; background: {T['bg_card_hover']}; border-left: 4px solid {T['primary']}; border-radius: 6px;">
                            Chúng tôi cung cấp dịch vụ mô phỏng PLS nhằm kiểm tra và hoàn thiện chương trình PLC trong môi trường ảo trước khi cài đặt thực tế. Bằng cách kết nối các module I/O ảo từ phần mềm Visual Components với chương trình PLC, đội ngũ kỹ sư có thể phát hiện và xử lý lỗi trong không gian 3D. Quy trình này giúp tối ưu hóa hệ thống, rút ngắn thời gian lắp đặt và đảm bảo chất lượng vận hành cao nhất khi bàn giao cho khách hàng.
                        </p>
                        """, unsafe_allow_html=True)
                        st.markdown(f"""
                        <div style="padding: 18px; background: {T['bg_card']}; border: 1.5px dashed {T['primary']}; border-radius: 14px; margin-bottom: 14px; box-shadow: 0 2px 6px rgba(0,0,0,0.02);">
                            <div style="margin-bottom: 12px; display: flex; align-items: flex-start; gap: 10px;">
                                <span style="color: {T['text_primary']}; font-size: 18px; font-weight: bold;">✔</span>
                                <div><b style="color: {T['text_primary']}; font-size: 14.5px;">Chương trình điều khiển:</b> <span style="color: {T['text_secondary']}; font-size: 14px;">Đảm bảo logic vận hành chính xác theo yêu cầu kỹ thuật.</span></div>
                            </div>
                            <div style="margin-bottom: 12px; display: flex; align-items: flex-start; gap: 10px;">
                                <span style="color: {T['text_primary']}; font-size: 18px; font-weight: bold;">✔</span>
                                <div><b style="color: {T['text_primary']}; font-size: 14.5px;">Hệ thống Interlock:</b> <span style="color: {T['text_secondary']}; font-size: 14px;">Kiểm tra tính an toàn và ràng buộc giữa các cơ cấu bộ phận.</span></div>
                            </div>
                            <div style="margin-bottom: 12px; display: flex; align-items: flex-start; gap: 10px;">
                                <span style="color: {T['text_primary']}; font-size: 18px; font-weight: bold;">✔</span>
                                <div><b style="color: {T['text_primary']}; font-size: 14.5px;">Chế độ vận hành tự động:</b> <span style="color: {T['text_secondary']}; font-size: 14px;">Mô phỏng quá trình chạy tự động để tối ưu năng suất.</span></div>
                            </div>
                            <div style="margin-bottom: 12px; display: flex; align-items: flex-start; gap: 10px;">
                                <span style="color: {T['text_primary']}; font-size: 18px; font-weight: bold;">✔</span>
                                <div><b style="color: {T['text_primary']}; font-size: 14.5px;">Hệ thống cảnh báo (Alarm):</b> <span style="color: {T['text_secondary']}; font-size: 14px;">Kiểm tra các kịch bản báo lỗi và xử lý sự cố.</span></div>
                            </div>
                            <div style="display: flex; align-items: flex-start; gap: 10px;">
                                <span style="color: {T['text_primary']}; font-size: 18px; font-weight: bold;">✔</span>
                                <div><b style="color: {T['text_primary']}; font-size: 14.5px;">Debug hệ thống:</b> <span style="color: {T['text_secondary']}; font-size: 14px;">Hiệu chỉnh và tối ưu hóa mã nguồn toàn diện trong môi trường ảo.</span></div>
                            </div>
                        </div>
                        """, unsafe_allow_html=True)
                else:
                    st.markdown(f"""
                    <div style="padding: 18px 22px; background: {T['bg_card_hover']}; border: 1.5px solid {T['border']}; border-radius: 14px; margin-bottom: 20px;">
                        <div style="font-size: 18px; font-weight: 800; color: {T['text_primary']}; margin-bottom: 6px;">🧊 3D\u30b7\u30df\u30e5\u30ec\u30fc\u30b7\u30e7\u30f3\u8a2d\u8a08\u30bd\u30ea\u30e5\u30fc\u30b7\u30e7\u30f3</div>
                        <div style="font-size: 14px; color: {T['text_primary']}; line-height: 1.6;">\u5c02\u7528\u30bd\u30d5\u30c8\u30a6\u30a7\u30a2\u300cVisual Components\u300d\u3092\u6d3b\u7528\u3057\u3001\u4eee\u60f3\u74b0\u5883\u306b\u304a\u3044\u3066\u30c7\u30b8\u30bf\u30eb\u30e2\u30fc\u30b7\u30e7\u30f3\u30b7\u30df\u30e5\u30ec\u30fc\u30b7\u30e7\u30f3\uff08Animation\uff09\u3084PLC\u30d7\u30ed\u30b0\u30e9\u30df\u30f3\u30b0\u30b7\u30df\u30e5\u30ec\u30fc\u30b7\u30e7\u30f3\uff08PLS\uff09\u3092\u5b9f\u65bd\u3002</div>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    t1, t2 = st.tabs(["01 | ANIMATION – \u30e2\u30fc\u30b7\u30e7\u30f3\u30b7\u30df\u30e5\u30ec\u30fc\u30b7\u30e7\u30f3", "02 | PERFECT LADDER SIMULATION – PLS"])
                    
                    with t1:
                        st.markdown(f"""
                        <p style="font-style: italic; color: {T['text_secondary']}; font-size: 14px; line-height: 1.6; margin: 10px 0 16px 0; padding: 12px; background: {T['bg_card_hover']}; border-left: 4px solid {T['primary']}; border-radius: 6px;">
                            <b>VIET.MOS</b>\u3067\u306f\u3001\u5c02\u7528\u30bd\u30d5\u30c8\u30a6\u30a7\u30a2<b>\u300cVisual Components\u300d</b>\u3092\u6d3b\u7528\u3057\u3001\u5b9f\u969b\u306e\u88fd\u9020\u3092\u884c\u3046\u524d\u306b\u30c7\u30b8\u30bf\u30eb\u30a8\u30f3\u30b8\u30cb\u30a2\u30ea\u30f3\u30b0\u30b7\u30df\u30e5\u30ec\u30fc\u30b7\u30e7\u30f3\u3092\u5b9f\u65bd\u3057\u307e\u3059\u3002\u3053\u306e\u30bd\u30ea\u30e5\u30fc\u30b7\u30e7\u30f3\u306b\u3088\u308a\u3001\u751f\u7523\u6295\u8cc7\u3092\u6c7a\u5b9a\u3059\u308b\u524d\u306b\u3001\u30b7\u30b9\u30c6\u30e0\u306e\u6700\u3082\u30ea\u30a2\u30eb\u3067\u6b63\u78ba\u306a\u8996\u899a\u7684\u30a4\u30e1\u30fc\u30b8\u3092\u304a\u5ba2\u69d8\u306b\u63d0\u4f9b\u3057\u307e\u3059\u3002
                        </p>
                        """, unsafe_allow_html=True)
                        st.markdown(f"""
                        <div style="padding: 18px; background: {T['bg_card']}; border: 1.5px dashed {T['primary']}; border-radius: 14px; margin-bottom: 14px; box-shadow: 0 2px 6px rgba(0,0,0,0.02);">
                            <div style="margin-bottom: 12px; display: flex; align-items: flex-start; gap: 10px;">
                                <span style="color: {T['primary']}; font-size: 18px; font-weight: bold;">✔</span>
                                <div><b style="color: {T['text_primary']}; font-size: 14.5px;">\u56f3\u9762\u306e\u53ef\u8996\u5316\uff1a</b> <span style="color: {T['text_secondary']}; font-size: 14px;">\u30b7\u30b9\u30c6\u30e0\u5168\u4f53\u306e\u30ea\u30a2\u30eb\u306a\u52d5\u4f5c\u30a4\u30e1\u30fc\u30b8\u3092\u63d0\u4f9b\u3057\u307e\u3059\u3002</span></div>
                            </div>
                            <div style="margin-bottom: 12px; display: flex; align-items: flex-start; gap: 10px;">
                                <span style="color: {T['primary']}; font-size: 18px; font-weight: bold;">✔</span>
                                <div><b style="color: {T['text_primary']}; font-size: 14.5px;">\u8a2d\u8a08\u306e\u6700\u9069\u5316\uff1a</b> <span style="color: {T['text_secondary']}; font-size: 14px;">\u8a2d\u8a08\u6bb5\u968e\u304b\u3089\u5e72\u6e09\u78ba\u8a8d\u3084\u751f\u7523\u30bf\u30af\u30c8\u30bf\u30a4\u30e0\uff08Cycle Time\uff09\u306e\u691c\u8a3c\u3092\u884c\u3044\u307e\u3059\u3002</span></div>
                            </div>
                            <div style="margin-bottom: 12px; display: flex; align-items: flex-start; gap: 10px;">
                                <span style="color: {T['primary']}; font-size: 18px; font-weight: bold;">✔</span>
                                <div><b style="color: {T['text_primary']}; font-size: 14.5px;">\u30d1\u30d5\u30a9\u30fc\u30de\u30f3\u30b9\u7ba1\u7406\uff1a</b> <span style="color: {T['text_secondary']}; font-size: 14px;">\u5b9f\u969b\u306e\u751f\u7523\u91cf\u3092\u4e88\u6e2c\u3057\u3001\u5404\u88fd\u9020\u5de5\u7a0b\u9593\u306e\u9023\u643a\u3092\u691c\u8a3c\u3057\u307e\u3059\u3002</span></div>
                            </div>
                            <div style="display: flex; align-items: flex-start; gap: 10px;">
                                <span style="color: {T['primary']}; font-size: 18px; font-weight: bold;">✔</span>
                                <div><b style="color: {T['text_primary']}; font-size: 14.5px;">VR\u4f53\u9a13\uff1a</b> <span style="color: {T['text_secondary']}; font-size: 14px;">VR\u6280\u8853\u3092\u6d3b\u7528\u3057\u3066\u304a\u5ba2\u69d8\u304c\u4eee\u60f3\u7a7a\u9593\u5185\u3067\u6a5f\u68b0\u3068\u76f4\u63a5\u30a4\u30f3\u30bf\u30e9\u30af\u30b7\u30e7\u30f3\u3067\u304d\u308b\u3088\u3046\u306b\u3057\u3001\u5c0e\u5165\u524d\u306e\u5b8c\u5168\u306a\u5b9f\u73fe\u53ef\u80fd\u6027\u3092\u4fdd\u8a3c\u3057\u307e\u3059\u3002</span></div>
                            </div>
                        </div>
                        """, unsafe_allow_html=True)
                        
                    with t2:
                        st.markdown(f"""
                        <p style="font-style: italic; color: {T['text_secondary']}; font-size: 14px; line-height: 1.6; margin: 10px 0 16px 0; padding: 12px; background: {T['bg_card_hover']}; border-left: 4px solid {T['primary']}; border-radius: 6px;">
                            \u5b9f\u6a5f\u3067\u306e\u30bb\u30c3\u30c8\u30a2\u30c3\u30d7\u524d\u306b\u3001\u4eee\u60f3\u74b0\u5883\u4e0a\u3067PLC\u30d7\u30ed\u30b0\u30e9\u30e0\u306e\u691c\u8a3c\u304a\u3088\u3073\u5b8c\u6210\u3092\u56f3\u308bPLS\u30b7\u30df\u30e5\u30ec\u30fc\u30b7\u30e7\u30f3\u30b5\u30fc\u30d3\u30b9\u3092\u63d0\u4f9b\u3057\u307e\u3059\u3002Visual Components\u30bd\u30d5\u30c8\u306e\u4eee\u60f3I/O\u30e2\u30b8\u30e5\u30fc\u30eb\u3068PLC\u30d7\u30ed\u30b0\u30e9\u30e0\u3092\u63a5\u7d9a\u3059\u308b\u3053\u3068\u3067\u3001\u30a8\u30f3\u30b8\u30cb\u30a2\u306f3D\u7a7a\u9593\u5185\u3067\u30d0\u30b0\u3092\u767a\u898b\u30fb\u5bfe\u51e6\u3067\u304d\u307e\u3059\u3002\u3053\u306e\u30d7\u30ed\u30bb\u30b9\u306b\u3088\u308a\u3001\u30b7\u30b9\u30c6\u30e0\u306e\u6700\u9069\u5316\u3001\u73fe\u5730\u7acb\u3061\u4e0a\u3052\u6642\u9593\u306e\u77ed\u7e2e\u3001\u304a\u5ba2\u69d8\u3078\u306e\u5f15\u304d\u6e21\u3057\u6642\u306e\u6700\u9ad8\u54c1\u8cea\u306e\u7a3c\u50cd\u3092\u5b9f\u73fe\u3057\u307e\u3059\u3002
                        </p>
                        """, unsafe_allow_html=True)
                        st.markdown(f"""
                        <div style="padding: 18px; background: {T['bg_card']}; border: 1.5px dashed {T['primary']}; border-radius: 14px; margin-bottom: 14px; box-shadow: 0 2px 6px rgba(0,0,0,0.02);">
                            <div style="margin-bottom: 12px; display: flex; align-items: flex-start; gap: 10px;">
                                <span style="color: {T['text_primary']}; font-size: 18px; font-weight: bold;">✔</span>
                                <div><b style="color: {T['text_primary']}; font-size: 14.5px;">\u5236\u5fa1\u30d7\u30ed\u30b0\u30e9\u30e0\uff1a</b> <span style="color: {T['text_secondary']}; font-size: 14px;">\u6280\u8853\u8981\u4ef6\u306b\u57fa\u3065\u304f\u6b63\u78ba\u306a\u52d5\u4f5c\u30ed\u30b8\u30c3\u30af\u3092\u4fdd\u8a3c\u3057\u307e\u3059\u3002</span></div>
                            </div>
                            <div style="margin-bottom: 12px; display: flex; align-items: flex-start; gap: 10px;">
                                <span style="color: {T['text_primary']}; font-size: 18px; font-weight: bold;">✔</span>
                                <div><b style="color: {T['text_primary']}; font-size: 14.5px;">\u30a4\u30f3\u30bf\u30fc\u30ed\u30c3\u30af\u30b7\u30b9\u30c6\u30e0\uff1a</b> <span style="color: {T['text_secondary']}; font-size: 14px;">\u5404\u6a5f\u69cb\u90e8\u54c1\u9593\u306e\u5b89\u5168\u78ba\u8a8d\u304a\u3088\u3073\u9023\u52d5\u5236\u7d04\u3092\u30c6\u30b9\u30c8\u3057\u307e\u3059\u3002</span></div>
                            </div>
                            <div style="margin-bottom: 12px; display: flex; align-items: flex-start; gap: 10px;">
                                <span style="color: {T['text_primary']}; font-size: 18px; font-weight: bold;">✔</span>
                                <div><b style="color: {T['text_primary']}; font-size: 14.5px;">\u81ea\u52d5\u904b\u8ee2\u30e2\u30fc\u30c9\uff1a</b> <span style="color: {T['text_secondary']}; font-size: 14px;">\u81ea\u52d5\u904b\u8ee2\u30d7\u30ed\u30bb\u30b9\u3092\u30b7\u30df\u30e5\u30ec\u30fc\u30b7\u30e7\u30f3\u3057\u3001\u751f\u7523\u6027\u3092\u6700\u9069\u5316\u3057\u307e\u3059\u3002</span></div>
                            </div>
                            <div style="margin-bottom: 12px; display: flex; align-items: flex-start; gap: 10px;">
                                <span style="color: {T['text_primary']}; font-size: 18px; font-weight: bold;">✔</span>
                                <div><b style="color: {T['text_primary']}; font-size: 14.5px;">\u30a2\u30e9\u30fc\u30e0\u30b7\u30b9\u30c6\u30e0\uff1a</b> <span style="color: {T['text_secondary']}; font-size: 14px;">\u30a8\u30e9\u30fc\u767a\u751f\u30b7\u30ca\u30ea\u30aa\u304a\u3088\u3073\u30c8\u30e9\u30d6\u30eb\u30b7\u30e5\u30fc\u30c6\u30a3\u30f3\u30b0\u624b\u9806\u3092\u691c\u8a3c\u3057\u307e\u3059\u3002</span></div>
                            </div>
                            <div style="display: flex; align-items: flex-start; gap: 10px;">
                                <span style="color: {T['text_primary']}; font-size: 18px; font-weight: bold;">✔</span>
                                <div><b style="color: {T['text_primary']}; font-size: 14.5px;">\u30b7\u30b9\u30c6\u30e0\u30c7\u30d0\u30c3\u30b0\uff1a</b> <span style="color: {T['text_secondary']}; font-size: 14px;">\u4eee\u60f3\u74b0\u5883\u5185\u306b\u304a\u3044\u3066\u30bd\u30fc\u30b9\u30b3\u30fc\u30c9\u3092\u5305\u62ec\u7684\u306b\u8abf\u6574\u30fb\u6700\u9069\u5316\u3057\u307e\u3059\u3002</span></div>
                            </div>
                        </div>
                        """, unsafe_allow_html=True)
            else:
                st.markdown(f"""
                <div style="padding: 15px 18px; background: #F0F9FF; border-left: 4.5px solid #0EA5E9; border-radius: 10px; margin-bottom: 22px; font-size: 14px; color: {T['primary']}; line-height: 1.5;">
                    💡 <b>{t("auto_text_app_22")}</b> {t("auto_text_app_23")}
                </div>
                """, unsafe_allow_html=True)

        # Lọc trùng tên, gộp mã NV
        unique_members = {}
        for m in members:
            ten = m.get('ten', '').strip()
            ma = m.get('ma', '').strip()
            if not ten: continue
            if ten not in unique_members:
                unique_members[ten] = dict(m)
                unique_members[ten]['ma_list'] = [ma] if ma else []
            else:
                if ma and ma not in unique_members[ten]['ma_list']:
                    unique_members[ten]['ma_list'].append(ma)
        
        unique_members_list = list(unique_members.values())

        if dept_name == 'Tổng giám đốc':
            st.markdown(f"<h4 style='color:#0F172A; margin-bottom: 14px; font-size: 17px; font-weight:800;'>{t("auto_text_app_24")}</h4>", unsafe_allow_html=True)
        else:
            st.markdown(f"<h4 style='color:#0F172A; margin-bottom: 14px; font-size: 17px; font-weight:800;'>👥 {t("auto_text_app_25")} ({len(unique_members_list)})</h4>", unsafe_allow_html=True)

        if not unique_members_list:
            st.info(t("auto_text_app_26"))
        else:
            for m in unique_members_list:
                ma = ", ".join(m['ma_list'])
                ten = m.get('ten', '')
                cv = m.get('cv', '') or (t("auto_text_app_27"))
                char_avt = ten[0].upper() if ten else "?"
                if 'GD01' in ma or 'otaki' in ten.lower() or 'masahide' in ten.lower():
                    if OTAKI_B64:
                        avt_html = f'<img src="data:image/png;base64,{OTAKI_B64}" style="width: 48px; height: 48px; border-radius: 50%; object-fit: cover; border: 2px solid #F59E0B; box-shadow: 0 2px 8px rgba(245,158,11,0.25); flex-shrink: 0;">'
                    else:
                        avt_html = f'<div style="width: 48px; height: 48px; border-radius: 50%; background: linear-gradient(135deg, #F59E0B, #D97706); color: white; display: flex; align-items: center; justify-content: center; font-weight: 800; font-size: 18px; flex-shrink: 0;">{char_avt}</div>'
                else:
                    avt_html = f'<div style="width: 48px; height: 48px; border-radius: 50%; background: {T['bg_card']}; color: {T['primary']}; border: 1.5px solid {T['border']}; display: flex; align-items: center; justify-content: center; font-weight: 800; font-size: 18px; flex-shrink: 0; box-shadow: 0 2px 6px rgba(0,0,0,0.05);">{char_avt}</div>'
                ma_str = "" if ('GD01' in ma or 'otaki' in ten.lower() or 'masahide' in ten.lower()) else f' <span style="font-size: 12px; color: #64748B; font-weight: 600; background: {T['bg_card_hover']}; padding: 2px 8px; border-radius: 6px; margin-left: 6px; vertical-align: middle;">[{ma}]</span>'
                st.markdown(f"""<div style="display: flex; align-items: center; gap: 14px; padding: 14px 16px; background: {T['bg_card']}; border: 1.5px solid #E2E8F0; border-radius: 14px; margin-bottom: 12px; box-shadow: 0 2px 6px rgba(0,0,0,0.02); transition: all 0.2s;">
{avt_html}
<div style="overflow: hidden; flex-grow: 1;">
<div style="font-weight: 800; color: {T['text_primary']}; font-size: 15.5px; display: flex; align-items: center; flex-wrap: wrap; gap: 4px; font-family: 'Plus Jakarta Sans', sans-serif;">{ten}{ma_str}</div>
<div style="color: {T['primary']}; font-size: 13.5px; font-weight: 600; margin-top: 4px; font-family: 'Plus Jakarta Sans', sans-serif;">{cv}</div>
</div>
</div>""", unsafe_allow_html=True)


    all_org_depts = [
        "Tổng giám đốc", "Bộ phận kỹ thuật", "Thiết kế cơ khí", 
        "Thiết kế điện điều khiển", "Thiết kế điện", "Lập trình điều khiển", 
        "Mô phỏng 3D", "Bộ phận hành chính - kế toán"
    ]
    clicked_modal_dept = None

    # Put trigger buttons in sidebar hidden permanently with pure CSS so they never show up on screen
    with st.sidebar:
        st.markdown('<div id="secret-modal-triggers-anchor"></div>', unsafe_allow_html=True)
        st.markdown(f"""
        <style>
        div[data-testid="stSidebar"] div:has(#secret-modal-triggers-anchor) ~ div[data-testid="stElementContainer"] {{
            display: none !important; position: absolute !important; opacity: 0 !important; pointer-events: none !important;
        }}
        </style>
        """, unsafe_allow_html=True)
        for d_name in all_org_depts:
            if st.button(f"ORG_BTN_{d_name}", key=f"btn_org_modal_{d_name}"):
                clicked_modal_dept = d_name
                st.session_state.last_clicked_org_dept = d_name

    if clicked_modal_dept and clicked_modal_dept != "Bộ phận kỹ thuật":
        show_dept_modal_dialog(clicked_modal_dept, depts.get(clicked_modal_dept, []), is_vi_nf)

    st.markdown(f"<div style='margin-bottom: 8px;'><b style='color:#0F172A; font-size:15px;'>{'🔍 Tìm kiếm nhanh nhân sự trong hệ thống:' if is_vi_nf else '🔍 \u793e\u54e1\u30af\u30a4\u30c3\u30af\u691c\u7d22:'}</b></div>", unsafe_allow_html=True)
    search_org_query = st.text_input("", placeholder="Nhập tên, mã NV hoặc chức vụ (ví dụ: VM011, Long, Kế toán)..." if is_vi_nf else "\u793e\u54e1ID\u307e\u305f\u306f\u6c0f\u540d\u3092\u5165\u529b (\u4f8b: VM011, \u30ed\u30f3)...", key="org_search_input", label_visibility="collapsed")
    if search_org_query and search_org_query.strip():
        sq = search_org_query.strip().lower()
        matched_emps = [e for e in all_emps.values() if sq in str(e.get('ma','')).lower() or sq in str(e.get('ten','')).lower() or sq in str(e.get('cv','')).lower() or sq in str(e.get('pb','')).lower()]
        st.markdown(f"<div style='padding: 12px 16px; background: #F0F9FF; border: 1.5px solid #0EA5E9; border-radius: 12px; margin: 12px 0;'><b style='color:#0284C7;'>⚡ {'Kết quả tìm kiếm cho' if is_vi_nf else '\u691c\u7d22\u7d50\u679c'} \"{search_org_query}\": {len(matched_emps)} {'nhân sự phù hợp' if is_vi_nf else '\u540d'}</b></div>", unsafe_allow_html=True)
        if matched_emps:
            sc1, sc2 = st.columns(2)
            for m_idx, me in enumerate(matched_emps):
                with (sc1 if m_idx % 2 == 0 else sc2):
                    char_a = me.get('ten', 'N')[0].upper()
                    st.markdown(f"""
                    <div style="display: flex; align-items: center; padding: 14px; background: {T['bg_card']}; border: 1.5px solid #E2E8F0; border-radius: 12px; margin-bottom: 12px; box-shadow: 0 4px 10px rgba(0,0,0,0.03);">
                        <div style="width: 44px; height: 44px; border-radius: 50%; background: linear-gradient(135deg, #0EA5E9, #3B82F6); color: white; display: flex; align-items: center; justify-content: center; font-weight: 800; font-size: 16px; margin-right: 14px; flex-shrink: 0;">{char_a}</div>
                        <div>
                            <div style="font-weight: 800; color: {T['text_primary']}; font-size: 15px;">{me.get('ten')} <span style="color:#64748B; font-size:13px; font-weight:600;">({me.get('ma')})</span></div>
                            <div style="color: {T['primary']}; font-size: 13px; font-weight: 700; margin-top: 2px;">📁 {me.get('pb', 'Chưa phân bổ' if is_vi_nf else '未配属')} | {me.get('cv', 'Nhân viên' if is_vi_nf else '社員')}</div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
        else:
            st.warning("Không tìm thấy nhân sự nào khớp với từ khóa tìm kiếm." if is_vi_nf else "\u8a72\u5f53\u3059\u308b\u793e\u54e1\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093\u3002")

    active_d = st.session_state.get('last_clicked_org_dept', '')
    otaki_mini_img = f'<img src="data:image/png;base64,{OTAKI_B64}" style="width:26px;height:26px;border-radius:50%;object-fit:cover;border:1.5px solid #F59E0B;vertical-align:middle;margin-right:6px;">' if OTAKI_B64 else ''

    lbl_tgd = "Tổng giám đốc" if is_vi_nf else "\u6700\u9ad8\u7d4c\u55b6\u8cac\u4efb\u8005 (CEO)"
    lbl_kt = "Bộ phận kỹ thuật" if is_vi_nf else "\u6280\u8853\u90e8"
    lbl_ck = "Thiết kế cơ khí" if is_vi_nf else "\u6a5f\u68b0\u8a2d\u8a08\u90e8"
    lbl_dk = "Thiết kế điện điều khiển" if is_vi_nf else "\u5236\u5fa1\u30fb\u96fb\u6c17\u8a2d\u8a08\u90e8"
    lbl_dien = "Thiết kế điện" if is_vi_nf else "\u96fb\u6c17\u8a2d\u8a08"
    lbl_lt = "Lập trình điều khiển" if is_vi_nf else "\u5236\u5fa1\u30d7\u30ed\u30b0\u30e9\u30df\u30f3\u30b0"
    lbl_mp = "Mô phỏng 3D" if is_vi_nf else "3D\u30b7\u30df\u30e5\u30ec\u30fc\u30b7\u30e7\u30f3"
    lbl_hc = "Bộ phận hành chính - kế toán" if is_vi_nf else "\u7dcf\u52d9\u30fb\u7d4c\u7406\u90e8"

    emp_kt_list = [e for e in all_emps.values() if e.get('pb') not in ['Tổng giám đốc', 'Bộ phận hành chính - kế toán']]
    count_tgd = len(depts.get('Tổng giám đốc', [])) or 1
    count_kt = len(emp_kt_list) if emp_kt_list else 11
    count_ck = len(depts.get('Thiết kế cơ khí', [])) or 1
    count_dk = len(depts.get('Thiết kế điện điều khiển', [])) or 1
    count_dien = len(depts.get('Thiết kế điện', [])) or 1
    count_lt = len(depts.get('Lập trình điều khiển', [])) or 1
    count_mp = len(depts.get('Mô phỏng 3D', [])) or 4
    count_hc = len(depts.get('Bộ phận hành chính - kế toán', [])) or 2

    sub_tgd = f"{count_tgd} nhân sự" if is_vi_nf else f"{count_tgd} 名"
    sub_kt = f"{count_kt} nhân sự" if is_vi_nf else f"{count_kt} 名"
    sub_ck = f"{count_ck} nhân sự" if is_vi_nf else f"{count_ck} 名"
    sub_dk = f"{count_dk} nhân sự" if is_vi_nf else f"{count_dk} 名"
    sub_dien = f"{count_dien} nhân sự" if is_vi_nf else f"{count_dien} 名"
    sub_lt = f"{count_lt} nhân sự" if is_vi_nf else f"{count_lt} 名"
    sub_mp = f"{count_mp} nhân sự" if is_vi_nf else f"{count_mp} 名"
    sub_hc = f"{count_hc} nhân sự" if is_vi_nf else f"{count_hc} 名"

    list_view_html = ""
    for d_name, d_members in depts.items():
        if not d_members: continue
        d_title = d_name
        if not is_vi_nf:
            d_ja_map = {
                'Tổng giám đốc': '最高経営責任者 (CEO)',
                'Bộ phận kỹ thuật': '技術部',
                'Thiết kế cơ khí': '機械設計部',
                'Thiết kế điện điều khiển': '制御・電気設計部',
                'Thiết kế điện': '電気設計',
                'Lập trình điều khiển': '制御プログラミング',
                'Mô phỏng 3D': '3Dシミュレーション',
                'Bộ phận hành chính - kế toán': '総務・経理部'
            }
            d_title = d_ja_map.get(d_name, d_name)
        if d_name == 'Tổng giám đốc':
            count_span = ""
            icon_span = ""
        else:
            count_span = f'<span style="font-size: 12.5px; background: {T["bg_card_hover"]}; color: {T["primary"]}; padding: 4px 12px; border-radius: 20px; font-weight: 800; border: 1px solid {T["border"]};">{len(d_members)} {"nhân sự" if is_vi_nf else "名"}</span>'
            icon_span = '<span style="display: inline-flex; align-items: center; justify-content: center; width: 26px; height: 26px; border-radius: 7px; background: linear-gradient(135deg, #FEF08A, #FACC15); color: #713F12; font-size: 14px; margin-right: 8px; box-shadow: 0 2px 5px rgba(234,179,8,0.2);">📂</span>'
        list_view_html += f"""<div style="margin-bottom: 22px; background: #F8FAFC; border: 1.5px solid #E2E8F0; border-radius: 14px; padding: 18px; box-shadow: 0 2px 6px rgba(0,0,0,0.02);">
<div style="font-size: 16px; font-weight: 800; color: #0F172A; margin-bottom: 14px; padding-bottom: 8px; border-bottom: 1.5px solid #CBD5E1; display: flex; align-items: center; justify-content: space-between;">
<span style="display: flex; align-items: center;">{icon_span}<span style="font-size: 17px;">{d_title}</span></span>
{count_span}
</div>
<div style="display: grid; grid-template-columns: repeat(auto-fill, minmax(260px, 1fr)); gap: 12px;">"""
        for emp in d_members:
            e_name = emp.get('ten', '')
            e_code = emp.get('ma', '')
            e_cv = emp.get('cv', 'Nhân viên' if is_vi_nf else '社員')
            char_icon = e_name[0].upper() if e_name else 'N'
            if 'GD01' in str(e_code) or 'otaki' in e_name.lower() or 'masahide' in e_name.lower():
                if OTAKI_B64:
                    avt_html = f'<img src="data:image/png;base64,{OTAKI_B64}" style="width: 44px; height: 44px; border-radius: 50%; object-fit: cover; border: 2px solid #F59E0B; box-shadow: 0 2px 6px rgba(245,158,11,0.25); flex-shrink: 0;">'
                else:
                    avt_html = f'<div style="width: 44px; height: 44px; border-radius: 50%; background: linear-gradient(135deg, #F59E0B, #D97706); color: white; font-weight: 800; font-size: 16px; display: flex; align-items: center; justify-content: center; flex-shrink: 0;">{char_icon}</div>'
                code_span = ""
            else:
                avt_html = f'<div style="width: 44px; height: 44px; border-radius: 50%; background: {T["bg_card"]}; color: {T["primary"]}; border: 1.5px solid {T["border"]}; font-weight: 800; font-size: 16px; display: flex; align-items: center; justify-content: center; flex-shrink: 0; box-shadow: 0 2px 5px rgba(0,0,0,0.05);">{char_icon}</div>'
                code_span = f' <span style="font-size: 12px; color: {T["text_secondary"]}; font-weight: 600; background: {T["bg_card_hover"]}; padding: 2px 8px; border-radius: 6px; margin-left: 6px; vertical-align: middle;">[{e_code}]</span>'
            list_view_html += f"""<div style="background: {T['bg_card']}; border: 1.5px solid {T['border']}; border-radius: 14px; padding: 14px 16px; display: flex; align-items: center; gap: 14px; box-shadow: 0 2px 6px rgba(0,0,0,0.02); transition: all 0.2s;">
{avt_html}
<div style="overflow: hidden; flex-grow: 1;">
<div style="font-weight: 800; color: {T['text_primary']}; font-size: 15px; display: flex; align-items: center; flex-wrap: wrap; gap: 4px; font-family: 'Plus Jakarta Sans', sans-serif;">{e_name}{code_span}</div>
<div style="color: {T['primary']}; font-size: 13px; font-weight: 600; margin-top: 4px; font-family: 'Plus Jakarta Sans', sans-serif; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">{e_cv}</div>
</div>
</div>"""
        list_view_html += """</div>
</div>"""

    st.markdown(f"""<style>
.org-tree * {{margin: 0; padding: 0;}}
.org-tree-card {{
    background: {T['bg_card']};
    border: 1.5px solid rgba(14, 165, 233, 0.35);
    border-radius: 16px;
    box-shadow: 0 10px 30px rgba(14, 165, 233, 0.25), 0 4px 12px rgba(15, 23, 42, 0.08);
    padding: 24px 20px 15px 20px;
    margin-top: 20px;
    margin-bottom: 30px;
    position: relative;
}}
.org-tree {{
    display: flex; justify-content: center; overflow-x: auto; padding-bottom: 25px; margin-top: 15px;
}}
.org-tree ul {{
    padding-top: 20px; position: relative;
    transition: all 0.5s; display: flex; justify-content: center;
}}
.org-tree li {{
    float: left; text-align: center; list-style-type: none;
    position: relative; padding: 20px 10px 0 10px; transition: all 0.5s;
}}
.org-tree li::before, .org-tree li::after{{
    content: ''; position: absolute; top: 0; right: 50%;
    border-top: 2px solid #0EA5E9; width: 50%; height: 20px;
}}
.org-tree li::after{{ right: auto; left: 50%; border-left: 2px solid #0EA5E9; }}
.org-tree li:only-child::after, .org-tree li:only-child::before {{ display: none; }}
.org-tree li:only-child{{ padding-top: 0;}}
.org-tree li:first-child::before, .org-tree li:last-child::after{{ border: 0 none; }}
.org-tree li:last-child::before{{
    border-right: 2px solid #0EA5E9; border-radius: 0 5px 0 0;
}}
.org-tree li:first-child::after{{ border-radius: 5px 0 0 0; }}
.org-tree ul ul::before{{
    content: ''; position: absolute; top: 0; left: 50%;
    border-left: 2px solid #0EA5E9; width: 0; height: 20px;
}}
.org-tree li .org-node {{
    border: 1.5px solid rgba(14, 165, 233, 0.45);
    padding: 14px 20px;
    text-decoration: none;
    display: inline-block;
    border-radius: 12px;
    transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    background: {T['bg_card']};
    box-shadow: 0 6px 18px rgba(14, 165, 233, 0.18), 0 2px 6px rgba(15, 23, 42, 0.08);
    white-space: nowrap;
    cursor: pointer;
    min-width: 175px;
    text-align: center;
}}
.org-tree li .org-node .dept-title {{
    color: {T['text_primary']};
    font-family: 'Plus Jakarta Sans', sans-serif;
    font-size: 15px;
    font-weight: 800;
    margin-bottom: 5px;
    transition: color 0.3s;
}}
.org-tree li .org-node .dept-sub {{
    color: {T['text_secondary']};
    font-family: 'Plus Jakarta Sans', sans-serif;
    font-size: 13px;
    font-weight: 600;
    transition: color 0.3s;
}}
.org-tree li .org-node:hover, .org-tree li .org-node.active-node {{
    background: linear-gradient(135deg, #0EA5E9 0%, #0284C7 100%) !important;
    border-color: #0284C7 !important;
    box-shadow: 0 10px 25px rgba(14, 165, 233, 0.35), 0 4px 10px rgba(15, 23, 42, 0.12) !important;
    transform: translateY(-4px);
}}
.org-tree li .org-node:hover .dept-title, .org-tree li .org-node:hover .dept-sub,
.org-tree li .org-node.active-node .dept-title, .org-tree li .org-node.active-node .dept-sub {{
    color: #FFFFFF !important;
}}
/* Đặc quyền cấp cao nhất: CEO Node */
.org-tree li .org-node.ceo-node, .org-tree li .org-node.ceo-node.active-node {{
    background: linear-gradient(135deg, #F59E0B 0%, #D97706 100%) !important;
    border: 2.5px solid #FEF3C7 !important;
    padding: 16px 28px !important;
    min-width: 240px !important;
    box-shadow: 0 10px 25px rgba(245, 158, 11, 0.35), 0 4px 10px rgba(15, 23, 42, 0.12) !important;
}}
.org-tree li .org-node.ceo-node:hover {{
    transform: translateY(-4px) scale(1.02) !important;
    box-shadow: 0 14px 32px rgba(245, 158, 11, 0.45), 0 6px 14px rgba(15, 23, 42, 0.15) !important;
    background: linear-gradient(135deg, #D97706 0%, #B45309 100%) !important;
    border-color: #FFFBEB !important;
}}
.org-tree li .org-node.ceo-node:hover+ul li::after, 
.org-tree li .org-node.ceo-node:hover+ul li::before, 
.org-tree li .org-node.ceo-node:hover+ul::before, 
.org-tree li .org-node.ceo-node:hover+ul ul::before{{
    border-color: #F59E0B !important;
}}
.org-tool-btn {{
    transition: all 0.2s ease !important;
    font-family: 'Plus Jakarta Sans', sans-serif !important;
}}
.org-tool-btn:hover {{
    background: #E0F2FE !important;
    border-color: #0EA5E9 !important;
    color: #0284C7 !important;
    transform: translateY(-1px);
}}
#btn-toggle-view:hover {{
    background: linear-gradient(135deg, #0284C7, #0369A1) !important;
    color: white !important;
    box-shadow: 0 6px 14px rgba(14,165,233,0.35) !important;
}}
@media print {{
    .org-tool-btn, #btn-toggle-view, [data-testid="stSidebar"], .stAppHeader {{
        display: none !important;
    }}
    .org-tree-card {{
        border: none !important;
        box-shadow: none !important;
    }}
}}
</style>
<div class="org-tree-card">
<div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px; padding-bottom: 16px; border-bottom: 1.5px dashed rgba(14, 165, 233, 0.3); flex-wrap: wrap; gap: 14px;">
<div style="display: flex; align-items: center; gap: 12px;">
<div style="width: 40px; height: 40px; border-radius: 12px; background: linear-gradient(135deg, #0EA5E9, #0284C7); color: white; display: flex; align-items: center; justify-content: center; font-size: 20px; box-shadow: 0 4px 12px rgba(14,165,233,0.3); flex-shrink: 0;">🏢</div>
<div>
<div style="font-weight: 800; font-size: 17px; color: {T['text_primary']}; font-family: 'Plus Jakarta Sans', sans-serif;">{'Sơ đồ Cấu trúc & Phân bổ Phòng ban' if is_vi_nf else '\u90e8\u7f72\u69cb\u6210\u30fb\u914d\u5c5e\u56f3'}</div>
<div style="font-size: 13.5px; color: {T['text_secondary']}; font-weight: 500; margin-top: 2px;">{'Nhấp vào từng phòng ban để xem chi tiết danh sách nhân sự' if is_vi_nf else '\u5404\u90e8\u7f72\u3092\u30af\u30ea\u30c3\u30af\u3057\u3066\u30e1\u30f3\u30d0\u30fc\u8a73\u7d30\u3092\u8868\u793a'}</div>
</div>
</div>
<div style="display: flex; align-items: center; gap: 8px; flex-wrap: wrap;">
<div style="display: flex; align-items: center; background: {T['bg_card']}; border: 1.5px solid {T['border']}; border-radius: 10px; overflow: hidden; box-shadow: 0 2px 5px rgba(0,0,0,0.03);">
<button id="btn-zoom-out" class="org-tool-btn" title="Thu nhỏ" style="padding: 6px 12px; border: none; background: transparent; font-weight: 800; color: {T['text_primary']}; cursor: pointer; font-size: 15px; border-right: 1px solid {T['border']};">−</button>
<span id="org-zoom-label" style="padding: 0 10px; font-size: 12.5px; font-weight: 700; color: {T['text_secondary']};">100%</span>
<button id="btn-zoom-in" class="org-tool-btn" title="Phóng to" style="padding: 6px 12px; border: none; background: transparent; font-weight: 800; color: {T['text_primary']}; cursor: pointer; font-size: 15px; border-left: 1px solid {T['border']};">+</button>
</div>
<button id="btn-org-print" class="org-tool-btn" style="display: flex; align-items: center; gap: 6px; padding: 6px 14px; background: {T['bg_card']}; border: 1.5px solid {T['border']}; border-radius: 10px; font-size: 13px; font-weight: 700; color: {T['text_primary']}; cursor: pointer; box-shadow: 0 2px 5px rgba(0,0,0,0.03);">
🖨️ {'Xuất file / In' if is_vi_nf else '印刷 / PDF'}
</button>
<button id="btn-toggle-view" class="org-tool-btn" style="display: flex; align-items: center; gap: 6px; padding: 6px 14px; background: linear-gradient(135deg, #0EA5E9, #0284C7); border: 1.5px solid #0284C7; border-radius: 10px; font-size: 13px; font-weight: 700; color: white; cursor: pointer; box-shadow: 0 4px 10px rgba(14,165,233,0.25);">
🔄 <span id="label-toggle-view">{'Xem Danh bạ' if is_vi_nf else 'リスト表示'}</span>
</button>
</div>
</div>
<div id="org-view-tree" class="org-tree">
<ul>
<li>
<div class="org-node ceo-node {'active-node' if active_d == 'Tổng giám đốc' else ''}" data-dept="Tổng giám đốc">
<div style="display: flex; align-items: center; justify-content: center; gap: 8px;">
{otaki_mini_img}
<span class="dept-title" style="font-size: 18px; font-weight: 800; color: white; letter-spacing: 0.3px;">{lbl_tgd}</span>
</div>
</div>
<ul>
<li>
<div class="org-node {'active-node' if active_d == 'Bộ phận kỹ thuật' else ''}" data-dept="Bộ phận kỹ thuật">
<div class="dept-title">{lbl_kt}</div>
<div class="dept-sub">{sub_kt}</div>
</div>
<ul>
<li>
<div class="org-node {'active-node' if active_d == 'Thiết kế cơ khí' else ''}" data-dept="Thiết kế cơ khí">
<div class="dept-title">{lbl_ck}</div>
<div class="dept-sub">{sub_ck}</div>
</div>
</li>
<li>
<div class="org-node {'active-node' if active_d == 'Thiết kế điện điều khiển' else ''}" data-dept="Thiết kế điện điều khiển">
<div class="dept-title">{lbl_dk}</div>
<div class="dept-sub">{sub_dk}</div>
</div>
<ul>
<li>
<div class="org-node {'active-node' if active_d == 'Thiết kế điện' else ''}" data-dept="Thiết kế điện">
<div class="dept-title">{lbl_dien}</div>
<div class="dept-sub">{sub_dien}</div>
</div>
</li>
<li>
<div class="org-node {'active-node' if active_d == 'Lập trình điều khiển' else ''}" data-dept="Lập trình điều khiển">
<div class="dept-title">{lbl_lt}</div>
<div class="dept-sub">{sub_lt}</div>
</div>
</li>
<li>
<div class="org-node {'active-node' if active_d == 'Mô phỏng 3D' else ''}" data-dept="Mô phỏng 3D">
<div class="dept-title">{lbl_mp}</div>
<div class="dept-sub">{sub_mp}</div>
</div>
</li>
</ul>
</li>
</ul>
</li>
<li>
<div class="org-node {'active-node' if active_d == 'Bộ phận hành chính - kế toán' else ''}" data-dept="Bộ phận hành chính - kế toán">
<div class="dept-title">{lbl_hc}</div>
<div class="dept-sub">{sub_hc}</div>
</div>
</li>
</ul>
</li>
</ul>
</div>
<div id="org-view-list" class="org-list-view" style="display: none; padding: 15px 5px;">
{list_view_html}
</div>
</div>""", unsafe_allow_html=True)

    components.html(f"""
    <script>
    (function() {{
        const parentDoc = window.parent.document;
        let bindTimer = null;
        let currentZoom = 1.0;
        const isVi = {str(is_vi_nf).lower()};

        function hideOrgBtns() {{
            const sidebar = parentDoc.querySelector('[data-testid="stSidebar"]');
            if (!sidebar) return;
            const btns = sidebar.querySelectorAll('button');
            btns.forEach(b => {{
                if (b.innerText && b.innerText.includes('ORG_BTN_')) {{
                    let el = b;
                    for (let i = 0; i < 8; i++) {{
                        if (!el.parentElement) break;
                        el = el.parentElement;
                        if (el.getAttribute('data-testid') === 'stElementContainer') break;
                    }}
                    el.style.cssText = 'display:none!important;visibility:hidden!important;';
                }}
            }});
        }}

        function bindTreeClicks() {{
            const nodes = parentDoc.querySelectorAll('.org-tree .org-node:not([data-bound="true"])');
            nodes.forEach(node => {{
                node.setAttribute('data-bound', 'true');
                node.style.cursor = 'pointer';
                node.addEventListener('click', function(e) {{
                    e.preventDefault();
                    e.stopPropagation();
                    const dept = this.getAttribute('data-dept');
                    if (!dept || dept === 'Bộ phận kỹ thuật') return;
                    const sidebar = parentDoc.querySelector('[data-testid="stSidebar"]');
                    if (!sidebar) return;
                    const btns = Array.from(sidebar.querySelectorAll('button'));
                    const target = btns.find(b => b.innerText && b.innerText.includes('ORG_BTN_' + dept));
                    if (target) target.click();
                }});
            }});
        }}

        function bindOrgTools() {{
            const btnZoomIn = parentDoc.querySelector('#btn-zoom-in:not([data-bound="true"])');
            const btnZoomOut = parentDoc.querySelector('#btn-zoom-out:not([data-bound="true"])');
            const btnPrint = parentDoc.querySelector('#btn-org-print:not([data-bound="true"])');
            const btnToggle = parentDoc.querySelector('#btn-toggle-view:not([data-bound="true"])');
            
            if (btnZoomIn) {{
                btnZoomIn.setAttribute('data-bound', 'true');
                btnZoomIn.addEventListener('click', function(e) {{
                    e.preventDefault();
                    currentZoom = Math.min(currentZoom + 0.1, 1.6);
                    applyZoom();
                }});
            }}
            if (btnZoomOut) {{
                btnZoomOut.setAttribute('data-bound', 'true');
                btnZoomOut.addEventListener('click', function(e) {{
                    e.preventDefault();
                    currentZoom = Math.max(currentZoom - 0.1, 0.5);
                    applyZoom();
                }});
            }}
            function applyZoom() {{
                const tree = parentDoc.querySelector('#org-view-tree');
                const label = parentDoc.querySelector('#org-zoom-label');
                if (tree) {{
                    tree.style.zoom = currentZoom;
                }}
                if (label) {{
                    label.innerText = Math.round(currentZoom * 100) + '%';
                }}
            }}
            if (btnPrint) {{
                btnPrint.setAttribute('data-bound', 'true');
                btnPrint.addEventListener('click', function(e) {{
                    e.preventDefault();
                    window.parent.print();
                }});
            }}
            if (btnToggle) {{
                btnToggle.setAttribute('data-bound', 'true');
                btnToggle.addEventListener('click', function(e) {{
                    e.preventDefault();
                    const treeView = parentDoc.querySelector('#org-view-tree');
                    const listView = parentDoc.querySelector('#org-view-list');
                    const label = parentDoc.querySelector('#label-toggle-view');
                    if (!treeView || !listView) return;
                    if (treeView.style.display !== 'none') {{
                        treeView.style.display = 'none';
                        listView.style.display = 'block';
                        if (label) label.innerText = isVi ? 'Sơ đồ cây' : 'ツリー表示';
                    }} else {{
                        treeView.style.display = 'flex';
                        listView.style.display = 'none';
                        if (label) label.innerText = isVi ? 'Xem Danh bạ' : 'リスト表示';
                    }}
                }});
            }}
        }}

        function doSetup() {{
            hideOrgBtns();
            bindTreeClicks();
            bindOrgTools();
        }}

        // Run once immediately
        doSetup();

        const observer = new MutationObserver(function(mutations) {{
            if (bindTimer) clearTimeout(bindTimer);
            bindTimer = setTimeout(doSetup, 80);
        }});

        observer.observe(parentDoc.body, {{
            childList: true,
            subtree: true,
            attributes: false,
            characterData: false
        }});
    }})();
    </script>
    """, height=0, width=0)

    st.divider()

    if not all_emps:
        st.markdown(f"""<div style='text-align:center; padding: 60px; background: rgba(255,255,255,0.6); border-radius: 20px; border: 2px dashed #CBD5E1; margin: 30px 0;'>
            <div style='font-size: 64px; margin-bottom: 16px;'>🏢</div>
            <b style='color:#64748B; font-size:18px;'>{'Chưa có dữ liệu nhân sự. Vui lòng tải lên Bảng chấm công hoặc thêm nhân viên thủ công ở Cài đặt!' if is_vi_nf else '\u793e\u54e1\u30c7\u30fc\u30bf\u304c\u3042\u308a\u307e\u305b\u3093\u3002'}</b>
        </div>""", unsafe_allow_html=True)
    else:
        dept_order = [
            "Tổng giám đốc", "Tổng Giám Đốc",
            "Bộ phận kỹ thuật",
            "Thiết kế cơ khí",
            "Thiết kế điện",
            "Lập trình điều khiển",
            "Mô phỏng 3D",
            "Bộ phận hành chính - kế toán", "Hành chính kế toán", "Hành chính"
        ]
        
        def get_dept_order(dept_name):
            for i, d in enumerate(dept_order):
                if d.lower() in dept_name.lower():
                    return i
            return 999
            
        sorted_depts = sorted(depts.items(), key=lambda x: get_dept_order(x[0]))
        
        # Departments to hide in the directory list below (but keep on org chart)
        hidden_in_list = {'Lập trình điều khiển'}

        for pb, p_emps in sorted_depts:
            if pb in hidden_in_list:
                continue
            display_pb = pb
            if not is_vi_nf:
                dept_ja_map = {
                    'Tổng giám đốc': '\u6700\u9ad8\u7d4c\u55b6\u8cac\u4efb\u8005 (CEO)',
                    'Bộ phận kỹ thuật': '\u6280\u8853\u90e8',
                    'Thiết kế cơ khí': '\u6a5f\u68b0\u8a2d\u8a08\u90e8',
                    'Thiết kế điện điều khiển': '\u5236\u5fa1\u30fb\u96fb\u6c17\u8a2d\u8a08\u90e8',
                    'Thiết kế điện': '\u96fb\u6c17\u8a2d\u8a08',
                    'Lập trình điều khiển': '\u5236\u5fa1\u30d7\u30ed\u30b0\u30e9\u30df\u30f3\u30b0',
                    'Mô phỏng 3D': '3D\u30b7\u30df\u30e5\u30ec\u30fc\u30b7\u30e7\u30f3',
                    'Bộ phận hành chính - kế toán': '\u7dcf\u52d9\u30fb\u7d4c\u7406\u90e8'
                }
                display_pb = dept_ja_map.get(pb, pb)
            if pb == 'Tổng giám đốc':
                icon_html = ""
                sub_lbl = "Ban lãnh đạo cấp cao" if is_vi_nf else "経営陣"
                count_badge = ""
            else:
                icon_html = '<div style="width: 28px; height: 28px; border-radius: 8px; background: linear-gradient(135deg, #FEF08A, #FACC15); color: #713F12; display: flex; align-items: center; justify-content: center; font-size: 15px; box-shadow: 0 2px 6px rgba(234, 179, 8, 0.2); flex-shrink: 0;">📂</div>'
                sub_lbl = "Danh sách nhân sự thuộc bộ phận" if is_vi_nf else "所属メンバー一覧"
                count_badge = f'<div style="background: {T["bg_card_hover"]}; color: {T["primary"]}; padding: 6px 14px; border-radius: 20px; font-weight: 800; font-size: 13.5px; border: 1px solid {T["border"]};">{len(p_emps)} {"nhân sự" if is_vi_nf else "名"}</div>'

            dept_card_html = f"""<div style="background: {T['bg_card_hover']}; border: 1.5px solid {T['border']}; border-radius: 24px; padding: 24px 28px; margin-bottom: 28px; box-shadow: 0 4px 16px rgba(0,0,0,0.03);">
<div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 20px; padding-bottom: 16px; border-bottom: 1.5px dashed {T['border']}; flex-wrap: wrap; gap: 12px;">
<div style="display: flex; align-items: center; gap: 14px;">
{icon_html}
<div>
<div style="font-size: 18px; font-weight: 800; color: {T['text_primary']}; font-family: 'Plus Jakarta Sans', sans-serif;">{display_pb}</div>
<div style="font-size: 13px; color: {T['text_secondary']}; font-weight: 600; margin-top: 2px;">{sub_lbl}</div>
</div>
</div>
{count_badge}
</div>
<div style="display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 16px;">"""
            
            for emp in p_emps:
                cv = emp.get('cv', '').strip()
                if not cv:
                    cv = "Nhân viên" if is_vi_nf else "\u793e\u54e1"
                char_avt = emp['ten'][0].upper() if emp.get('ten') else "?"
                emp_ma = emp.get('ma', '')
                emp_ten = emp.get('ten', '')
                is_ceo_card = (emp_ma == 'GD01' or 'otaki' in emp_ten.lower() or 'masahide' in emp_ten.lower() or '\u5927\u6edd' in emp_ten)
                if is_ceo_card:
                    if OTAKI_B64:
                        avt_html = f'<img src="data:image/png;base64,{OTAKI_B64}" style="width: 50px; height: 50px; border-radius: 50%; object-fit: cover; border: 2px solid #F59E0B; box-shadow: 0 2px 8px rgba(245,158,11,0.25); flex-shrink: 0;">'
                    else:
                        avt_html = f'<div style="width: 50px; height: 50px; border-radius: 50%; background: linear-gradient(135deg, #F59E0B, #D97706); display: flex; align-items: center; justify-content: center; color: white; font-weight: 800; font-size: 20px; flex-shrink: 0;">{char_avt}</div>'
                    id_part = ""
                else:
                    avt_html = f'<div style="width: 50px; height: 50px; border-radius: 50%; background: {T["bg_card"]}; color: {T["primary"]}; border: 1.5px solid {T["border"]}; display: flex; align-items: center; justify-content: center; font-weight: 800; font-size: 20px; flex-shrink: 0; box-shadow: 0 2px 6px rgba(0,0,0,0.05);">{char_avt}</div>'
                    id_part = f' <span style="font-size: 12px; color: {T["text_tertiary"]}; font-weight: 600; background: {T["bg_card_hover"]}; padding: 2px 8px; border-radius: 6px; margin-left: 6px; vertical-align: middle;">[{emp_ma}]</span>'

                dept_card_html += f"""<div style="background: {T['bg_card']}; padding: 18px 20px; border-radius: 16px; border: 1.5px solid {T['border']}; box-shadow: 0 4px 10px rgba(0,0,0,0.02); display: flex; align-items: center; gap: 16px; transition: all 0.2s;">
{avt_html}
<div style="overflow: hidden; flex-grow: 1;">
<div style="font-weight: 800; color: {T['text_primary']}; font-size: 16px; display: flex; align-items: center; flex-wrap: wrap; gap: 4px; font-family: 'Plus Jakarta Sans', sans-serif;" title="{emp_ten}">{emp_ten}{id_part}</div>
<div style="color: {T['primary']}; font-size: 13.5px; font-weight: 600; margin-top: 4px; font-family: 'Plus Jakarta Sans', sans-serif; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">{cv}</div>
</div>
</div>"""

            dept_card_html += """</div>
</div>"""
            st.markdown(dept_card_html, unsafe_allow_html=True)
    st.stop()

if st.session_state.app_page == "kpi_schedule":
    render_kpi_schedule_page()
    st.stop()

if st.session_state.app_page == "checkin":
    render_checkin_page()
    st.stop()

if st.session_state.app_page == "history":
    render_history_page()
    st.stop()






if st.session_state.get("show_history"):
    st.session_state.show_history = False
    st.rerun()



# ----- UPLOAD FILE (CHỈ XỬ LÝ KHI Ở TRANG CHẤM CÔNG - uploader ở main area) -----

if st.session_state.get('app_page', 'overview') == 'chamcong':
    uploaded_file = None  # placeholder

    is_vi_cc_exp = (st.session_state.get('lang', 'vi') == 'vi')

    # ===== FILE UPLOAD LUON HIEN THI (Chỉ khi đã có file) =====
    if st.session_state.get('df_raw') is not None:
        _fn = st.session_state.get('_loaded_filename', '')
        if _fn:
            _fn_disp = (f"\U0001f4cb File hi\u1ec7n t\u1ea1i: **{_fn}**" if is_vi_cc_exp
                        else f"\U0001f4cb \u73fe\u5728\u306e\u30d5\u30a1\u30a4\u30eb: **{_fn}**")
            st.caption(_fn_disp)
    
        _up_label = ("\U0001f4e4 T\u1ea3i file ch\u1ea5m c\u00f4ng (k\u00e9o th\u1ea3 ho\u1eb7c b\u1ea5m ch\u1ecdn \u2014 s\u1ebd thay th\u1ebf file hi\u1ec7n t\u1ea1i)"
                     if is_vi_cc_exp else
                     "\U0001f4e4 \u52e4\u52e0\u30d5\u30a1\u30a4\u30eb\u3092\u30a2\u30c3\u30d7\u30ed\u30fc\u30c9 (\u30c9\u30e9\u30c3\u30b0&\u30c9\u30ed\u30c3\u30d7 \u307e\u305f\u306f\u30af\u30ea\u30c3\u30af \u2014 \u73fe\u5728\u306e\u30d5\u30a1\u30a4\u30eb\u3092\u5207\u308a\u6233\u3048\u308b)")
        inline_new_file = st.file_uploader(
            _up_label,
            type=["xlsx", "xls", "csv", "txt", "dat", "tsv", "xlsm", "xlsb"],
            key="inline_file_uploader",
            label_visibility="visible",
            accept_multiple_files=True
        )
        if inline_new_file:
            for _k in ['df_raw', 'mapping_auto', 'manual_leave', 'manual_ot',
                       'manual_ot_reason', 'manual_hc', 'manual_total']:
                st.session_state.pop(_k, None)
            st.session_state['_pending_new_file'] = inline_new_file
            st.rerun()

    # Nut xoa du lieu (chi hien khi da co file)
    if st.session_state.get('df_raw') is not None:
        _lbl_reset = "\U0001f5d1\ufe0f X\u00f3a d\u1eef li\u1ec7u" if is_vi_cc_exp else "\U0001f5d1\ufe0f \u30ea\u30bb\u30c3\u30c8"
        if st.button(_lbl_reset, key="btn_reset_file_top", type="secondary"):
            for _k in ['df_raw', 'mapping_auto', '_pending_new_file', 'manual_leave', 'manual_ot',
                       'manual_ot_reason', 'manual_hc', 'manual_total', '_loaded_filename']:
                st.session_state.pop(_k, None)
            st.rerun()

    # ===== EXPANDER CAI DAT (gon) =====
    _settings_lbl = ("\u2699\ufe0f Thi\u1ebft l\u1eadp & C\u1ea5u h\u00ecnh"
                     if is_vi_cc_exp else "\u2699\ufe0f \u8a2d\u5b9a\u3068\u69cb\u6210")
    _settings_expanded = st.session_state.get('_settings_expanded', False)
    with st.expander(_settings_lbl, expanded=_settings_expanded):
        if _settings_expanded:
            render_integrated_settings_content()
            if st.button("\u2715 " + ("Thu g\u1ecdn" if is_vi_cc_exp else "\u9589\u3058\u308b"), key="cc_close_settings"):
                st.session_state['_settings_expanded'] = False
                st.rerun()
        else:
            if st.button("\u2699\ufe0f " + ("M\u1edf c\u00e0i \u0111\u1eb7t" if is_vi_cc_exp else "\u8a2d\u5b9a\u3092\u958b\u304f"), key="cc_open_settings", use_container_width=True):
                st.session_state['_settings_expanded'] = True
                st.rerun()
else:
    uploaded_file = None




# ----- WELCOME CARD & UPLOAD ZONE -----
if st.session_state.get('df_raw') is None and st.session_state.get('app_page', 'overview') == 'chamcong':
    # ===== KHU VỰC KÉO THẢ FILE CHO TRANG CHẤM CÔNG =====
    is_vi_cc = (st.session_state.get('lang', 'vi') == 'vi')
    is_sepia = (st.session_state.theme_mode == 'sepia') if 'theme_mode' in st.session_state else False

    # --- Step indicator ---
    _s1 = "Tải file" if is_vi_cc else "\u30a2\u30c3\u30d7\u30ed\u30fc\u30c9"
    _s2 = "Xác nhận cột" if is_vi_cc else "\u5217\u30de\u30c3\u30d4\u30f3\u30b0"
    _s3 = "Xem kết quả" if is_vi_cc else "\u7d50\u679c\u3092\u78ba\u8a8d"
    st.markdown(f"""
    <div class="vmos-step-bar" style="padding:18px 24px; background: {T['bg_card']}BF;
         backdrop-filter:blur(14px); border-radius:16px; border:1px solid {T['border']};
         box-shadow:{T['shadow']}; margin-bottom:24px;">
        <div class="vmos-step">
            <div class="vmos-step-dot active">1</div>
            <span class="vmos-step-label active">{_s1}</span>
        </div>
        <div class="vmos-step-connector"></div>
        <div class="vmos-step">
            <div class="vmos-step-dot pending">2</div>
            <span class="vmos-step-label">{_s2}</span>
        </div>
        <div class="vmos-step-connector"></div>
        <div class="vmos-step">
            <div class="vmos-step-dot pending">3</div>
            <span class="vmos-step-label">{_s3}</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown(f"""
    <style>
    /* Ẩn padding mặc định của Streamlit để uploader nằm gọn */
    .upload-zone-wrapper {{ margin-top: 30px; }}
    div[data-testid="stFileUploader"] > label {{ display: none !important; }}
    div[data-testid="stFileUploader"] section {{
        background: {T['bg_card']} !important;
        border: 3px dashed {T['border']} !important;
        border-radius: 28px !important;
        padding: 55px 30px !important;
        text-align: center !important;
        box-shadow: {T['shadow']} !important;
        transition: all 0.3s ease !important;
    }}
    div[data-testid="stFileUploader"] section:hover {{
        border-color: {T['primary']} !important;
        background: {T['bg_content']} !important;
        box-shadow: {T['shadow_glow']} !important;
        transform: translateY(-3px) !important;
    }}
    div[data-testid="stFileUploader"] section > div > div > p {{
        color: {T['text_primary']} !important;
        font-size: 18px !important;
        font-weight: 700 !important;
        font-family: 'Plus Jakarta Sans', 'Inter', sans-serif !important;
    }}
    div[data-testid="stFileUploader"] section > div > div > small {{
        color: {T['text_secondary']} !important;
        font-size: 14px !important;
        font-weight: 600 !important;
    }}
    div[data-testid="stFileUploader"] section svg {{
        color: {T['primary']} !important;
        width: 60px !important;
        height: 60px !important;
        margin-bottom: 10px !important;
    }}
    div[data-testid="stFileUploader"] button[kind="secondary"] {{
        background: {T['primary_gradient']} !important;
        color: white !important;
        border: none !important;
        border-radius: 14px !important;
        font-weight: 800 !important;
        font-size: 15px !important;
        padding: 10px 28px !important;
        box-shadow: {T['shadow']} !important;
        margin-top: 12px !important;
    }}
    div[data-testid="stFileUploader"] button[kind="secondary"]:hover {{
        background: {T['accent_gradient']} !important;
        transform: translateY(-2px) !important;
        box-shadow: {T['shadow_glow']} !important;
    }}
    .chamcong-upload-header {{
        text-align: center;
        margin-bottom: 18px;
    }}
    .chamcong-upload-title {{
        font-family: 'Plus Jakarta Sans', 'Inter', sans-serif !important;
        font-size: 32px;
        font-weight: 800;
        background: {T['primary_gradient']};
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        letter-spacing: 0.5px;
        margin-bottom: 8px;
        filter: drop-shadow(0 2px 6px rgba(0,0,0,0.18));
    }}
    .chamcong-upload-sub {{
        font-family: 'Plus Jakarta Sans', 'Inter', sans-serif !important;
        font-size: 15px;
        color: {T['text_primary']};
        font-weight: 600;
        background: {T['bg_app']}D9;
        display: inline-block;
        padding: 6px 20px;
        border-radius: 20px;
        backdrop-filter: blur(8px);
        margin-bottom: 0;
        letter-spacing: 0.2px;
    }}
    </style>
    """, unsafe_allow_html=True)

    st.markdown(f"""
    <div class="chamcong-upload-header">
        <div class="chamcong-upload-title">📋 {'Tải lên File Chấm công' if is_vi_cc else '\u30c1\u30a7\u30c3\u30af\u30d5\u30a1\u30a4\u30eb\u3092\u30a2\u30c3\u30d7\u30ed\u30fc\u30c9'}</div>
        <div class="chamcong-upload-sub">{'Hỗ trợ Excel (.xlsx, .xls, .xlsm) · CSV · TXT · DAT' if is_vi_cc else 'Excel (.xlsx, .xls, .xlsm) · CSV · TXT · DAT \u306b\u5bfe\u5fdc'}</div>
    </div>
    """, unsafe_allow_html=True)

    uploaded_file = st.file_uploader(
        "chamcong_main_upload",
        type=["xlsx","xls","csv","txt","dat","tsv","xlsm","xlsb"],
        label_visibility="collapsed",
        key="main_file_uploader",
        accept_multiple_files=True
    )
    # Nếu người dùng dùng nút "Đổi file" từ settings, xử lý file pending
    if not uploaded_file and st.session_state.get('_pending_new_file'):
        uploaded_file = st.session_state.pop('_pending_new_file')

elif st.session_state.get('df_raw') is None:
    # Trang không phải chamcong, không có file → hiển thị Enterprise Dashboard thay vì card trắng trống
    render_enterprise_dashboard()


# ----- LANDING PAGE CŨ (Đã vô hiệu hóa) -----
if False and uploaded_file is None and st.session_state.df_raw is None:
    pass

    st.markdown(f"""
<style>
/* --- PREMIUM 3-STEP GUIDE CSS --- */
.chamcong-landing {{
    position: fixed;
    top: 50%;
    left: 55%; /* Offset for sidebar */
    transform: translate(-50%, -50%);
    text-align: center;
    font-family: 'Inter', 'Be Vietnam Pro', sans-serif;
    pointer-events: none;
    width: 85%;
    max-width: 900px;
}}

.guide-title {{
    font-size: 34px;
    font-weight: 800;
    margin-bottom: 50px;
    background: linear-gradient(135deg, #0284C7, #0EA5E9, #38BDF8);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    letter-spacing: -0.5px;
    filter: drop-shadow(0 2px 4px rgba(14,165,233,0.15));
}}

.guide-steps {{
    display: flex;
    justify-content: space-between;
    gap: 25px;
}}

.guide-step {{
    flex: 1;
    background: rgba(255, 255, 255, 0.7);
    backdrop-filter: blur(16px);
    -webkit-backdrop-filter: blur(16px);
    padding: 40px 25px;
    border-radius: 24px;
    box-shadow: 
        0 4px 6px -1px rgba(0, 0, 0, 0.05),
        0 10px 15px -3px rgba(0, 0, 0, 0.05),
        inset 0 1px 0 rgba(255, 255, 255, 0.9);
    border: 1px solid rgba(255, 255, 255, 0.5);
    position: relative;
    transition: all 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.275);
}}

/* Hover effects */
.guide-steps:hover .guide-step:not(:hover) {{
    transform: scale(0.95);
    opacity: 0.7;
}}

.guide-step:hover {{
    transform: translateY(-10px) scale(1.02);
    box-shadow: 
        0 20px 25px -5px rgba(14, 165, 233, 0.15),
        0 10px 10px -5px rgba(14, 165, 233, 0.1),
        inset 0 1px 0 rgba(255, 255, 255, 1);
    background: rgba(255, 255, 255, 0.9);
}}

.icon-wrapper {{
    width: 80px;
    height: 80px;
    border-radius: 50%;
    margin: 0 auto 25px auto;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 36px;
    background: linear-gradient(135deg, #F0F9FF, #E0F2FE);
    box-shadow: 0 10px 20px rgba(14, 165, 233, 0.1);
    position: relative;
}}

.icon-wrapper::after {{
    content: '';
    position: absolute;
    top: -5px; left: -5px; right: -5px; bottom: -5px;
    border-radius: 50%;
    background: linear-gradient(135deg, rgba(14,165,233,0.3), rgba(56,189,248,0));
    z-index: -1;
}}

.guide-step-title {{
    font-size: 20px;
    font-weight: 700;
    color: #0F172A;
    margin-bottom: 12px;
}}

.guide-step-desc {{
    font-size: 14px;
    color: #475569;
    line-height: 1.6;
    font-weight: 400;
}}

.guide-arrow {{
    position: absolute;
    top: 50%;
    right: -30px;
    transform: translateY(-50%);
    font-size: 28px;
    color: #94A3B8;
    z-index: 10;
    filter: drop-shadow(0 2px 4px rgba(0,0,0,0.05));
}}
.guide-step:last-child .guide-arrow {{
    display: none;
}}

.upload-pulse .icon-wrapper {{
    animation: pulse-glow 2.5s infinite;
}}

@keyframes pulse-glow {{
    0% {{ box-shadow: 0 0 0 0 rgba(14, 165, 233, 0.4); }}
    70% {{ box-shadow: 0 0 0 20px rgba(14, 165, 233, 0); }}
    100% {{ box-shadow: 0 0 0 0 rgba(14, 165, 233, 0); }}
}}
/* Copyright — sát xuống chân trang */
.vimos-copyright {{
    position: fixed;
    bottom: 8px;
    left: 50%;
    transform: translateX(-50%);
    width: 100%;
    text-align: center;
    color: #94A3B8;
    font-size: 12px;
    font-weight: 400;
    z-index: 10;
    pointer-events: none;
}}
</style>

<div class="chamcong-landing">
    <div class="guide-title">{t('guide_title')}</div>
    <div class="guide-steps">
        <div class="guide-step upload-pulse">
            <div class="icon-wrapper">📁</div>
            <div class="guide-step-title">{t('guide_step1_title')}</div>
            <div class="guide-step-desc">{t('guide_step1_desc')}</div>
            <span class="guide-arrow">➔</span>
        </div>
        <div class="guide-step">
            <div class="icon-wrapper">⚙️</div>
            <div class="guide-step-title">{t('guide_step2_title')}</div>
            <div class="guide-step-desc">{t('guide_step2_desc')}</div>
            <span class="guide-arrow">➔</span>
        </div>
        <div class="guide-step">
            <div class="icon-wrapper">📊</div>
            <div class="guide-step-title">{t('guide_step3_title')}</div>
            <div class="guide-step-desc">{t('guide_step3_desc')}</div>
        </div>
    </div>
</div>
<div class="vimos-copyright">Copyright &copy; 2026 V-mos System</div>
""", unsafe_allow_html=True)

# ----- XỬ LÝ FILE -----
if uploaded_file:
    file_list = uploaded_file if isinstance(uploaded_file, list) else [uploaded_file]
    current_names = " + ".join([f.name for f in file_list]) if len(file_list) <= 3 else f"{len(file_list)} files ({file_list[0].name}, ...)"
    if st.session_state.df_raw is None or st.session_state.get('last_uploaded') != current_names:
        with st.spinner("Đang đọc và gộp dữ liệu file chấm công..."):
            dfs = []
            for f in file_list:
                f.seek(0)
                df_sub = parse_excel_file(f)
                if df_sub is not None and not df_sub.empty:
                    dfs.append(df_sub)
            if dfs:
                df = pd.concat(dfs, ignore_index=True) if len(dfs) > 1 else dfs[0]
                if file_list:
                    file_list[0].seek(0)
                    st.session_state['uploaded_file_bytes'] = file_list[0].read()
                st.session_state.df_raw = df
                st.session_state.last_uploaded = current_names
                st.session_state['_loaded_filename'] = current_names
                mapping_auto = auto_detect_columns(df)
                req_keys = ['ma_nv', 'ten_nv', 'ngay', 'gio_vao', 'gio_ra']
                if all(k in mapping_auto for k in req_keys):
                    st.session_state.mapping = mapping_auto
                    st.session_state.step = 3
                    st.rerun()
                else:
                    st.session_state.step = 2
                    st.session_state.mapping_auto = mapping_auto
elif st.session_state.df_raw is not None and st.session_state.get('app_page', 'chamcong') not in ('chamcong',):
    # Chỉ xóa df_raw khi đả chuyển sang trang khác và không có file nào cả
    pass

# ----- MAPPING CỘT -----
if st.session_state.get('app_page', 'overview') == 'chamcong' and st.session_state.step == 2 and st.session_state.df_raw is not None:
    df_raw = st.session_state.df_raw
    mapping_auto = st.session_state.mapping_auto
    st.markdown("### Dữ liệu thô (10 dòng đầu)")
    st.dataframe(df_raw.head(10), use_container_width=True, hide_index=True)
    st.sidebar.markdown("### 🔗 Bước 2: Xác nhận mapping cột")
    st.sidebar.warning("⚠️ Không nhận diện đủ 5 cột.")
    columns = ["-- Chọn cột --"] + list(df_raw.columns)
    def get_index(col_name):
        return columns.index(col_name) if col_name in columns else 0
    map_emp_id = st.sidebar.selectbox("Mã NV", columns, index=get_index(mapping_auto.get('ma_nv')))
    map_emp_name = st.sidebar.selectbox("Tên NV", columns, index=get_index(mapping_auto.get('ten_nv')))
    map_date = st.sidebar.selectbox("Ngày làm việc", columns, index=get_index(mapping_auto.get('ngay')))
    map_in = st.sidebar.selectbox("Giờ vào", columns, index=get_index(mapping_auto.get('gio_vao')))
    map_out = st.sidebar.selectbox("Giờ ra", columns, index=get_index(mapping_auto.get('gio_ra')))
    
    st.sidebar.markdown("<small style='color:#64748B'>— Tuỳ chọn (để trống nếu không có) —</small>", unsafe_allow_html=True)
    opt_cols = ["(Không có)"] + list(df_raw.columns)
    map_chuc_vu = st.sidebar.selectbox("Chức vụ (tuỳ chọn)", opt_cols, index=0)
    map_phong_ban = st.sidebar.selectbox("Phòng ban (tuỳ chọn)", opt_cols, index=0)
    
    if st.sidebar.button("Xác nhận Mapping", type="primary", use_container_width=True):
        if "-- Chọn cột --" in [map_emp_id, map_emp_name, map_date, map_in, map_out]:
            st.sidebar.error("❌ Vui lòng chọn đủ 5 cột bắt buộc!")
        else:
            new_mapping = {'ma_nv': map_emp_id, 'ten_nv': map_emp_name, 'ngay': map_date, 'gio_vao': map_in, 'gio_ra': map_out}
            if map_chuc_vu != "(Không có)": new_mapping['chuc_vu'] = map_chuc_vu
            if map_phong_ban != "(Không có)": new_mapping['phong_ban'] = map_phong_ban
            st.session_state.mapping = new_mapping
            st.session_state.step = 3
            st.rerun()

# ----- TÍNH TOÁN & KẾT QUẢ -----
if st.session_state.get('app_page', 'overview') == 'chamcong' and st.session_state.step >= 3 and "mapping" in st.session_state and st.session_state.get('df_raw') is not None:
    m = st.session_state.mapping
    df_process = st.session_state.df_raw.copy()
    df_process = df_process.dropna(subset=[m['ma_nv']])
    df_process = df_process[df_process[m['ma_nv']].astype(str).str.strip() != '']
    
    @st.cache_data(show_spinner=False)
    def cached_clean_date(col): return col.apply(clean_date)
    df_process["_parsed_date"] = cached_clean_date(df_process[m['ngay']])
    df_process = df_process.dropna(subset=["_parsed_date"])

    if df_process.empty:
        st.warning("Không có dữ liệu hợp lệ (Lỗi đọc ngày tháng hoặc Mã NV trống).")
    else:
        min_date = df_process["_parsed_date"].min()
        max_date = df_process["_parsed_date"].max()
        default_start = datetime.date(2026, 5, 21)
        default_end = datetime.date(2026, 6, 20)
        if min_date.date() > default_end or max_date.date() < default_start:
            default_start = min_date.date()
            default_end = max_date.date()

        start_date = st.session_state.get("main_start_date", default_start)
        end_date = st.session_state.get("main_end_date", default_end)

        years_in_range = range(start_date.year, end_date.year + 1)
        fixed_holidays = set(get_fixed_holidays_for_years(years_in_range))
        
        if "custom_holidays" not in st.session_state:
            st.session_state.custom_holidays = set()
        if "custom_workdays" not in st.session_state:
            st.session_state.custom_workdays = set()

        holiday_dates = fixed_holidays | st.session_state.custom_holidays
        makeup_dates = st.session_state.custom_workdays
        working_days, holiday_info, makeup_info = calculate_working_days(start_date, end_date, holiday_dates, makeup_dates)

        if makeup_info:
            st.markdown('<div class="card"><div class="card-title"><span class="card-icon">💼</span>Ngày làm bù trong kỳ</div></div>', unsafe_allow_html=True)
            for m in sorted(makeup_info, key=lambda x: x["date"]):
                date_str = m["date"].strftime("%d/%m/%Y")
                st.info(f"💼 **{date_str} ({m['weekday_label']})** được tính là ngày làm việc (Ngày làm bù).")

        if holiday_info:
            st.markdown('<div class="card"><div class="card-title"><span class="card-icon">📅</span>Ngày lễ trong kỳ</div></div>', unsafe_allow_html=True)
            for h in sorted(holiday_info, key=lambda x: x["date"]):
                date_str = h["date"].strftime("%d/%m/%Y")
                if h["is_workday"]:
                    st.warning(f"⚠️ **{date_str} ({h['weekday_label']})** trùng vào ngày làm việc — đã trừ 1 ngày công.")
                else:
                    st.info(f"ℹ️ {date_str} ({h['weekday_label']}) rơi vào ngày nghỉ cuối tuần.")

        mask = (df_process["_parsed_date"].dt.date >= start_date) & (df_process["_parsed_date"].dt.date <= end_date)
        df_filtered = df_process.loc[mask].copy()
        
        if not m.get('phong_ban'): m['phong_ban'] = 'Phòng ban'
        if not m.get('chuc_vu'): m['chuc_vu'] = 'Chức vụ'

        if st.session_state.manual_emps:
            new_rows = []
            curr_d = start_date
            while curr_d <= end_date:
                for emp in st.session_state.manual_emps:
                    row_data = {
                        m['ma_nv']: emp['ma'],
                        m['ten_nv']: emp['ten'],
                        "_parsed_date": pd.to_datetime(curr_d),
                        m['ngay']: curr_d.strftime('%d/%m/%Y'),
                        m['gio_vao']: pd.NaT,
                        m['gio_ra']: pd.NaT
                    }
                    if 'chuc_vu' in m: row_data[m['chuc_vu']] = emp['cv']
                    if 'phong_ban' in m: row_data[m['phong_ban']] = emp['pb']
                    new_rows.append(row_data)
                curr_d += datetime.timedelta(days=1)
            if new_rows:
                df_new = pd.DataFrame(new_rows)
                df_filtered = pd.concat([df_filtered, df_new], ignore_index=True)


        # APPLY DELETED EMPS AND EDITS
        if st.session_state.get('deleted_emps'):
            df_filtered = df_filtered[~df_filtered[m['ma_nv']].astype(str).str.strip().isin(st.session_state.deleted_emps)]
            
        if st.session_state.get('edited_emps'):
            for ma, edits in st.session_state.edited_emps.items():
                mask_emp = df_filtered[m['ma_nv']].astype(str).str.strip() == ma
                if 'phong_ban' in m and edits.get('pb'):
                    df_filtered.loc[mask_emp, m['phong_ban']] = edits['pb']
                if 'chuc_vu' in m and edits.get('cv'):
                    df_filtered.loc[mask_emp, m['chuc_vu']] = edits['cv']

        if df_filtered.empty:
            st.warning("Không có dữ liệu trong khoảng thời gian đã chọn.")
        else:
            if "gps_synced" not in st.session_state: st.session_state.gps_synced = set()
            try:
                df_gps_all = get_field_checkins(limit=10000)
                if not df_gps_all.empty:
                    emp_info_map = {}
                    for _, erow in df_filtered.iterrows():
                        ema = str(erow[m['ma_nv']]).strip().upper()
                        if ema not in emp_info_map:
                            emp_info_map[ema] = {
                                'ten': erow.get(m['ten_nv'], ''),
                                'cv': erow.get(m.get('chuc_vu', 'Chức vụ'), ''),
                                'pb': erow.get(m.get('phong_ban', 'Phòng ban'), '')
                            }

                    gps_rows_to_add = []
                    for _, g_row in df_gps_all.iterrows():
                        g_ma = str(g_row['ma_nv']).split(' - ')[0].strip().upper()
                        g_time = pd.to_datetime(g_row['thoi_gian'], format='%d/%m/%Y %H:%M:%S', errors='coerce')
                        if pd.notna(g_time) and start_date <= g_time.date() <= end_date:
                            g_date_str = g_time.strftime('%d/%m/%Y')
                            g_hour_str = g_time.strftime('%H:%M')
                            g_loai = str(g_row['loai'])
                            
                            mask_match = (df_filtered[m['ma_nv']].astype(str).str.strip().str.upper() == g_ma) & (df_filtered["_parsed_date"].dt.strftime('%d/%m/%Y') == g_date_str)
                            if mask_match.any():
                                for idx_m in df_filtered[mask_match].index:
                                    vao_ht = df_filtered.loc[idx_m, m['gio_vao']]
                                    ra_ht = df_filtered.loc[idx_m, m['gio_ra']]
                                    vao_tr = pd.isna(vao_ht) or str(vao_ht).strip().lower() in ['', 'nan', 'none', 'nat']
                                    ra_tr = pd.isna(ra_ht) or str(ra_ht).strip().lower() in ['', 'nan', 'none', 'nat']
                                    
                                    if 'Vào ca' in g_loai or 'Check-in' in g_loai or '\u51fa\u52e4' in g_loai:
                                        if vao_tr: df_filtered.loc[idx_m, m['gio_vao']] = g_hour_str
                                    elif 'Tan ca' in g_loai or 'Check-out' in g_loai or '\u9000\u52e4' in g_loai:
                                        if ra_tr: df_filtered.loc[idx_m, m['gio_ra']] = g_hour_str
                                        
                                    vao_sau = df_filtered.loc[idx_m, m['gio_vao']]
                                    ra_sau = df_filtered.loc[idx_m, m['gio_ra']]
                                    vao_s_tr = pd.isna(vao_sau) or str(vao_sau).strip().lower() in ['', 'nan', 'none', 'nat']
                                    ra_s_tr = pd.isna(ra_sau) or str(ra_sau).strip().lower() in ['', 'nan', 'none', 'nat']
                                    
                                    if not vao_s_tr and ra_s_tr:
                                        df_filtered.loc[idx_m, m['gio_ra']] = "17:00"
                                    elif vao_s_tr and not ra_s_tr:
                                        df_filtered.loc[idx_m, m['gio_vao']] = "08:00"
                            else:
                                einfo = emp_info_map.get(g_ma, {'ten': str(g_row['ten_nv']), 'cv': '', 'pb': ''})
                                vao_val = g_hour_str if ('Vào ca' in g_loai or 'Check-in' in g_loai or '\u51fa\u52e4' in g_loai) else "08:00"
                                ra_val = g_hour_str if ('Tan ca' in g_loai or 'Check-out' in g_loai or '\u9000\u52e4' in g_loai) else "17:00"
                                
                                new_g_row = {
                                    m['ma_nv']: g_ma,
                                    m['ten_nv']: einfo['ten'],
                                    "_parsed_date": pd.to_datetime(g_time.date()),
                                    m['ngay']: g_date_str,
                                    m['gio_vao']: vao_val,
                                    m['gio_ra']: ra_val
                                }
                                if 'chuc_vu' in m: new_g_row[m['chuc_vu']] = einfo['cv']
                                if 'phong_ban' in m: new_g_row[m['phong_ban']] = einfo['pb']
                                gps_rows_to_add.append(new_g_row)
                                
                            st.session_state.gps_synced.add((g_ma, g_date_str))
                            
                    if gps_rows_to_add:
                        df_filtered = pd.concat([df_filtered, pd.DataFrame(gps_rows_to_add)], ignore_index=True)
            except Exception:
                pass

            with st.spinner("⏳ Đang tính toán giờ làm..."):
                @st.cache_data(show_spinner=False)
                def cached_calc_hours(df_times, vc, rc, tb, tk, max_h):
                    return df_times.apply(lambda row: calculate_working_hours(
                        row['vao'], row['ra'],
                        start_chuan=time_to_float(vc), end_chuan=time_to_float(rc),
                        lunch_start=time_to_float(tb), lunch_end=time_to_float(tk),
                        max_hours=max_h
                    ), axis=1)
                
                df_times = pd.DataFrame({'vao': df_filtered[m['gio_vao']], 'ra': df_filtered[m['gio_ra']]})
                df_calc = cached_calc_hours(
                    df_times, 
                    st.session_state.gio_vao_chuan, st.session_state.gio_ra_chuan,
                    st.session_state.nghi_trua_bat_dau, st.session_state.nghi_trua_ket_thuc,
                    st.session_state.so_gio_toi_da
                )
                df_filtered["Giờ hành chính"] = df_calc.apply(lambda x: x['admin_hours'] if isinstance(x, dict) else 0.0)
                df_filtered["_is_chieu"] = df_calc.apply(lambda x: x.get('is_chieu', False) if isinstance(x, dict) else False)
                if 'di_tre' in m and m['di_tre'] in df_filtered.columns:
                    df_filtered["Phút đi trễ"] = pd.to_numeric(df_filtered[m['di_tre']], errors='coerce').fillna(0)
                else:
                    df_filtered["Phút đi trễ"] = df_calc.apply(lambda x: x['di_tre'] if isinstance(x, dict) else 0)
                
                if 've_som' in m and m['ve_som'] in df_filtered.columns:
                    df_filtered["Phút về sớm"] = pd.to_numeric(df_filtered[m['ve_som']], errors='coerce').fillna(0)
                else:
                    df_filtered["Phút về sớm"] = df_calc.apply(lambda x: x['ve_som'] if isinstance(x, dict) else 0)
                if 'ot' in m and m['ot'] in df_filtered.columns:
                    df_filtered["Giờ OT"] = pd.to_numeric(df_filtered[m['ot']], errors='coerce').fillna(0)
                else:
                    df_filtered["Giờ OT"] = 0.0
                    
                # Tính Tổng giờ và Giờ hành chính
                # Nếu có tong_gio trong mapping (nghĩa là tải từ DB), ta ưu tiên dùng dữ liệu đã lưu
                if 'tong_gio' in m and m['tong_gio'] in df_filtered.columns:
                    df_filtered["Số giờ làm thực tế"] = pd.to_numeric(df_filtered[m['tong_gio']], errors='coerce').fillna(0)
                    if 'hc_manual' in m and m['hc_manual'] in df_filtered.columns:
                        df_filtered["Giờ hành chính"] = pd.to_numeric(df_filtered[m['hc_manual']], errors='coerce').fillna(0)
                    else:
                        df_filtered["Giờ hành chính"] = df_filtered["Số giờ làm thực tế"] - df_filtered["Giờ OT"]
                else:
                    df_filtered["Giờ hành chính"] = df_calc.apply(lambda x: x['admin_hours'] if isinstance(x, dict) else 0.0)
                    df_filtered["Số giờ làm thực tế"] = pd.to_numeric(df_filtered["Giờ hành chính"], errors='coerce').fillna(0) + pd.to_numeric(df_filtered["Giờ OT"], errors='coerce').fillna(0)

                df_filtered["Ngày"] = df_filtered["_parsed_date"].dt.strftime('%d/%m/%Y')
                df_filtered = df_filtered.sort_values(by=[m['ma_nv'], "_parsed_date"])
                df_filtered["_nv_label"] = df_filtered[m['ma_nv']].astype(str) + " - " + df_filtered[m['ten_nv']].astype(str)
                danh_sach_nv = sorted(df_filtered["_nv_label"].unique().tolist())

            pass  # Filter và kỳ công đã chuyển sang main area


            # ----- BƯỚC 3 & 4: KỲ CÔNG + LỌC DỮ LIỆU (Hiển thị ở Main Area) -----
            is_vi = (st.session_state.get('lang', 'vi') == 'vi')
            is_vi_r = is_vi
            st.markdown(f"""
            <div style="background:rgba(255,255,255,0.94); backdrop-filter:blur(16px); border-radius:18px; padding:18px 22px 14px 22px; border:1.5px solid rgba(14,165,233,0.25); box-shadow: 0 8px 25px rgba(14,165,233,0.25); margin-bottom:16px;">
                <div style="font-weight:900; font-size:15px; color:#0284C7; margin-bottom:12px; display:flex; gap:8px; align-items:center;">{t("auto_text_app_28")}</div>
            </div>
            """, unsafe_allow_html=True)
            
            if st.button(t("auto_text_app_29"), use_container_width=False, key="btn_q_std"):
                    st.session_state["main_start_date"] = datetime.date(2026, 5, 21)
                    st.session_state["main_end_date"] = datetime.date(2026, 6, 20)
                    st.rerun()

            ctrl_c1, ctrl_c2, ctrl_c3 = st.columns([1.2, 1.2, 2])
            with ctrl_c1:
                st.markdown(f"<div style='font-size:13px; font-weight:700; color:#0F172A; margin-bottom:4px;'>{t("auto_text_app_30")}</div>", unsafe_allow_html=True)
                start_date_main = st.date_input("", value=start_date, key="main_start_date", label_visibility="collapsed")
            with ctrl_c2:
                st.markdown(f"<div style='font-size:13px; font-weight:700; color:#0F172A; margin-bottom:4px;'>{t("auto_text_app_31")}</div>", unsafe_allow_html=True)
                end_date_main = st.date_input("", value=end_date, key="main_end_date", label_visibility="collapsed")
            with ctrl_c3:
                st.markdown(f"<div style='font-size:13px; font-weight:700; color:#0F172A; margin-bottom:4px;'>{t("auto_text_app_32")}</div>", unsafe_allow_html=True)
                chon_nv = st.multiselect("", options=danh_sach_nv, default=[], placeholder=t("auto_text_app_33"), label_visibility="collapsed", key="main_chon_nv")
            
            # Cập nhật lại nếu người dùng đổi date ở main area
            if start_date_main != start_date or end_date_main != end_date:
                start_date = start_date_main
                end_date = end_date_main
                mask = (df_process["_parsed_date"].dt.date >= start_date) & (df_process["_parsed_date"].dt.date <= end_date)
                df_filtered = df_process[mask].copy()

            st.markdown(t("result_title"))
            if chon_nv:
                df_filtered = df_filtered[df_filtered["_nv_label"].isin(chon_nv)]

            @st.fragment
            def render_interactive_dashboard(df_base):
                df_filtered = df_base.copy()
                editor_key = f"data_editor_ot_{st.session_state.get('editor_key_counter', 0)}"

                # Áp dụng Giờ HC thủ công
                def apply_manual_hc(row):
                    ma = str(row[m['ma_nv']]).strip().upper()
                    if ma.endswith('.0'): ma = ma[:-2]
                    ngay = row["_parsed_date"].strftime('%d/%m/%Y')
                    if "manual_hc" in st.session_state and (ma, ngay) in st.session_state.manual_hc:
                        return min(8.0, float(st.session_state.manual_hc[(ma, ngay)]))
                    return min(8.0, float(row["Giờ hành chính"]))
                df_filtered["Giờ hành chính"] = df_filtered.apply(apply_manual_hc, axis=1)

                # Áp dụng OT thủ công
                def apply_manual_ot(row):
                    ma = str(row[m['ma_nv']]).strip().upper()
                    if ma.endswith('.0'): ma = ma[:-2]
                    ngay = row["_parsed_date"].strftime('%d/%m/%Y')
                    if "manual_ot" in st.session_state and (ma, ngay) in st.session_state.manual_ot:
                        return float(st.session_state.manual_ot[(ma, ngay)])
                    return row["Giờ OT"]
                df_filtered["Giờ OT"] = df_filtered.apply(apply_manual_ot, axis=1)

                # Tính lại Tổng giờ = Giờ HC + Giờ OT (Chỉ áp dụng nếu không phải lỗi -1, trừ khi có sửa thủ công)
                def apply_manual_total(row):
                    ma = str(row[m['ma_nv']]).strip().upper()
                    if ma.endswith('.0'): ma = ma[:-2]
                    ngay = row["_parsed_date"].strftime('%d/%m/%Y')
                    
                    if "manual_total" in st.session_state and (ma, ngay) in st.session_state.manual_total:
                        return float(st.session_state.manual_total[(ma, ngay)])

                    has_hc_override = "manual_hc" in st.session_state and (ma, ngay) in st.session_state.manual_hc
                    has_ot_override = "manual_ot" in st.session_state and (ma, ngay) in st.session_state.manual_ot
                    
                    if row["Số giờ làm thực tế"] == -1 and not (has_hc_override or has_ot_override):
                        return -1.0
                        
                    ot = float(row["Giờ OT"]) if pd.notna(row["Giờ OT"]) and str(row["Giờ OT"]).strip() != "" else 0.0
                    hc = float(row["Giờ hành chính"]) if pd.notna(row["Giờ hành chính"]) and str(row["Giờ hành chính"]).strip() != "" else 0.0
                    return hc + ot
                df_filtered["Số giờ làm thực tế"] = df_filtered.apply(apply_manual_total, axis=1)

                # Kiểm tra lỗi check-out TRƯỚC khi replace -1 → 0
                has_checkout_error = (df_filtered["Số giờ làm thực tế"] == -1).any()
                if has_checkout_error:
                    st.warning("⚠️ Phát hiện một số dòng có giờ ra sớm hơn giờ vào (lỗi check-out), đã được tính là 0 giờ.")

                def check_anomaly(row):
                    ma = str(row[m['ma_nv']]).strip().upper()
                    if ma.endswith('.0'): ma = ma[:-2]
                    ngay = row["_parsed_date"].strftime('%d/%m/%Y')
                    
                    if "manual_notes" in st.session_state and (ma, ngay) in st.session_state.manual_notes:
                        return st.session_state.manual_notes[(ma, ngay)]

                    vao = row[m['gio_vao']]
                    ra = row[m['gio_ra']]
                    is_wd = is_workday_func(row["_parsed_date"])
                    thu = row["_parsed_date"].weekday()  # 5=T7, 6=CN
                    has_ot_override = "manual_ot" in st.session_state and (ma, ngay) in st.session_state.manual_ot
                    has_hc_override = "manual_hc" in st.session_state and (ma, ngay) in st.session_state.manual_hc
                    has_ot_reason = "manual_ot_reason" in st.session_state and (ma, ngay) in st.session_state.manual_ot_reason
                    has_leave = st.session_state.get("manual_leave", {}).get((ma, ngay), False)
                    has_gps = "gps_synced" in st.session_state and (ma, ngay) in st.session_state.gps_synced

                    if has_gps:
                        return "📍 Công tác" if st.session_state.lang == 'vi' else "📍 \u51fa\u5f35"

                    if has_ot_override or has_hc_override or has_ot_reason:
                        return ""

                    try:
                        vao_trong = pd.isna(vao) or str(vao).strip().lower() in ['', 'nan', 'none', 'nat']
                    except (TypeError, ValueError):
                        vao_trong = False
                    try:
                        ra_trong = pd.isna(ra) or str(ra).strip().lower() in ['', 'nan', 'none', 'nat']
                    except (TypeError, ValueError):
                        ra_trong = False

                    notes = []
                    is_ja = st.session_state.lang == 'ja'

                    if has_leave:
                        d_check = row["_parsed_date"].date() if hasattr(row["_parsed_date"], 'date') else row["_parsed_date"]
                        if is_wd and vao_trong and ra_trong and is_last_saturday_of_month(d_check):
                            notes.append("🔴 \u5fc5\u9808\u571f\u66dc\u65e5\u6b20\u52e4" if is_ja else "🔴 Vắng Thứ 7 bắt buộc")
                        return " | ".join(notes) if notes else ""

                    if not is_wd and vao_trong and ra_trong:
                        return " | ".join(notes) if notes else ""

                    if vao_trong and not ra_trong: notes.append("⚠️ \u51fa\u52e4\u6253\u523b\u5fd8\u308c" if is_ja else "⚠️ Thiếu giờ vào")
                    elif ra_trong and not vao_trong: notes.append("⚠️ \u9000\u52e4\u6253\u523b\u5fd8\u308c" if is_ja else "⚠️ Thiếu giờ ra")
                    elif vao_trong and ra_trong:
                        if is_wd:
                            d_check = row["_parsed_date"].date() if hasattr(row["_parsed_date"], 'date') else row["_parsed_date"]
                            if is_last_saturday_of_month(d_check):
                                notes.append("🔴 \u5fc5\u9808\u571f\u66dc\u65e5\u6b20\u52e4" if is_ja else "🔴 Vắng Thứ 7 bắt buộc")
                            else:
                                notes.append("🔴 \u7121\u65ad\u6b20\u52e4" if is_ja else "🔴 Nghỉ không phép")
                    elif row["Số giờ làm thực tế"] == -1: notes.append("🟣 \u9000\u52e4\u30a8\u30e9\u30fc" if is_ja else "🟣 Lỗi check-out")
                    elif 0 < float(row["Số giờ làm thực tế"]) < 4: notes.append("🟠 \u5b9f\u50cd\u4e0d\u8db3 (< 4h)" if is_ja else "🟠 Làm thiếu giờ (< 4h)")

                    if has_leave and not (vao_trong and ra_trong):
                        notes.append("🟢 \u6709\u7d66\u4f11\u6687" if is_ja else "🟢 Nghỉ có phép")

                    if row.get("_is_chieu", False):
                        notes.append("🔵 \u5348\u5f8c\u51fa\u52e4" if is_ja else "🔵 Làm ca chiều")

                    return " | ".join(notes)
                df_filtered["Ghi chú"] = df_filtered.apply(check_anomaly, axis=1)
                df_filtered["Số giờ làm thực tế"] = pd.to_numeric(df_filtered["Số giờ làm thực tế"].replace(-1, 0), errors='coerce').fillna(0.0).round(1)
                df_filtered["Giờ hành chính"] = pd.to_numeric(df_filtered["Giờ hành chính"], errors='coerce').fillna(0.0).round(1)
                df_filtered["Giờ OT"] = pd.to_numeric(df_filtered["Giờ OT"], errors='coerce').fillna(0.0).round(1)

                def get_has_leave(row):
                    ma = str(row[m['ma_nv']]).strip().upper()
                    if ma.endswith('.0'): ma = ma[:-2]
                    ngay = row["_parsed_date"].strftime('%d/%m/%Y')
                    return st.session_state.get("manual_leave", {}).get((ma, ngay), False)

                df_filtered["Có phép"] = df_filtered.apply(get_has_leave, axis=1)

                def get_ot_reason(row):
                    ma = str(row[m['ma_nv']]).strip().upper()
                    if ma.endswith('.0'): ma = ma[:-2]
                    ngay = row["_parsed_date"].strftime('%d/%m/%Y')
                    return st.session_state.get("manual_ot_reason", {}).get((ma, ngay), "")

                df_filtered["Lý do tăng ca"] = df_filtered.apply(get_ot_reason, axis=1)

                with st.spinner("⏳ Đang tổng hợp kết quả..."):
                    total_rows = len(df_filtered)
                    total_emps = df_filtered[m['ma_nv']].nunique()
                    df_numeric = df_filtered[pd.to_numeric(df_filtered['Số giờ làm thực tế'], errors='coerce').notnull()]
                    total_hours = df_numeric['Số giờ làm thực tế'].sum() if not df_numeric.empty else 0
                    ngay_nghi = int((df_filtered["Số giờ làm thực tế"] == 0).sum())

                    # Bỏ thanh tổng quan kỳ công theo yêu cầu
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric(t("total_days"), working_days, help=t("total_days_help"))
                    c2.metric(t("num_employees"), total_emps)
                    c3.metric(t("total_hours"), f"{format_gio_lam(total_hours)} {t('hours_unit')}")
                    c4.metric(t("total_days_off"), ngay_nghi)

                    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

                    lang = st.session_state.lang
                    t_data = get_data_t(lang)
                    
                    chuc_vu_vals = [t_data(x) for x in (df_filtered[m['chuc_vu']].values if 'chuc_vu' in m and m['chuc_vu'] in df_filtered else [""] * len(df_filtered))]
                    phong_ban_vals = [t_data(x) for x in (df_filtered[m['phong_ban']].values if 'phong_ban' in m and m['phong_ban'] in df_filtered else [""] * len(df_filtered))]
                    ten_nv_vals = [translate_name(x, lang) for x in df_filtered[m['ten_nv']].values]

                    weekday_map_vi = {0:'Hai',1:'Ba',2:'Tư',3:'Năm',4:'Sáu',5:'Bảy',6:'CN'}
                    weekday_map_ja = {0:'\u6708',1:'\u706b',2:'\u6c34',3:'\u6728',4:'\u91d1',5:'\u571f',6:'\u65e5'}
                    weekday_map = weekday_map_ja if lang == 'ja' else weekday_map_vi

                    def format_time_str(val):
                        if pd.isna(val) or str(val).strip().lower() in ['nan', 'none', 'nat', '']:
                            return ""
                        if hasattr(val, 'strftime'):
                            return val.strftime('%H:%M')
                        return str(val).strip()

                    gio_vao_vals = [format_time_str(x) for x in (df_filtered[m['gio_vao']].values if 'gio_vao' in m and m['gio_vao'] in df_filtered else [""] * len(df_filtered))]
                    gio_ra_vals = [format_time_str(x) for x in (df_filtered[m['gio_ra']].values if 'gio_ra' in m and m['gio_ra'] in df_filtered else [""] * len(df_filtered))]

                    df_result_ui = pd.DataFrame({
                        "STT": range(1, len(df_filtered) + 1),
                        "Mã NV": df_filtered[m['ma_nv']].values,
                        "Tên nhân viên": ten_nv_vals,
                        "Chức vụ": chuc_vu_vals,
                        "Phòng ban": phong_ban_vals,
                        "Thứ": df_filtered["_parsed_date"].dt.weekday.map(weekday_map).values,
                        "Ngày": df_filtered["_parsed_date"].dt.strftime('%Y/%m/%d' if lang == 'ja' else '%d/%m/%Y').values,
                        "Giờ vào": gio_vao_vals,
                        "Giờ ra": gio_ra_vals,
                        "Giờ làm thực tế": df_filtered["Giờ hành chính"].values,
                        "OT": df_filtered["Giờ OT"].values,
                        "Tổng giờ": df_filtered["Số giờ làm thực tế"].values,
                        "Lý do tăng ca": df_filtered["Lý do tăng ca"].values,
                        "Ghi chú": df_filtered["Ghi chú"].values
                    })
                    st.session_state['df_result'] = df_result_ui
                    st.session_state['df_filtered_for_chat'] = df_filtered.copy(deep=True)

                    def get_loai_ngay(row):
                        try:
                            ngay = row["_parsed_date"].date() if hasattr(row["_parsed_date"], 'date') else pd.to_datetime(row["_parsed_date"]).date()
                            ngay_str = ngay.strftime('%d/%m/%Y')
                        except Exception as e: logger.warning(f"Lỗi: {e}", exc_info=True); return 'binh_thuong'
                            
                        gio_vao_raw = row.get(m['gio_vao'], pd.NA)
                        gio_ra_raw = row.get(m['gio_ra'], pd.NA)
                        ma = str(row[m['ma_nv']]).strip().upper()
                        if ma.endswith('.0'): ma = ma[:-2]
                        has_leave = st.session_state.get("manual_leave", {}).get((ma, ngay_str), False)
                        
                        thu = ngay.weekday()  # 5=T7, 6=CN
                        is_wd = is_workday_func(ngay)
                        vao_trong = pd.isna(gio_vao_raw) or str(gio_vao_raw).strip().lower() in ['', 'nan', 'none', 'nat']
                        ra_trong = pd.isna(gio_ra_raw) or str(gio_ra_raw).strip().lower() in ['', 'nan', 'none', 'nat']

                        if not is_wd and vao_trong and ra_trong:
                            return 'cuoi_tuan'   # T7 / CN không có dữ liệu

                        if is_wd and vao_trong and ra_trong:
                            if has_leave:
                                return 'nghi_co_phep'
                            return 'nghi_khong_phep'

                        return 'binh_thuong'

                    df_result_ui["_loai"] = df_filtered.apply(get_loai_ngay, axis=1).values

                    def style_row(row):
                        loai = df_result_ui.loc[row.name, "_loai"]
                        ngay_str = df_result_ui.loc[row.name, "Ngày"]
                        try:
                            fmt = '%Y/%m/%d' if lang == 'ja' else '%d/%m/%Y'
                            d_obj_row = datetime.datetime.strptime(ngay_str, fmt).date()
                            thu = d_obj_row.weekday()
                            is_weekend = (thu in [5, 6]) and not (thu == 5 and is_last_saturday_of_month(d_obj_row))
                        except Exception as e: logger.warning(f"Lỗi: {e}", exc_info=True); is_weekend = False

                        styles = [""] * len(row)
                        if is_weekend:
                            styles = ["background-color: #F1F5F9"] * len(row)
                            
                        idx_gio = list(row.index).index("Tổng giờ")
                        idx_ghi_chu = list(row.index).index("Ghi chú")
                        
                        if loai == 'cuoi_tuan':
                            styles[idx_gio] = "background-color: #F1F5F9; color: #64748B"
                        elif loai == 'nghi_khong_phep':
                            styles = ["background-color: #FEE2E2; color: #991B1B"] * len(row)
                            styles[idx_gio] = "background-color: #FEE2E2; color: #991B1B; font-weight: 600"
                        elif loai == 'nghi_co_phep':
                            styles[idx_gio] = "background-color: #F1F5F9; color: #0EA5E9; font-weight: 600" if is_weekend else "color: #0EA5E9; font-weight: 600"
                        else:
                            styles[idx_gio] = "background-color: #F1F5F9; color: #0EA5E9; font-weight: 600" if is_weekend else "color: #0EA5E9; font-weight: 600"
                            
                        val_str = str(row["Ghi chú"])
                        if "Nghỉ không phép" in val_str or "Vắng Thứ 7 bắt buộc" in val_str or "\u7121\u65ad\u6b20\u52e4" in val_str or "\u5fc5\u9808\u571f\u66dc\u65e5\u6b20\u52e4" in val_str:
                            styles[idx_ghi_chu] = 'background-color: #FEE2E2; color: #991B1B; font-weight: 600'
                        elif "Lỗi check-out" in val_str or "\u9000\u52e4\u30a8\u30e9\u30fc" in val_str:
                            styles[idx_ghi_chu] = 'background-color: #F3E8FF; color: #6B21A8; font-weight: 500'
                        elif "Thiếu giờ vào" in val_str or "Thiếu giờ ra" in val_str or "Làm thiếu giờ" in val_str or "\u51fa\u52e4\u6253\u523b\u5fd8\u308c" in val_str or "\u9000\u52e4\u6253\u523b\u5fd8\u308c" in val_str or "\u5b9f\u50cd\u4e0d\u8db3" in val_str:
                            styles[idx_ghi_chu] = 'background-color: #FEF3C7; color: #92400E; font-weight: 500'
                        elif "OT thủ công" in val_str or "\u624b\u52d5OT" in val_str:
                            styles[idx_ghi_chu] = 'background-color: #E0F2FE; color: #0369A1; font-weight: 500'
                        elif "ca chiều" in val_str or "\u5348\u5f8c\u51fa\u52e4" in val_str:
                            styles[idx_ghi_chu] = 'background-color: #F8FAFC; color: #475569; font-weight: 500'

                        return styles

                    # Bỏ thanh kết quả chi tiết theo yêu cầu
                    df_display = df_result_ui.drop(columns=["_loai"])
                    # Do not apply fmt_num to df_display here, keep as floats for data_editor
                    def apply_fmt_num(df):
                        df_out = df.copy()
                        def fmt_num(x):
                            try:
                                if pd.isna(x) or str(x).strip() == '': return ''
                                v = round(float(x), 1)
                                return f"{int(v)}" if v.is_integer() else f"{v}"
                            except: return x
                        for col in ['Giờ làm thực tế', 'OT', 'Tổng giờ']:
                            if col in df_out.columns:
                                df_out[col] = df_out[col].apply(fmt_num)
                        return df_out

                    tab_table, tab_chart = st.tabs([t("tab_table"), t("tab_chart")])
                    with tab_table:
                        import hashlib
                        try:
                            df_hash = hashlib.md5(pd.util.hash_pandas_object(df_display).values).hexdigest()
                        except:
                            df_hash = "fallback"
                        edit_mode = st.toggle("✏️ Bật Chế độ chỉnh sửa" if lang == 'vi' else "✏️ 編集モードを有効にする", value=False)
                        if edit_mode:
                            st.markdown("<p style='font-size: 13px; color: #64748B;'>💡 Nhấp đúp vào ô để chỉnh sửa. Dữ liệu tự động đồng bộ khi xuất file/lưu DB.</p>" if lang == 'vi' else "<p style='font-size: 13px; color: #64748B;'>💡 セルをダブルクリックして編集します。データはファイル出力／DB保存時に自動同期されます。</p>", unsafe_allow_html=True)
                            editor_key = f"editor_{df_hash}"
                            
                            col_cfg = {"version": None}
                            if lang == 'ja':
                                ja_col_labels = {
                                    "STT": t("mos_col_stt"),
                                    "Mã NV": t("export_col_emp_code"),
                                    "Tên nhân viên": t("export_col_emp_name"),
                                    "Chức vụ": t("emp_position"),
                                    "Phòng ban": t("emp_dept"),
                                    "Thứ": t("export_col_weekday"),
                                    "Ngày": t("export_col_date"),
                                    "Giờ vào": t("export_col_in"),
                                    "Giờ ra": t("export_col_out"),
                                    "Giờ làm thực tế": t("export_col_actual_hours"),
                                    "OT": t("export_col_ot"),
                                    "Tổng giờ": t("export_col_total_hours"),
                                    "Lý do tăng ca": t("export_col_ot_reason"),
                                    "Ghi chú": t("export_col_note")
                                }
                                for col_k, ja_lbl in ja_col_labels.items():
                                    if col_k in df_display.columns:
                                        if col_k in ['Giờ làm thực tế', 'OT', 'Tổng giờ']:
                                            df_display[col_k] = pd.to_numeric(df_display[col_k], errors='coerce').fillna(0.0).astype(float)
                                            col_cfg[col_k] = st.column_config.NumberColumn(label=ja_lbl, format="%g", step=0.01)
                                        else:
                                            col_cfg[col_k] = st.column_config.Column(label=ja_lbl)
                            else:
                                for col in ['Giờ làm thực tế', 'OT', 'Tổng giờ']:
                                    if col in df_display.columns:
                                        df_display[col] = pd.to_numeric(df_display[col], errors='coerce').fillna(0.0).astype(float)
                                        col_cfg[col] = st.column_config.NumberColumn(format="%g", step=0.01)
                                    
                            if editor_key in st.session_state and isinstance(st.session_state[editor_key], dict):
                                edited_rows = st.session_state[editor_key].get("edited_rows", {})
                                for r_idx_str, edits in edited_rows.items():
                                    try:
                                        r_idx = int(r_idx_str)
                                        if r_idx in df_display.index:
                                            if "Giờ làm thực tế" in edits or "OT" in edits:
                                                hc_val = edits.get("Giờ làm thực tế", df_display.loc[r_idx, "Giờ làm thực tế"])
                                                ot_val = edits.get("OT", df_display.loc[r_idx, "OT"])
                                                try:
                                                    hc_f = float(hc_val) if pd.notna(hc_val) and str(hc_val).strip() != "" else 0.0
                                                    ot_f = float(ot_val) if pd.notna(ot_val) and str(ot_val).strip() != "" else 0.0
                                                    edits["Tổng giờ"] = round(hc_f + ot_f, 1)
                                                except ValueError:
                                                    pass
                                            elif "Tổng giờ" in edits:
                                                tot_val = edits["Tổng giờ"]
                                                ot_val = edits.get("OT", df_display.loc[r_idx, "OT"])
                                                try:
                                                    tot_f = float(tot_val) if pd.notna(tot_val) and str(tot_val).strip() != "" else 0.0
                                                    ot_f = float(ot_val) if pd.notna(ot_val) and str(ot_val).strip() != "" else 0.0
                                                    edits["Giờ làm thực tế"] = min(8.0, max(0.0, round(tot_f - ot_f, 1)))
                                                except ValueError:
                                                    pass
                                    except Exception:
                                        pass

                            edited_df_display = st.data_editor(df_display, use_container_width=True, hide_index=True, height=500, column_config=col_cfg, key=f"editor_{df_hash}")
                            
                            # Sync edits back to df_filtered
                            if "OT" in edited_df_display.columns:
                                df_filtered["Giờ OT"] = pd.to_numeric(edited_df_display["OT"], errors='coerce').fillna(0).round(1)
                            if "Giờ làm thực tế" in edited_df_display.columns:
                                df_filtered["Giờ hành chính"] = pd.to_numeric(edited_df_display["Giờ làm thực tế"], errors='coerce').fillna(0).round(1).apply(lambda x: min(8.0, float(x)))
                            if "Tổng giờ" in edited_df_display.columns:
                                df_filtered["Số giờ làm thực tế"] = pd.to_numeric(edited_df_display["Tổng giờ"], errors='coerce').fillna(0).round(1)
                            if "Giờ vào" in edited_df_display.columns:
                                df_filtered["Giờ vào"] = edited_df_display["Giờ vào"].values
                                if 'gio_vao' in m and m['gio_vao'] in df_filtered.columns:
                                    df_filtered[m['gio_vao']] = edited_df_display["Giờ vào"].values
                            if "Giờ ra" in edited_df_display.columns:
                                df_filtered["Giờ ra"] = edited_df_display["Giờ ra"].values
                                if 'gio_ra' in m and m['gio_ra'] in df_filtered.columns:
                                    df_filtered[m['gio_ra']] = edited_df_display["Giờ ra"].values
                            if "Ghi chú" in edited_df_display.columns:
                                df_filtered["Ghi chú"] = edited_df_display["Ghi chú"].values
                            if "Lý do tăng ca" in edited_df_display.columns:
                                df_filtered["Lý do tăng ca"] = edited_df_display["Lý do tăng ca"].values
                            if "Có phép" in edited_df_display.columns:
                                df_filtered["Có phép"] = edited_df_display["Có phép"].values
                                
                            col_save1, col_save2 = st.columns([3, 1])
                            with col_save1:
                                if st.button("💾 Lưu thay đổi" if st.session_state.get('lang', 'vi') == 'vi' else "💾 \u5909\u66f4\u3092\u4fdd\u5b58", type="primary", use_container_width=True):
                                    st.session_state['undo_db_backup'] = st.session_state.get('df_filtered_for_chat').copy(deep=True)
                                    with st.spinner(t("spinner_save")):
                                        try:
                                            # Tự động tính lại Tổng giờ trước khi lưu DB
                                            df_filtered["Số giờ làm thực tế"] = (pd.to_numeric(df_filtered["Giờ hành chính"], errors='coerce').fillna(0) + pd.to_numeric(df_filtered["Giờ OT"], errors='coerce').fillna(0)).round(1)
                                            
                                            conflicts = save_to_db(df_filtered, m)
                                            if conflicts:
                                                msg_conflict = f"❌ XUNG ĐỘT DỮ LIỆU: Các nhân viên sau đã bị người khác thay đổi trong lúc bạn đang sửa: {', '.join(conflicts)}. Vui lòng tải lại trang!" if lang == 'vi' else f"❌ データ競合: 編集前に他者によって更新された社員がいます: {', '.join(conflicts)}。ページを再読み込みしてください！"
                                                st.error(msg_conflict)
                                            else:
                                                st.success(t("success_save"))
                                                # Reload df_raw to prevent UI from reverting
                                                if st.session_state.get('last_uploaded') == "Database":
                                                    conn = sqlite3.connect(DB_FILE)
                                                    st.session_state.df_raw = pd.read_sql_query("SELECT * FROM records", conn)
                                                    conn.close()
                                                else:
                                                    if st.session_state.get('df_raw') is not None:
                                                        import numpy as np
                                                        df_raw = st.session_state.df_raw
                                                        
                                                        # Thêm cột vào df_raw nếu chưa có
                                                        if 'tong_gio' not in df_raw.columns:
                                                            df_raw['tong_gio'] = np.nan
                                                        if 'ot_manual' not in df_raw.columns:
                                                            df_raw['ot_manual'] = np.nan
                                                        if 'hc_manual' not in df_raw.columns:
                                                            df_raw['hc_manual'] = np.nan
                                                            
                                                        # Cập nhật giá trị chính xác theo Mã NV và Ngày làm việc
                                                        df_raw_dates = df_raw[m['ngay']].apply(clean_date)
                                                        df_raw_mas = df_raw[m['ma_nv']].astype(str).str.strip()
                                                        
                                                        for _, f_row in df_filtered.iterrows():
                                                            f_ma = str(f_row[m['ma_nv']]).strip()
                                                            f_date = f_row["_parsed_date"]
                                                            match_mask = (df_raw_mas == f_ma) & (df_raw_dates == f_date)
                                                            if match_mask.any():
                                                                df_raw.loc[match_mask, 'tong_gio'] = f_row["Số giờ làm thực tế"]
                                                                df_raw.loc[match_mask, 'ot_manual'] = f_row["Giờ OT"]
                                                                df_raw.loc[match_mask, 'hc_manual'] = f_row["Giờ hành chính"]
                                                                if 'gio_vao' in m and m['gio_vao'] in df_raw.columns:
                                                                    df_raw.loc[match_mask, m['gio_vao']] = f_row.get("Giờ vào", f_row.get(m['gio_vao']))
                                                                if 'gio_ra' in m and m['gio_ra'] in df_raw.columns:
                                                                    df_raw.loc[match_mask, m['gio_ra']] = f_row.get("Giờ ra", f_row.get(m['gio_ra']))
                                                                    
                                                        # Cập nhật mapping để lần process_data sau sẽ ưu tiên dùng dữ liệu đã lưu
                                                        st.session_state.mapping['tong_gio'] = 'tong_gio'
                                                        st.session_state.mapping['ot'] = 'ot_manual'
                                                        st.session_state.mapping['hc_manual'] = 'hc_manual'
                                                        
                                                        
                                                # Set flag to show success message after rerun
                                                st.session_state.save_success = True
                                                st.rerun()
                                        except Exception as e:
                                            logger.error(f"Lỗi khi lưu Database: {e}", exc_info=True)
                                            st.error(f"❌ Có lỗi xảy ra khi lưu Database: {e}")
                            with col_save2:
                                if st.session_state.get('undo_db_backup') is not None:
                                    if st.button("↩️ Hoàn tác" if lang == 'vi' else "↩️ 元に戻す", use_container_width=True):
                                        with st.spinner("Đang khôi phục..." if lang == 'vi' else "復元中..."):
                                            try:
                                                conflicts = save_to_db(st.session_state['undo_db_backup'], m)
                                                if conflicts:
                                                    msg_undo_conf = f"❌ XUNG ĐỘT DỮ LIỆU: Dữ liệu đã bị thay đổi bởi người khác: {', '.join(conflicts)}" if lang == 'vi' else f"❌ データ競合: データは他者によって変更されました: {', '.join(conflicts)}"
                                                    st.error(msg_undo_conf)
                                                else:
                                                    st.session_state['undo_db_backup'] = None
                                                    st.success("Đã khôi phục dữ liệu gốc!" if lang == 'vi' else "元のデータを復元しました！")
                                            except Exception as e:
                                                logger.error(f"Lỗi khi khôi phục Database: {e}", exc_info=True)
                                                st.error(f"❌ Lỗi khôi phục: {e}")
                                        import time
                                        time.sleep(0.5)
                                        st.rerun()
                        else:
                            df_display_styled = apply_fmt_num(df_display)
                            col_cfg_view = {}
                            if lang == 'ja':
                                ja_col_labels_view = {
                                    "STT": t("mos_col_stt"),
                                    "Mã NV": t("export_col_emp_code"),
                                    "Tên nhân viên": t("export_col_emp_name"),
                                    "Chức vụ": t("emp_position"),
                                    "Phòng ban": t("emp_dept"),
                                    "Thứ": t("export_col_weekday"),
                                    "Ngày": t("export_col_date"),
                                    "Giờ vào": t("export_col_in"),
                                    "Giờ ra": t("export_col_out"),
                                    "Giờ làm thực tế": t("export_col_actual_hours"),
                                    "OT": t("export_col_ot"),
                                    "Tổng giờ": t("export_col_total_hours"),
                                    "Lý do tăng ca": t("export_col_ot_reason"),
                                    "Ghi chú": t("export_col_note")
                                }
                                col_cfg_view = {k: st.column_config.Column(label=ja_col_labels_view[k]) for k in df_display.columns if k in ja_col_labels_view}
                            st.dataframe(df_display_styled.style.apply(style_row, axis=1), use_container_width=True, hide_index=True, height=500, column_config=col_cfg_view)
                    with tab_chart:
                        try:
                            import plotly.express as px
                            df_c = df_filtered.copy()
                            df_c['Giờ hành chính'] = pd.to_numeric(df_c['Giờ hành chính'], errors='coerce').fillna(0)
                            df_c['Giờ OT'] = pd.to_numeric(df_c['Giờ OT'], errors='coerce').fillna(0)
                            
                            is_vi = st.session_state.get('lang', 'vi') == 'vi'
                            col_chart1, col_chart2 = st.columns(2)
                            
                            with col_chart1:
                                df_emp = df_c.groupby(m['ten_nv'])[['Giờ hành chính', 'Giờ OT']].sum().reset_index()
                                df_emp = df_emp.rename(columns={'Giờ hành chính': t("auto_text_app_34"), 'Giờ OT': t("auto_text_app_35")})
                                fig1 = px.bar(
                                    df_emp, x=m['ten_nv'], y=[t("auto_text_app_34"), t("auto_text_app_35")],
                                    title=t("auto_text_app_38"),
                                    labels={"value": t("auto_text_app_39"), m['ten_nv']: t("auto_text_app_40")},
                                    barmode='stack', color_discrete_sequence=['#38BDF8', '#F43F5E']
                                )
                                fig1.update_layout(legend_title_text="", margin=dict(t=50, b=0, l=0, r=0))
                                st.plotly_chart(fig1, use_container_width=True)
                                
                            with col_chart2:
                                df_date = df_c.groupby('_parsed_date')[['Giờ hành chính', 'Giờ OT']].sum().reset_index()
                                df_date = df_date.rename(columns={'Giờ hành chính': t("auto_text_app_34"), 'Giờ OT': t("auto_text_app_35")})
                                fig2 = px.line(
                                    df_date, x='_parsed_date', y=[t("auto_text_app_34"), t("auto_text_app_35")],
                                    title=t("auto_text_app_45"),
                                    labels={"value": t("auto_text_app_39"), "_parsed_date": t("auto_text_app_47")},
                                    markers=True, color_discrete_sequence=['#38BDF8', '#F43F5E']
                                )
                                fig2.update_layout(legend_title_text="", margin=dict(t=50, b=0, l=0, r=0))
                                st.plotly_chart(fig2, use_container_width=True)
                        except Exception as e:
                            print(f"Lỗi vẽ biểu đồ: {e}")
                            st.info('Không thể vẽ biểu đồ do dữ liệu chưa phù hợp.' if st.session_state.get('lang', 'vi') == 'vi' else '\u30c7\u30fc\u30bf\u304c\u4e0d\u9069\u5207\u306a\u305f\u3081\u3001\u30c1\u30e3\u30fc\u30c8\u3092\u63cf\u753b\u3067\u304d\u307e\u305b\u3093\u3002')

                    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
                    st.markdown("## ⬇️ \u30d5\u30a1\u30a4\u30eb\u51fa\u529b\u3068\u4fdd\u5b58" if st.session_state.lang == 'ja' else "## ⬇️ Xuất file & Lưu trữ")
                    col_exp1, col_exp2, col_exp3 = st.columns([2, 1, 1])
                    
                    file_prefix = "\u30bf\u30a4\u30e0\u30ab\u30fc\u30c9" if st.session_state.lang == 'ja' else "Bang cong"
                    file_name_export = f"{file_prefix} {start_date.strftime('%d%m')}-{end_date.strftime('%d%m')}.xlsx"
                    
                    with col_exp1:
                        st.markdown(f"""<div style='background:#F0FDF8;border:0.5px solid #6EE7C0;border-radius:10px;padding:12px 16px;font-size:13px;color:#0F6E56'>
    📄 File: <b>{file_name_export}</b><br>
    <span style='color:#6B7280;font-size:12px'>{t("export_file_rows", rows=total_rows, emps=total_emps)}</span>
    </div>""", unsafe_allow_html=True)
                    with col_exp2:
                        if st.button(t("btn_save_db"), use_container_width=True):
                            with st.spinner(t("spinner_save")):
                                try:
                                    conflicts = save_to_db(df_filtered, m)
                                    if conflicts:
                                        st.error(f"❌ XUNG ĐỘT DỮ LIỆU: Các bản ghi sau đã bị thay đổi bởi người khác: {', '.join(conflicts)}")
                                    else:
                                        st.success(t("success_save"))
                                        if st.session_state.get('last_uploaded') == "Database":
                                            conn = sqlite3.connect(DB_FILE)
                                            st.session_state.df_raw = pd.read_sql_query("SELECT * FROM records", conn)
                                            conn.close()
                                except Exception as e:
                                    logger.error(f"Lỗi khi lưu Database (từ export): {e}", exc_info=True)
                                    st.error(f"❌ Có lỗi xảy ra khi lưu Database: {e}")
                    with col_exp3:
                        excel_data = None
                        try:
                            total_wd_tuple = calculate_working_days(start_date, end_date, st.session_state.custom_holidays, st.session_state.custom_workdays)
                            total_wd = total_wd_tuple[0]
                            with st.spinner("⏳ Đang kết xuất Bảng Chấm Công ra file Excel..."):
                                excel_data = export_excel_tong_hop(df_filtered, m, start_date, end_date, total_wd)
                        except Exception as e:
                            logger.error(f"Lỗi xuất Excel: {e}", exc_info=True)
                            st.error(f"❌ Lỗi xuất file: {e}")
                        if excel_data is not None:
                            st.download_button(
                                label=t("btn_download_excel"),
                                data=excel_data,
                                file_name=file_name_export,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                type="primary", use_container_width=True,
                            )
                        else:
                            st.button("⬇️ Tải Excel", disabled=True, use_container_width=True, help="Không có dữ liệu để xuất")

            st.session_state['current_df_filtered'] = df_filtered
            with st.spinner("⏳ Đang tính toán công số và đồng bộ Bảng chấm công chi tiết AI..."):
                render_interactive_dashboard(df_filtered)


# trigger reload
