import streamlit as st
from theme import get_theme
from log_config import logger
import pandas as pd
import sqlite3
from translations import get_t, translate_name
from db import DB_FILE

@st.cache_data(ttl=300)
def _get_db_gamification_field():
    import sqlite3
    import pandas as pd
    try:
        conn = sqlite3.connect(DB_FILE)
        df = pd.read_sql_query("SELECT ten_nv, COUNT(id) as total_checkin FROM field_checkins GROUP BY ma_nv, ten_nv ORDER BY total_checkin DESC LIMIT 3", conn)
        conn.close()
        return df
    except Exception as e:
        logger.warning(f"Error querying gamification field checkins: {e}")
        return pd.DataFrame(columns=['ten_nv', 'total_checkin'])

@st.cache_data(ttl=300)
def _get_db_gamification_records():
    import sqlite3
    import pandas as pd
    try:
        conn = sqlite3.connect(DB_FILE)
        df_ot = pd.read_sql_query("SELECT ten_nv, SUM(CAST(ot AS FLOAT)) as total_ot FROM records WHERE ot != '' AND ot != '0' AND ot != '0.0' GROUP BY ma_nv, ten_nv ORDER BY total_ot DESC LIMIT 3", conn)
        df_days = pd.read_sql_query("SELECT ten_nv, COUNT(DISTINCT ngay) as total_days FROM records GROUP BY ma_nv, ten_nv ORDER BY total_days DESC LIMIT 3", conn)
        conn.close()
        return df_ot, df_days
    except Exception as e:
        logger.warning(f"Error querying gamification records: {e}")
        return pd.DataFrame(columns=['ten_nv', 'total_ot']), pd.DataFrame(columns=['ten_nv', 'total_days'])

def render_gamification_dashboard():
    t = get_t(st.session_state.get('lang', 'vi'))
    is_vi = st.session_state.get('lang', 'vi') == 'vi'
    theme_mode = st.session_state.get('theme_mode', 'light')
    gT = get_theme(theme_mode)
    is_sepia = (theme_mode == 'sepia')

    # ─── Bảng theme trung tâm – dùng chung toàn file ──────────────────────
    T = {
        "bg_card":      "rgba(254,252,232,0.96)"  if is_sepia else "rgba(255,255,255,0.95)",
        "bg_others":    "rgba(120,53,15,0.04)"     if is_sepia else "rgba(0,0,0,0.02)",
        "text_h1":      "#78350F"                  if is_sepia else "#1E293B",
        "text_sub":     "#92400E"                  if is_sepia else "#64748B",
        "text_body":    "#78350F"                  if is_sepia else "#334155",
        "text_muted":   "#92400E"                  if is_sepia else "#64748B",
        "text_name":    "#451A03"                  if is_sepia else "#1E293B",
        "divider":      "rgba(120,53,15,0.15)"     if is_sepia else "rgba(0,0,0,0.1)",
    }
    # ───────────────────────────────────────────────────────────────────────

    title_text = "Top 3 Đóng Góp Xuất Sắc" if is_vi else "トップ3 優秀貢献社員"
    sub_text = "Hệ thống tự động vinh danh" if is_vi else "システムによる自動表彰"

    st.markdown(f"""
        <div style="text-align: center; margin-bottom: 25px;">
            <h2 style="color: {T['text_h1']}; margin: 0; font-size: 28px;">🏆 {title_text}</h2>
            <p style="color: {T['text_sub']}; margin: 5px 0 0 0; font-size: 14px;">✨ {sub_text} ✨</p>
        </div>
    """, unsafe_allow_html=True)

@st.cache_data(ttl=300)
def _process_gamification_df(df_curr):
    import pandas as pd
    ma_nv_col = df_curr.columns[0] if len(df_curr.columns) > 0 else "ma_nv"
    ten_nv_col = "Tên nhân viên" if "Tên nhân viên" in df_curr.columns else "ten_nv"
    ot_col = "Giờ OT" if "Giờ OT" in df_curr.columns else "ot"
    ngay_col = "Ngày" if "Ngày" in df_curr.columns else "ngay"
    
    df_curr_copy = df_curr.copy()
    df_curr_copy[ot_col] = pd.to_numeric(df_curr_copy[ot_col].astype(str).str.replace(',', '.').replace('', '0'), errors='coerce').fillna(0)
    df_ot = df_curr_copy[df_curr_copy[ot_col] > 0].groupby([ma_nv_col, ten_nv_col])[ot_col].sum().reset_index().rename(columns={ten_nv_col: 'ten_nv', ot_col: 'total_ot'}).sort_values('total_ot', ascending=False).head(3)
    df_days = df_curr_copy.groupby([ma_nv_col, ten_nv_col])[ngay_col].nunique().reset_index().rename(columns={ten_nv_col: 'ten_nv', ngay_col: 'total_days'}).sort_values('total_days', ascending=False).head(3)
    return df_ot, df_days

def render_gamification_dashboard():
    t = get_t(st.session_state.get('lang', 'vi'))
    is_vi = st.session_state.get('lang', 'vi') == 'vi'
    theme_mode = st.session_state.get('theme_mode', 'light')
    gT = get_theme(theme_mode)
    is_sepia = (theme_mode == 'sepia')

    # ─── Bảng theme trung tâm – dùng chung toàn file ──────────────────────
    T = {
        "bg_card":      "rgba(254,252,232,0.96)"  if is_sepia else "rgba(255,255,255,0.95)",
        "bg_others":    "rgba(120,53,15,0.04)"     if is_sepia else "rgba(0,0,0,0.02)",
        "text_h1":      "#78350F"                  if is_sepia else "#1E293B",
        "text_sub":     "#92400E"                  if is_sepia else "#64748B",
        "text_body":    "#78350F"                  if is_sepia else "#334155",
        "text_muted":   "#92400E"                  if is_sepia else "#64748B",
        "text_name":    "#451A03"                  if is_sepia else "#1E293B",
        "divider":      "rgba(120,53,15,0.15)"     if is_sepia else "rgba(0,0,0,0.1)",
    }
    # ───────────────────────────────────────────────────────────────────────

    # Tiêu đề và mô tả đã được chuyển vào thanh ngang rút gọn

    try:
        import pandas as pd
        df_field = _get_db_gamification_field()
        
        if 'current_df_filtered' in st.session_state and st.session_state.current_df_filtered is not None and not st.session_state.current_df_filtered.empty:
            df_ot, df_days = _process_gamification_df(st.session_state.current_df_filtered)
        else:
            df_ot, df_days = _get_db_gamification_records()
            
    except Exception as e:
        logger.warning(f"Error processing gamification data: {e}", exc_info=True)
        return

    # Tạo HTML hiển thị rút gọn (Thanh ngang góc phải)
    def get_top1_info(df, unit):
        if df.empty:
            return "?" if is_vi else "?", 0
        df = df.copy()
        df['ten_nv'] = [translate_name(str(x), st.session_state.get('lang', 'vi')) for x in df['ten_nv']]
        top_name = str(df.iloc[0]['ten_nv']).split()[-1] if df.iloc[0]['ten_nv'] else '?'
        val = round(df.iloc[0][unit], 1) if unit == 'total_ot' else df.iloc[0][unit]
        return top_name, val

    top_ot_name, val_ot = get_top1_info(df_ot, 'total_ot')
    top_days_name, val_days = get_top1_info(df_days, 'total_days')
    top_chk_name, val_chk = get_top1_info(df_field, 'total_checkin')
    
    unit_ot = "h"
    unit_days = "d" if not is_vi else "ng"
    unit_chk = "l" if is_vi else "回"

    st.markdown(f"""
        <div style="position: fixed; bottom: 20px; right: 20px; z-index: 999999; 
                    background: {T['bg_card']}; padding: 10px 16px; border-radius: 50px; 
                    box-shadow: 0 8px 25px rgba(0,0,0,0.12); display: flex; gap: 15px; 
                    border: 1px solid {T['divider']}; align-items: center; font-size: 13px; 
                    backdrop-filter: blur(10px); transition: transform 0.3s ease;">
            <div style="font-weight: 800; color: {T['text_h1']}; margin-right: 5px;">🏆 Vinh danh:</div>
            <div style="display: flex; gap: 12px;">
                <div style="color: #EC4899; font-weight: 600; display: flex; align-items: center; gap: 4px;" title="Chiến Thần OT">
                    🔥 <span>{top_ot_name} ({val_ot}{unit_ot})</span>
                </div>
                <div style="color: #EAB308; font-weight: 600; display: flex; align-items: center; gap: 4px;" title="Ong Chăm Chỉ">
                    🐝 <span>{top_days_name} ({val_days}{unit_days})</span>
                </div>
                <div style="color: #EC4899; font-weight: 600; display: flex; align-items: center; gap: 4px;" title="Thánh Check-in">
                    📍 <span>{top_chk_name} ({val_chk}{unit_chk})</span>
                </div>
            </div>
        </div>
    """, unsafe_allow_html=True)



if __name__ == '__main__':
    render_gamification_dashboard()
