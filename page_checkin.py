import streamlit as st
from theme import get_theme
from log_config import logger
import pandas as pd
import datetime
import sqlite3
from db import save_field_checkin, get_field_checkins, get_company_emp_options, get_company_emp_dict, DB_FILE
from translations import get_t, translate_name, translate_dia_diem

def render_checkin_page():
    t = get_t(st.session_state.get('lang', 'vi'))
    is_vi = st.session_state.get('lang', 'vi') == 'vi'
    
    title_text = "📱 Cổng Check-in GPS Hiện Trường" if is_vi else "📱 フィールド打刻 GPS"
    sub_text = "Dành cho Kỹ sư đi lắp đặt, bảo trì công trình tại nhà máy khách hàng" if is_vi else "顧客工場での現地設置・保守作業員向け"
    st.markdown(f"""
        <div style="margin-bottom: 5px;">
            <h1 style='margin:0; color:#0F172A; font-size:28px;'>{title_text}</h1>
            <p style='color:#64748B; margin:4px 0 0 0;'>{sub_text}</p>
        </div>
    """, unsafe_allow_html=True)
    
    theme_mode = st.session_state.get('theme_mode', 'light')
    gT = get_theme(theme_mode)
    is_sepia = (theme_mode == 'sepia')

    # ── Bảng theme trung tâm ──────────────────────────────────────────
    T = {
        "bg_page":    gT['bg_app'],
        "bg_form":    gT['bg_card'],
        "bg_expander":gT['bg_card_hover'],
        "border":     gT['border'],
        "border_dash":gT['primary'],
        "text_label": gT['text_secondary'],
    }
    # ─────────────────────────────────────────────────────────────────

    st.markdown(f"""
        <style>
        /* Chỉ đổi màu nền nếu bật chế độ Sepia bảo vệ mắt, nếu không thì giữ nguyên hình nền gốc */
        {f'''.stApp, .main, [data-testid="stAppViewContainer"] {{
            background: {T["bg_page"]} !important;
        }}''' if is_sepia else ''}

        /* Chỉnh lại màu của toàn bộ bảng (form và lịch sử) bên cạnh nhau */
        .st-key-field_form_container, .st-key-history_table_container {{
            margin-top: -10px !important;
            background: {T['bg_form']} !important;
            border-radius: 20px !important;
            padding: 25px 25px !important;
            box-shadow: 0 15px 45px rgba(0, 0, 0, 0.08), 0 4px 12px rgba(0,0,0,0.04) !important;
            border: 1px solid {T['border']} !important;
        }}
        .st-key-field_form_container [data-testid="stExpander"], .st-key-history_table_container [data-testid="stExpander"] {{
            background: {T['bg_expander']} !important;
            border-radius: 12px !important;
            border: 1.5px dashed {T['border_dash']} !important;
            margin-bottom: 20px !important;
        }}
        .st-key-field_form_container label {{
            font-weight: 600 !important;
            color: {T['text_label']} !important;
            font-size: 14px !important;
            margin-bottom: 5px !important;
        }}
        .st-key-btn_submit_field > button {{
            background: linear-gradient(135deg, #EC4899 0%, #DB2777 100%) !important;
            color: white !important;
            border: none !important;
            border-radius: 12px !important;
            font-weight: 700 !important;
            font-size: 16px !important;
            height: 52px !important;
            margin-top: 15px !important;
            box-shadow: 0 8px 20px rgba(236, 72, 153, 0.25) !important;
            transition: all 0.3s ease !important;
        }}
        .st-key-btn_submit_field > button:hover {{
            transform: translateY(-2px) !important;
            box-shadow: 0 12px 28px rgba(236, 72, 153, 0.4) !important;
            background: linear-gradient(135deg, #DB2777 0%, #0369A1 100%) !important;
        }}
        </style>
    """, unsafe_allow_html=True)

    col_form, col_hist = st.columns([5, 7], gap="large")
    with col_form:
        with st.container(key="field_form_container"):
            is_vi = (st.session_state.get('lang', 'vi') == 'vi')
            st.markdown("<h2 style='text-align: center; color: #0F172A; font-weight: 800; margin-bottom: 25px; font-size: 24px; letter-spacing: -0.5px;'>" + ("📍 Xác Nhận Hiện Trường" if is_vi else "📍 現地打刻の確認") + "</h2>", unsafe_allow_html=True)
            if "manual_emps" not in st.session_state:
                st.session_state.manual_emps = []

            emp_options = get_company_emp_options(st.session_state.lang)

            sel_emp_gps = st.selectbox("👤 Chọn nhân viên:" if is_vi else "👤 担当者を選択:", emp_options, key="sb_sel_emp_gps")
            ma_nv = sel_emp_gps.split(" - ")[0].strip() if sel_emp_gps else ""
            ten_nv = sel_emp_gps.split(" - ")[1].strip() if (sel_emp_gps and " - " in sel_emp_gps) else ""
            
            with st.expander("➕ Thêm nhân sự mới (nếu chưa có trong danh sách)" if is_vi else "➕ 新規メンバー追加 (リストにない場合)"):
                with st.form("form_add_emp_gps"):
                    col_m1, col_m2 = st.columns([1, 2])
                    new_ma_gps = col_m1.text_input("Mã NV:" if is_vi else "社員ID:", placeholder="VD: NV099" if is_vi else "例: NV099")
                    new_ten_gps = col_m2.text_input("Tên nhân viên:" if is_vi else "氏名:", placeholder="VD: Nguyễn Văn A" if is_vi else "例: グエン・ヴァン・A")
                    if st.form_submit_button("➕ Thêm vào danh sách" if is_vi else "➕ リストに追加", type="secondary", use_container_width=True):
                        if new_ma_gps and new_ten_gps:
                            m_clean = new_ma_gps.strip().upper()
                            t_clean = new_ten_gps.strip()
                            st.session_state.manual_emps.append({
                                "ma": m_clean,
                                "ten": t_clean,
                                "cv": "",
                                "pb": ""
                            })
                            try:
                                conn = sqlite3.connect(DB_FILE)
                                conn.execute("INSERT OR REPLACE INTO employees VALUES (?, ?, ?, ?)", (m_clean, t_clean, "", ""))
                                conn.commit(); conn.close()
                            except Exception as e: logger.warning(f"Lỗi: {e}", exc_info=True); pass
                            st.success(f"✅ Đã thêm NV {m_clean} - {t_clean}!" if is_vi else f"✅ {m_clean} を追加しました！")
                            st.rerun()
                        else:
                            st.error("⚠️ Vui lòng nhập Mã NV và Tên NV!" if is_vi else "⚠️ 社員IDと氏名を入力してください！")
            
            st.markdown("<hr style='margin: 10px 0 25px 0; border: none; border-top: 1px dashed #E2E8F0;'/>", unsafe_allow_html=True)
            
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                ngay_bat_dau = st.date_input("📅 Từ ngày:" if is_vi else "📅 開始日:", value=datetime.date.today(), key="gps_d_start")
            with col_d2:
                ngay_ket_thuc = st.date_input("📅 Đến ngày:" if is_vi else "📅 終了日:", value=datetime.date.today(), key="gps_d_end")

            col_t1, col_t2 = st.columns([3, 2])
            with col_t1:
                loai = st.radio("🔴 Loại check-in:" if is_vi else "🔴 打刻種別:", ["🟢 Vào ca (Check-in)", "🟣 Tan ca (Check-out)"] if is_vi else ["🟢 出勤", "🟣 退勤"], horizontal=True)
            with col_t2:
                is_vao = ('Vào ca' in loai or '出勤' in loai)
                default_time = datetime.time(8, 0) if is_vao else datetime.time(17, 0)
                gio_checkin = st.time_input("⏰ Giờ check-in:" if is_vi else "⏰ 打刻時刻:", value=default_time, key=f"gps_t_{'in' if is_vao else 'out'}")

            dia_diem = st.text_input("📍 Địa điểm hiện trường:" if is_vi else "📍 現地場所:", placeholder="VD: Nhà máy Canon Bắc Ninh" if is_vi else "例: キヤノンバクニン工場")
            ghi_chu = st.text_area("📝 Chi tiết công việc:" if is_vi else "📝 作業詳細:", placeholder="VD: Kiểm tra cảm biến tại dây chuyền..." if is_vi else "例: ラインセンサーの点検...")
            
            if st.button("✅ XÁC NHẬN CHẤM CÔNG HIỆN TRƯỜNG" if is_vi else "✅ 現地打刻を確定する", type="primary", use_container_width=True, key="btn_submit_field"):
                if not ma_nv or not ten_nv or not dia_diem:
                    st.error("⚠️ Vui lòng chọn Nhân viên và nhập Địa điểm công tác!" if is_vi else "⚠️ 担当者を選択し、現地場所を入力してください！")
                elif ngay_bat_dau > ngay_ket_thuc:
                    st.error("⚠️ Ngày kết thúc không được nhỏ hơn ngày bắt đầu!" if is_vi else "⚠️ 終了日は開始日より前の日付にできません！")
                else:
                    curr_d = ngay_bat_dau
                    saved_count = 0
                    try:
                        while curr_d <= ngay_ket_thuc:
                            time_str = curr_d.strftime("%d/%m/%Y") + " " + gio_checkin.strftime("%H:%M:%S")
                            save_field_checkin(ma_nv.strip(), ten_nv.strip(), time_str, loai, dia_diem.strip(), "", ghi_chu.strip())
                            curr_d += datetime.timedelta(days=1)
                            saved_count += 1
                            
                        if saved_count == 1:
                            st.success(f"✅ Đã ghi nhận thành công ngày {ngay_bat_dau.strftime('%d/%m/%Y')}!" if is_vi else f"✅ {ngay_bat_dau.strftime('%Y/%m/%d')} の打刻を記録しました！")
                        else:
                            st.success(f"✅ Đã ghi nhận công tác thành công cho {saved_count} ngày (từ {ngay_bat_dau.strftime('%d/%m/%Y')} đến {ngay_ket_thuc.strftime('%d/%m/%Y')})!" if is_vi else f"✅ {saved_count}日間 ({ngay_bat_dau.strftime('%Y/%m/%d')} ~ {ngay_ket_thuc.strftime('%Y/%m/%d')}) の作業を記録しました！")
                    except Exception as e:
                        from log_config import logger
                        logger.error(f"Lỗi khi lưu khai báo: {e}", exc_info=True)
                        st.error(f"❌ Có lỗi xảy ra khi lưu vào CSDL: {e}")
                    st.rerun()

    with col_hist:
        with st.container(key="history_table_container"):
            df_history = get_field_checkins()
            is_vi = (st.session_state.lang == 'vi')
            
            col_h1, col_h2 = st.columns([3, 1.2])
            with col_h1:
                st.markdown("### 📋 Lịch sử check-in" if is_vi else "### 📋 最新の現地打刻履歴")
            with col_h2:
                if not df_history.empty:
                    if st.button("🗑️ Xóa toàn bộ" if is_vi else "🗑️ 全件削除", key="btn_del_all_gps", use_container_width=True):
                        st.session_state.confirm_del_gps_all = True
                    if st.session_state.get('confirm_del_gps_all', False):
                        st.warning("⚠️ Bạn có chắc chắn muốn xóa toàn bộ lịch sử Check-in hiện trường không?" if is_vi else "⚠️ 現地打刻履歴をすべて削除してもよろしいですか？")
                        col_y, col_n = st.columns(2)
                        if col_y.button("✔️ Có, xóa hết" if is_vi else "✔️ はい、すべて削除"):
                            conn = sqlite3.connect(DB_FILE)
                            conn.execute("DELETE FROM field_checkins")
                            conn.commit(); conn.close()
                            st.session_state.confirm_del_gps_all = False
                            st.rerun()
                        if col_n.button("❌ Hủy" if is_vi else "❌ キャンセル"):
                            st.session_state.confirm_del_gps_all = False
                            st.rerun()

            if df_history.empty:
                st.caption("Chưa có dữ liệu check-in hiện trường nào được ghi nhận." if is_vi else "現地打刻の記録がまだありません。")
            else:
                raw_df = df_history.copy()
                def translate_loai_str(val):
                    s = str(val)
                    if not is_vi:
                        if 'Vào ca' in s or 'Check-in' in s or '出勤' in s: return "🟢 出勤"
                        if 'Tan ca' in s or 'Check-out' in s or '退勤' in s: return "🟣 退勤"
                    else:
                        if 'Vào ca' in s or 'Check-in' in s or '出勤' in s: return "🟢 Vào ca (Check-in)"
                        if 'Tan ca' in s or 'Check-out' in s or '退勤' in s: return "🟣 Tan ca (Check-out)"
                    return val

                cur_lang = st.session_state.get('lang', 'vi')
                emp_dict_synced = get_company_emp_dict(cur_lang)
                if 'loai' in df_history.columns:
                    df_history['loai'] = df_history['loai'].apply(translate_loai_str)
                if 'ten_nv' in df_history.columns:
                    synced_names = []
                    for _, r_hist in df_history.iterrows():
                        m_h = str(r_hist.get('ma_nv', '')).strip()
                        if m_h in emp_dict_synced:
                            synced_names.append(emp_dict_synced[m_h])
                        else:
                            synced_names.append(translate_name(str(r_hist.get('ten_nv', '')), cur_lang))
                    df_history['ten_nv'] = synced_names
                if 'dia_diem' in df_history.columns:
                    df_history['dia_diem'] = [translate_dia_diem(str(x), cur_lang) for x in df_history['dia_diem']]
                if not is_vi and 'thoi_gian' in df_history.columns:
                    def fmt_jp_dt(val):
                        s = str(val).strip()
                        if len(s) >= 10 and s[2] == '/' and s[5] == '/':
                            return f"{s[6:10]}/{s[3:5]}/{s[0:2]}" + s[10:]
                        return s
                    df_history['thoi_gian'] = df_history['thoi_gian'].apply(fmt_jp_dt)
                for c in df_history.columns:
                    df_history[c] = ["" if str(v).strip().lower() in ['nan', '<na>', 'none', 'null'] else v for v in df_history[c]]
                if not is_vi:
                    df_history = df_history.rename(columns={
                        'ma_nv': '社員ID', 'ten_nv': '氏名', 'thoi_gian': '打刻日時',
                        'loai': '打刻種別', 'dia_diem': '現地場所', 'ghi_chu': '作業詳細'
                    })
                else:
                    df_history = df_history.rename(columns={
                        'ma_nv': 'Mã NV', 'ten_nv': 'Tên NV', 'thoi_gian': 'Thời gian',
                        'loai': 'Loại', 'dia_diem': 'Địa điểm', 'ghi_chu': 'Chi tiết'
                    })
                for drop_col in ['id', 'toa_do', '座標', 'Tọa độ']:
                    if drop_col in df_history.columns:
                        df_history = df_history.drop(columns=[drop_col])
                st.dataframe(df_history, use_container_width=True, hide_index=True)

                expander_title = "🗑️ Xóa từng dòng dữ liệu sai" if is_vi else "🗑️ 誤った打刻レコードの削除"
                with st.expander(expander_title):
                    del_opts = [f"ID {r['id']}: {r['thoi_gian']} | {r['ma_nv']} - {translate_name(r['ten_nv'], st.session_state.lang)} | {translate_loai_str(r['loai'])}" for _, r in raw_df.iterrows()]
                    placeholder_sel = "-- Chọn lượt check-in cần xóa --" if is_vi else "-- 削除する打刻を選択 --"
                    sel_label = "Chọn lượt check-in cần xóa:" if is_vi else "削除する打刻を選択:"
                    sel_del = st.selectbox(sel_label, [placeholder_sel] + del_opts, key="gps_single_del")
                    if sel_del and sel_del != placeholder_sel and str(sel_del).startswith("ID "):
                        try:
                            del_id = int(str(sel_del).split(":")[0].replace("ID", "").strip())
                        except ValueError:
                            del_id = None
                        if del_id is not None and st.button("🗑️ Xác nhận xóa dòng này" if is_vi else "🗑️ このレコードを削除", type="primary", key="btn_confirm_gps_del"):
                            conn = sqlite3.connect(DB_FILE)
                            conn.execute("DELETE FROM field_checkins WHERE id = ?", (del_id,))
                            conn.commit(); conn.close()
                            st.success("✅ Xóa thành công!" if is_vi else "✅ 削除しました！")
                            st.rerun()



if __name__ == '__main__':
    render_checkin_page()
