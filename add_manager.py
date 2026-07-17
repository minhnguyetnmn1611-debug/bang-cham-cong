import re

with open('app.py', 'r', encoding='utf-8') as f:
    code = f.read()

# 1. Insert the Quản lý nhân viên block into render_advanced_settings_sidebar
manager_code = '''
    if "deleted_emps" not in st.session_state:
        st.session_state.deleted_emps = set()
    if "edited_emps" not in st.session_state:
        st.session_state.edited_emps = {}

    with st.sidebar.expander("✏️ Quản lý nhân viên (Sửa/Xóa)"):
        st.markdown("<small>Sửa thông tin hoặc xóa nhân viên khỏi danh sách.</small>", unsafe_allow_html=True)
        if st.session_state.get('df_raw') is not None and 'mapping' in st.session_state:
            m_temp = st.session_state.mapping
            ma_nv_col = m_temp.get('ma_nv')
            ten_nv_col = m_temp.get('ten_nv')
            
            if ma_nv_col and ten_nv_col and ma_nv_col in st.session_state.df_raw.columns:
                raw_emps = st.session_state.df_raw[[ma_nv_col, ten_nv_col]].dropna().drop_duplicates()
                emp_dict = {str(row[ma_nv_col]).strip(): str(row[ten_nv_col]).strip() for _, row in raw_emps.iterrows()}
                
                # Add manual emps
                for me in st.session_state.manual_emps:
                    emp_dict[me['ma']] = me['ten']
                
                # Remove deleted
                for d_ma in list(st.session_state.deleted_emps):
                    emp_dict.pop(d_ma, None)
                    
                if emp_dict:
                    emp_list_display = [f"{k} - {v}" for k, v in emp_dict.items()]
                    sel_emp_str = st.selectbox("Chọn nhân viên", [""] + sorted(emp_list_display), key="sel_emp_manager")
                    if sel_emp_str:
                        sel_ma = sel_emp_str.split(" - ")[0]
                        st.markdown("---")
                        curr_pb = st.session_state.edited_emps.get(sel_ma, {}).get('pb', "")
                        curr_cv = st.session_state.edited_emps.get(sel_ma, {}).get('cv', "")
                        new_pb = st.text_input("Phòng ban mới", value=curr_pb, key="edit_pb")
                        new_cv = st.text_input("Chức vụ mới", value=curr_cv, key="edit_cv")
                        
                        col_e1, col_e2 = st.columns(2)
                        with col_e1:
                            if st.button("💾 Lưu", use_container_width=True, key="btn_save_emp"):
                                if sel_ma not in st.session_state.edited_emps:
                                    st.session_state.edited_emps[sel_ma] = {}
                                st.session_state.edited_emps[sel_ma]['pb'] = new_pb
                                st.session_state.edited_emps[sel_ma]['cv'] = new_cv
                                st.rerun()
                        with col_e2:
                            if st.button("🗑️ Xóa NV", type="primary", use_container_width=True, key="btn_del_emp"):
                                st.session_state.deleted_emps.add(sel_ma)
                                st.rerun()
                else:
                    st.info("Không có nhân viên nào.")
            else:
                st.info("Hãy hoàn thành Bước 2 (ghép cột).")
        else:
            st.info("Hãy tải file lên để xem danh sách NV.")
'''

old_manual_ot_check = '''    if "manual_ot" not in st.session_state:
        st.session_state.manual_ot = {}'''

new_manual_ot_check = manager_code + "\n" + old_manual_ot_check
code = code.replace(old_manual_ot_check, new_manual_ot_check)


# 2. Insert the filter logic into step 3 right after df_filtered is built (around "Không có dữ liệu trong khoảng thời gian đã chọn.")
filter_logic = '''
        # APPLY DELETED EMPS AND EDITS
        if st.session_state.get('deleted_emps'):
            df_filtered = df_filtered[~df_filtered[m['ma_nv']].astype(str).str.strip().isin(st.session_state.deleted_emps)]
            
        if st.session_state.get('edited_emps'):
            def apply_edits(row):
                ma = str(row[m['ma_nv']]).strip()
                if ma in st.session_state.edited_emps:
                    if 'phong_ban' in m and st.session_state.edited_emps[ma].get('pb'):
                        row[m['phong_ban']] = st.session_state.edited_emps[ma]['pb']
                    if 'chuc_vu' in m and st.session_state.edited_emps[ma].get('cv'):
                        row[m['chuc_vu']] = st.session_state.edited_emps[ma]['cv']
                return row
            df_filtered = df_filtered.apply(apply_edits, axis=1)

'''

target_if_empty = '''        if df_filtered.empty:'''
new_if_empty = filter_logic + target_if_empty
code = code.replace(target_if_empty, new_if_empty)


with open('app.py', 'w', encoding='utf-8') as f:
    f.write(code)

with open(r'C:\Users\kifukouza05\Desktop\app.py', 'w', encoding='utf-8') as f:
    f.write(code)

print("Added manager successfully")
