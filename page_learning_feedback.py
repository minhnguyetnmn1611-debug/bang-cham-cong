import streamlit as st
import datetime
import sqlite3
import os
import pandas as pd
from theme import get_theme
from db import get_company_emp_options, DB_FILE

def _translate_db_text(text, is_vi):
    if is_vi or not text or not isinstance(text, str):
        return text
    mapping_ja = {
        "catalogue": "カタログ (会社案内)",
        "tài liệu giới thiệu ngắn gọn về công ty": "会社概要・ショートパンフレット",
        "tài liệu giới thiệu": "会社紹介資料",
        "khóa học": "研修コース",
        "bài kiểm tra": "確認テスト",
        "sát hạch": "評価試験",
        "an toàn lao động": "安全衛生教育",
        "quy định công ty": "社内規定",
        "hướng dẫn sử dụng": "取扱説明書",
        "bài thi định kỳ": "定期試験",
        "đề thi": "試験問題",
        "hạn chót": "締切",
        "tải đề": "問題DL",
        "tài liệu": "資料",
        "bài tập": "課題"
    }
    lower_text = text.strip().lower()
    if lower_text in mapping_ja:
        return mapping_ja[lower_text]
    for k, v in mapping_ja.items():
        if k in lower_text and len(k) > 4:
            return v
    return text

def render_learning_feedback_page():
    is_vi = (st.session_state.get('lang', 'vi') == 'vi')
    T = get_theme(st.session_state.get('theme_mode', 'light'))
    
    st.markdown(f"<h2 style='color:#0F172A; font-size:28px; font-weight:800; margin-bottom: 8px;'>{'🎓 Đào tạo & Đánh giá' if is_vi else '🎓 研修 & 評価'}</h2>", unsafe_allow_html=True)
    st.markdown(f"<p style='color:#64748B; font-size: 15px; margin-bottom: 30px;'>{'Không gian học tập nội bộ và hệ thống phản hồi đa chiều.' if is_vi else '社内学習スペースと多面的なフィードバックシステム。'}</p>", unsafe_allow_html=True)

    tab1, tab2 = st.tabs(["Trung tâm Đào tạo" if is_vi else "研修センター", "Đánh giá" if is_vi else "評価"])

    with tab1:
        # ---- ADMIN AREA FOR UPLOADING COURSES ----
        with st.expander("🔐 Khu vực dành cho Quản lý (Admin Only)" if is_vi else "🔐 管理者専用エリア"):
            st.info("Nhập mật khẩu để truy cập chức năng tải lên tài liệu học tập." if is_vi else "学習資料のアップロード機能にアクセスするにはパスワードを入力してください。")
            admin_pwd = st.text_input("Mật khẩu" if is_vi else "パスワード", type="password")
            
            if admin_pwd == "admin123":
                st.success("Xác thực thành công. Bạn có thể tải lên tài liệu." if is_vi else "認証成功。資料をアップロードできます。")
                tab_admin1, tab_admin2, tab_admin3 = st.tabs(["Tài liệu & Khóa học" if is_vi else "資料＆コース", "Tạo Bài Thi (Trắc nghiệm)" if is_vi else "テスト作成", "Kết quả thi" if is_vi else "テスト結果"])
                
                with tab_admin1:
                    with st.form("upload_course_form"):
                        c_title = st.text_input("Tên khóa học / Tài liệu *" if is_vi else "コース・資料名 *")
                        c_desc = st.text_area("Mô tả ngắn gọn" if is_vi else "簡単な説明")
                        c_icon = st.selectbox("Biểu tượng (Icon)" if is_vi else "アイコン", ["🏢", "🛡️", "🗣️", "📖", "💻", "📊", "🔥"])
                        c_file = st.file_uploader("Chọn file tài liệu (PDF, MP4, DOCX)..." if is_vi else "資料ファイルを選択...", type=["pdf", "mp4", "docx", "pptx", "xlsx", "zip"])
                        
                        submitted = st.form_submit_button("Tải lên & Lưu" if is_vi else "アップロードして保存", type="primary")
                        if submitted:
                            if c_title and c_file:
                                try:
                                    # Save file to disk
                                    save_dir = os.path.join("uploads", "courses")
                                    os.makedirs(save_dir, exist_ok=True)
                                    
                                    # Generate unique filename to avoid overriding
                                    ts = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
                                    safe_filename = f"{ts}_{c_file.name}"
                                    file_path = os.path.join(save_dir, safe_filename)
                                    
                                    with open(file_path, "wb") as f:
                                        f.write(c_file.getbuffer())
                                    
                                    # Save to DB
                                    conn = sqlite3.connect(DB_FILE)
                                    c = conn.cursor()
                                    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                    c.execute("INSERT INTO courses (title, description, icon, file_name, file_path, timestamp) VALUES (?, ?, ?, ?, ?, ?)",
                                              (c_title.strip(), c_desc.strip(), c_icon, c_file.name, file_path, now))
                                    conn.commit()
                                    conn.close()
                                    
                                    st.success("Đã tải lên và tạo khóa học thành công!" if is_vi else "正常にアップロードおよびコース作成が完了しました！")
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"Lỗi hệ thống: {e}")
                            else:
                                st.warning("Vui lòng nhập Tên khóa học và Chọn file!" if is_vi else "コース名を入力し、ファイルを選択してください！")
                
                with tab_admin2:
                    st.markdown("### Giao Bài thi mới (Upload File Đề)" if is_vi else "### 新しいテストを配信")
                    with st.form("upload_exam_form"):
                        exam_title = st.text_input("Tên bài thi / sát hạch *" if is_vi else "テスト名 *")
                        exam_deadline = st.date_input("Hạn chót *" if is_vi else "期限 *")
                        exam_file = st.file_uploader("Chọn file đề thi (Word, Excel, PDF, ZIP)..." if is_vi else "テストファイルを選択...", type=["docx", "doc", "xlsx", "xls", "pdf", "zip"])
                        
                        submitted_ex = st.form_submit_button("Phát hành Bài Thi" if is_vi else "テストを公開", type="primary")
                        if submitted_ex:
                            if exam_title.strip() and exam_file:
                                try:
                                    save_dir = os.path.join("uploads", "exams")
                                    os.makedirs(save_dir, exist_ok=True)
                                    ts = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
                                    safe_filename = f"{ts}_{exam_file.name}"
                                    file_path = os.path.join(save_dir, safe_filename)
                                    with open(file_path, "wb") as f:
                                        f.write(exam_file.getbuffer())
                                    
                                    conn = sqlite3.connect(DB_FILE)
                                    c = conn.cursor()
                                    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                    c.execute("INSERT INTO exams (title, deadline, file_name, file_path, timestamp) VALUES (?, ?, ?, ?, ?)",
                                              (exam_title.strip(), str(exam_deadline), exam_file.name, file_path, now))
                                    conn.commit()
                                    conn.close()
                                    st.success("Giao bài thi thành công!" if is_vi else "テストが公開されました！")
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"Lỗi: {e}")
                            else:
                                st.warning("Vui lòng nhập Tên bài thi và chọn File đề thi!" if is_vi else "テスト名を入力し、ファイルを選択してください！")
                
                with tab_admin3:
                    st.markdown("### Danh sách bài nộp của nhân viên" if is_vi else "### 従業員の提出リスト")
                    conn = sqlite3.connect(DB_FILE)
                    c = conn.cursor()
                    c.execute('''
                        SELECT s.id, e.title, s.author, s.file_name, s.file_path, s.timestamp
                        FROM exam_submissions s
                        JOIN exams e ON s.exam_id = e.id
                        ORDER BY s.timestamp DESC
                    ''')
                    subs = c.fetchall()
                    conn.close()
                    
                    if not subs:
                        st.info("Chưa có nhân viên nào nộp bài." if is_vi else "まだ提出者がいません。")
                    else:
                        for sub in subs:
                            sub_id, ex_title, author, f_name, f_path, ts = sub
                            ex_title_disp = _translate_db_text(ex_title, is_vi)
                            lbl_emp = "Nhân viên" if is_vi else "従業員"
                            lbl_ex = "Bài thi" if is_vi else "テスト"
                            lbl_time = "Thời gian nộp" if is_vi else "提出日時"
                            with st.container(border=True):
                                col1, col2 = st.columns([3, 1])
                                with col1:
                                    st.markdown(f"**👤 {lbl_emp}:** {author} — **📋 {lbl_ex}:** {ex_title_disp}")
                                    st.caption(f"🕒 {lbl_time}: {ts} — 📁 File: `{f_name}`")
                                with col2:
                                    try:
                                        with open(f_path, "rb") as f:
                                            f_data = f.read()
                                        st.download_button(
                                            label="📥 Tải bài làm" if is_vi else "📥 ダウンロード",
                                            data=f_data,
                                            file_name=f"{author}_{f_name}",
                                            key=f"dl_sub_{sub_id}",
                                            use_container_width=True
                                        )
                                    except Exception:
                                        st.error("Lỗi file" if is_vi else "ファイルエラー")
            elif admin_pwd != "":
                st.error("Mật khẩu không chính xác!" if is_vi else "パスワードが間違っています！")

        # ---- DISPLAY COURSES FROM DB ----
        st.markdown(f"<h3 style='color:#0F172A; font-size:18px; font-weight:700; margin-bottom: 15px; margin-top: 20px;'>{'📚 Khóa học của tôi' if is_vi else '📚 マイコース'}</h3>", unsafe_allow_html=True)
        
        try:
            conn = sqlite3.connect(DB_FILE)
            c = conn.cursor()
            c.execute("SELECT id, title, description, icon, file_name, file_path FROM courses ORDER BY timestamp DESC")
            courses = c.fetchall()
            conn.close()
        except Exception as e:
            st.error(f"Lỗi truy xuất khóa học: {e}")
            courses = []
            
        if not courses:
            st.info("Chưa có tài liệu đào tạo nào. Quản lý cần tải lên trước." if is_vi else "学習資料がありません。管理者がアップロードする必要があります。")
        else:
            # We display courses in a grid (3 columns)
            cols = st.columns(3)
            for idx, course in enumerate(courses):
                c_id, title, desc, icon, file_name, file_path = course
                col = cols[idx % 3]
                title_disp = _translate_db_text(title, is_vi)
                desc_disp = _translate_db_text(desc, is_vi)
                
                with col:
                    # Dynamic color based on ID to make it look colorful
                    colors = [
                        ("linear-gradient(135deg, #3B82F6, #1D4ED8)", "#3B82F6"), # Blue
                        ("linear-gradient(135deg, #10B981, #047857)", "#10B981"), # Green
                        ("linear-gradient(135deg, #8B5CF6, #6D28D9)", "#8B5CF6"), # Purple
                        ("linear-gradient(135deg, #F59E0B, #B45309)", "#F59E0B"), # Orange
                        ("linear-gradient(135deg, #EC4899, #BE185D)", "#EC4899")  # Pink
                    ]
                    bg_grad, prog_color = colors[c_id % len(colors)]
                    
                    st.markdown(f"""
                    <div style="background:{T['bg_card']}; border-radius: 12px; border: 1px solid {T['border']}; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.05); margin-bottom: 20px;">
                        <div style="height: 120px; background: {bg_grad}; display: flex; align-items: center; justify-content: center; font-size: 40px;">{icon}</div>
                        <div style="padding: 15px;">
                            <h4 style="color:{T['text_primary']}; margin-top:0; margin-bottom:8px; font-size:16px;">{title_disp}</h4>
                            <div style="color:{T['text_tertiary']}; font-size:13px; margin-bottom:12px; height: 35px; overflow: hidden; text-overflow: ellipsis;">{desc_disp}</div>
                            <div style="background:{T['border']}; border-radius:10px; height:6px; width:100%; margin-bottom: 12px;">
                                <div style="background:{prog_color}; border-radius:10px; height:100%; width: 0%;"></div>
                            </div>
                            <div style="font-size:12px; color:{T['text_secondary']}; margin-bottom: 15px;">🗂️ {file_name}</div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    # Read file to create download button
                    try:
                        with open(file_path, "rb") as f:
                            file_data = f.read()
                        st.download_button(
                            label="📥 Tải xuống" if is_vi else "📥 ダウンロード",
                            data=file_data,
                            file_name=file_name,
                            key=f"dl_course_{c_id}",
                            use_container_width=True
                        )
                    except Exception:
                        st.error("Lỗi file" if is_vi else "ファイルエラー")
                        
        st.divider()
        st.markdown(f"<h3 style='color:#0F172A; font-size:18px; font-weight:700; margin-bottom: 15px;'>{'📋 Bài thi / Sát hạch định kỳ' if is_vi else '📋 定期テスト'}</h3>", unsafe_allow_html=True)
        
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute("SELECT id, title, deadline, file_name, file_path FROM exams WHERE status='active' ORDER BY id DESC")
        exams = c.fetchall()
        
        if not exams:
            st.info("Hiện tại chưa có bài sát hạch nào đến hạn." if is_vi else "現在、期限が迫っているテストはありません。")
        else:
            current_user = st.session_state.get('ten_nv', 'Anonymous')
            for ex in exams:
                ex_id, ex_title, ex_dl, ex_fname, ex_fpath = ex
                ex_title_disp = _translate_db_text(ex_title, is_vi)
                dl_lbl = "Hạn chót" if is_vi else "締切"
                with st.expander(f"📌 {ex_title_disp} ({dl_lbl}: {ex_dl})", expanded=False):
                    col_ex1, col_ex2 = st.columns([1, 1])
                    with col_ex1:
                        st.markdown("**1. Tải về đề thi:**" if is_vi else "**1. 問題ファイルをダウンロード:**")
                        try:
                            with open(ex_fpath, "rb") as f:
                                ex_data = f.read()
                            dl_text = f"📥 Tải đề: {ex_fname}" if is_vi else f"📥 問題DL: {ex_fname}"
                            st.download_button(
                                label=dl_text,
                                data=ex_data,
                                file_name=ex_fname,
                                key=f"dl_ex_{ex_id}",
                                use_container_width=True
                            )
                        except Exception:
                            st.error("Không tìm thấy file đề thi!" if is_vi else "問題ファイルが見つかりません！")
                            
                    with col_ex2:
                        st.markdown("**2. Nộp bài làm của bạn:**" if is_vi else "**2. 解答ファイルを提出:**")
                        c.execute("SELECT file_name, timestamp FROM exam_submissions WHERE exam_id=? AND author=? ORDER BY timestamp DESC LIMIT 1", (ex_id, current_user))
                        res_sub = c.fetchone()
                        if res_sub:
                            st.success(f"✅ Đã nộp: `{res_sub[0]}` ({res_sub[1]})" if is_vi else f"✅ 提出済み: `{res_sub[0]}` ({res_sub[1]})")
                            st.caption("Bạn có thể tải lên file khác bên dưới để nộp đè nếu muốn thay đổi." if is_vi else "変更したい場合は、下に新しいファイルをアップロードして上書きできます。")
                        
                        sub_file = st.file_uploader("Chọn file bài làm..." if is_vi else "解答ファイルを選択...", key=f"up_sub_{ex_id}", label_visibility="collapsed")
                        if st.button("📤 Gửi bài làm" if is_vi else "📤 提出する", key=f"btn_sub_{ex_id}", type="primary", use_container_width=True):
                            if sub_file:
                                try:
                                    save_dir = os.path.join("uploads", "submissions")
                                    os.makedirs(save_dir, exist_ok=True)
                                    ts = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
                                    safe_filename = f"{ts}_{current_user}_{sub_file.name}"
                                    file_path = os.path.join(save_dir, safe_filename)
                                    with open(file_path, "wb") as f:
                                        f.write(sub_file.getbuffer())
                                    
                                    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                    c.execute("INSERT INTO exam_submissions (exam_id, author, file_name, file_path, timestamp) VALUES (?, ?, ?, ?, ?)",
                                              (ex_id, current_user, sub_file.name, file_path, now))
                                    conn.commit()
                                    st.success("Nộp bài thành công!" if is_vi else "提出が完了しました！")
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"Lỗi: {e}")
                            else:
                                st.warning("Vui lòng chọn file bài làm trước khi gửi!" if is_vi else "送信する前にファイルを選択してください！")
        conn.close()

    with tab2:
        col_fb1, col_fb2 = st.columns([1, 1])
        with col_fb1:
            st.markdown(f"<h3 style='color:#0F172A; font-size:18px; font-weight:700; margin-bottom: 15px;'>{'✍️ Gửi Phiếu đánh giá' if is_vi else '✍️ 評価を送信'}</h3>", unsafe_allow_html=True)
            emps = get_company_emp_options('vi' if is_vi else 'ja')
            target_emp = st.selectbox("Chọn người bạn muốn đánh giá" if is_vi else "評価する人を選択", emps)
            st.slider("Kỹ năng chuyên môn (1-10)" if is_vi else "専門スキル (1-10)", 1, 10, 8)
            st.slider("Thái độ làm việc & Tương tác (1-10)" if is_vi else "勤務態度とコミュニケーション (1-10)", 1, 10, 8)
            st.text_area("Nhận xét điểm mạnh" if is_vi else "長所", placeholder="Điểm cộng của bạn ấy là..." if is_vi else "素晴らしい点は...")
            st.text_area("Điểm cần cải thiện" if is_vi else "改善点", placeholder="Để tốt hơn, bạn nên..." if is_vi else "改善すべき点は...")
            
            # Anonymous toggle
            st.checkbox("Gửi đánh giá ẩn danh (Anonymous)" if is_vi else "匿名で送信する", value=True)
            
            if st.button("Gửi đánh giá" if is_vi else "評価を送信", type="primary"):
                st.success("Đã gửi phiếu đánh giá thành công! Dữ liệu đã được lưu trữ bảo mật." if is_vi else "評価を送信しました！データは安全に保存されました。")
                
        with col_fb2:
            st.markdown(f"<h3 style='color:#0F172A; font-size:18px; font-weight:700; margin-bottom: 15px;'>{'📬 Đánh giá tôi nhận được' if is_vi else '📬 受け取った評価'}</h3>", unsafe_allow_html=True)
            
            from_anon = "Từ: Một đồng nghiệp giấu tên" if is_vi else "差出人: 匿名の同僚"
            time_anon = "Tháng trước" if is_vi else "先月"
            msg_anon = '"Bạn làm việc rất chăm chỉ và luôn hỗ trợ mọi người nhiệt tình. Tuy nhiên đôi khi bạn ôm đồm quá nhiều việc dẫn đến trễ deadline, hãy học cách ủy quyền nhé!"' if is_vi else '"いつも熱心に働いていただき、親身にサポートしてくれて感謝しています。ただ、一人で仕事を抱え込みすぎて締め切りに遅れることがあるので、任せることも覚えてください！"'
            
            from_mgr = "Từ: Quản lý trực tiếp" if is_vi else "差出人: 直属の上司"
            time_mgr = "Kỳ đánh giá Q2/2026" if is_vi else "2026年第2四半期評価"
            msg_mgr = '"Kỹ năng chuyên môn tuyệt vời, đã hoàn thành vượt chỉ tiêu dự án tháng trước. Rất đáng biểu dương!"' if is_vi else '"専門スキルが素晴らしく、先月のプロジェクト目標を大幅に達成しました。素晴らしい貢献です！"'

            st.markdown(f"""
            <div style="background:{T['bg_card']}; padding: 15px; border-radius: 12px; border: 1px solid {T['border']}; margin-bottom: 15px; border-left: 4px solid #3B82F6;">
                <div style="font-weight: 600; color: #0F172A; margin-bottom: 5px;">{from_anon}</div>
                <div style="color: #64748B; font-size: 13px; margin-bottom: 10px;">{time_anon}</div>
                <div style="color: #334155; font-size: 14px; font-style: italic;">{msg_anon}</div>
            </div>
            <div style="background:{T['bg_card']}; padding: 15px; border-radius: 12px; border: 1px solid {T['border']}; border-left: 4px solid #10B981;">
                <div style="font-weight: 600; color: #0F172A; margin-bottom: 5px;">{from_mgr}</div>
                <div style="color: #64748B; font-size: 13px; margin-bottom: 10px;">{time_mgr}</div>
                <div style="color: #334155; font-size: 14px; font-style: italic;">{msg_mgr}</div>
            </div>
            """, unsafe_allow_html=True)
