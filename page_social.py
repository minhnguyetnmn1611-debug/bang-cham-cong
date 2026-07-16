import streamlit as st
import datetime
import sqlite3
from theme import get_theme
from db import get_company_emp_options, DB_FILE
@st.fragment
def render_post_item(post, is_vi, is_admin, T, pre_fetched_comments=None):
    p_id, author, content, timestamp, initial_likes, real_author, is_anon = post
    
    # We will only query fresh data if the fragment is rerunning due to an interaction.
    # Otherwise, we use the pre-fetched data to save N+1 queries.
    if 'social_interacted_' + str(p_id) in st.session_state:
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute("SELECT likes FROM newsfeed WHERE id=?", (p_id,))
        row = c.fetchone()
        likes = row[0] if row else initial_likes
        
        try:
            c.execute("SELECT COUNT(*) FROM comments WHERE post_id=?", (p_id,))
            cmt_count = c.fetchone()[0]
            c.execute("SELECT author, content, timestamp FROM comments WHERE post_id=? ORDER BY timestamp ASC", (p_id,))
            cmts = c.fetchall()
        except:
            cmt_count = 0
            cmts = []
        conn.close()
    else:
        likes = initial_likes
        cmts = pre_fetched_comments if pre_fetched_comments is not None else []
        cmt_count = len(cmts)

    # Determine display name
    if is_anon:
        if is_admin:
            display_name = f"Người Dùng Ẩn Danh ({real_author or author})" if is_vi else f"匿名ユーザー ({real_author or author})"
        else:
            display_name = "Người Dùng Ẩn Danh" if is_vi else "匿名ユーザー"
    else:
        display_name = real_author if real_author else author
    
    try:
        dt = datetime.datetime.strptime(timestamp, "%Y-%m-%d %H:%M:%S")
        time_str = dt.strftime("%d/%m/%Y %H:%M")
    except:
        time_str = timestamp
        
    color_hex = "#" + hex(hash(display_name) & 0xFFFFFF)[2:].zfill(6)
    
    st.markdown(f"""
    <div style="background:{T['bg_card']}; padding: 20px 20px 10px 20px; border-radius: 12px 12px 0 0; border: 1px solid {T['border']}; border-bottom: none; margin-top: 20px;">
        <div style="display: flex; align-items: center; margin-bottom: 15px;">
            <div style="width: 40px; height: 40px; border-radius: 50%; background: {color_hex}; color: white; display: flex; align-items: center; justify-content: center; font-weight: bold; font-size: 18px; margin-right: 15px;">{display_name[0:1].upper() if display_name else '?'}</div>
            <div>
                <div style="font-weight: 700; color: {T['text_primary']}; font-size: 16px;">{display_name}</div>
                <div style="color: {T['text_tertiary']}; font-size: 12px;">{time_str}</div>
            </div>
        </div>
        <div style="color: {T['text_secondary']}; font-size: 15px; line-height: 1.6; margin-bottom: 10px;">
            {content}
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Interactive buttons row
    st.markdown(f"<div style='background:{T['bg_card']}; padding: 0 20px 10px 20px; border-radius: 0 0 12px 12px; border: 1px solid {T['border']}; border-top: none; box-shadow: 0 4px 6px rgba(0,0,0,0.05); margin-bottom: 20px;'>", unsafe_allow_html=True)
    
    col1, col2, _ = st.columns([1.5, 1.5, 5])
    with col1:
        if st.button(f"❤️ {likes} Thích", key=f"like_{p_id}", help="Thích bài viết này" if is_vi else "いいね"):
            try:
                conn = sqlite3.connect(DB_FILE)
                c = conn.cursor()
                c.execute("UPDATE newsfeed SET likes = likes + 1 WHERE id=?", (p_id,))
                conn.commit()
                conn.close()
                st.session_state['social_interacted_' + str(p_id)] = True
                st.rerun(scope="fragment")
            except Exception as e:
                st.error(f"Lỗi: {e}")
    with col2:
        if is_admin:
            if st.button("🗑️ Xóa", key=f"del_{p_id}", help="Xóa bài viết (Quyền Quản lý)" if is_vi else "削除"):
                try:
                    conn = sqlite3.connect(DB_FILE)
                    c = conn.cursor()
                    c.execute("DELETE FROM newsfeed WHERE id=?", (p_id,))
                    c.execute("DELETE FROM comments WHERE post_id=?", (p_id,))
                    conn.commit()
                    conn.close()
                    st.rerun() # Needs full rerun to remove from loop
                except Exception as e:
                    st.error(f"Lỗi xóa bài: {e}")
    
    with st.expander(f"💬 Bình luận ({cmt_count})" if is_vi else f"💬 コメント ({cmt_count})"):
        # Load comments
        if cmts:
            for cmt in cmts:
                c_author, c_content, c_time = cmt
                try:
                    dt_c = datetime.datetime.strptime(c_time, "%Y-%m-%d %H:%M:%S")
                    c_time_str = dt_c.strftime("%d/%m %H:%M")
                except:
                    c_time_str = c_time
                    
                st.markdown(f"<div style='margin-bottom: 5px;'><b style='color:{T['text_primary']}'>{c_author}</b> <span style='color:{T['text_tertiary']}; font-size:12px;'>({c_time_str})</span></div>", unsafe_allow_html=True)
                st.markdown(f"<div style='color:{T['text_secondary']}; margin-bottom: 10px;'>👉 {c_content}</div>", unsafe_allow_html=True)
                st.markdown("<hr style='margin: 5px 0; border: none; border-top: 1px dashed #E2E8F0;'>", unsafe_allow_html=True)
        else:
            st.info("Chưa có bình luận nào. Hãy là người đầu tiên!" if is_vi else "コメントはまだありません。")
            
        # Add comment input
        new_cmt = st.text_input("Viết bình luận..." if is_vi else "コメントを書く...", key=f"new_cmt_{p_id}")
        if st.button("Gửi bình luận" if is_vi else "送信", key=f"btn_cmt_{p_id}", type="primary"):
            if new_cmt.strip():
                try:
                    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    conn = sqlite3.connect(DB_FILE)
                    c = conn.cursor()
                    c.execute("INSERT INTO comments (post_id, author, content, timestamp) VALUES (?, ?, ?, ?)",
                              (p_id, "Người Dùng Ẩn Danh", new_cmt.strip(), now))
                    conn.commit()
                    conn.close()
                    st.session_state[f"new_cmt_{p_id}"] = ""
                    st.session_state['social_interacted_' + str(p_id)] = True
                    st.rerun(scope="fragment")
                except Exception as e:
                    st.error(f"Lỗi gửi bình luận: {e}")
    st.markdown("</div>", unsafe_allow_html=True)

def render_social_page():
    is_vi = (st.session_state.get('lang', 'vi') == 'vi')
    T = get_theme(st.session_state.get('theme_mode', 'light'))
    
    st.markdown(f"<h2 style='color:{T['text_primary']}; font-size:28px; font-weight:800; margin-bottom: 8px;'>{'🌟 Mạng xã hội & Vinh danh' if is_vi else '🌟 社内SNS & 表彰'}</h2>", unsafe_allow_html=True)
    st.markdown(f"<p style='color:{T['text_secondary']}; font-size: 15px; margin-bottom: 30px;'>{'Gắn kết tập thể, chia sẻ tin tức nội bộ và vinh danh những cá nhân xuất sắc.' if is_vi else '社内のニュースを共有し、優秀な個人を表彰する。'}</p>", unsafe_allow_html=True)

    tab1, tab2 = st.tabs(["Bảng tin" if is_vi else "ニュースフィード", "Vinh danh" if is_vi else "表彰"])

    with tab1:
        # ---- ADMIN MODE ----
        with st.expander("🔐 Chế độ Quản lý (Admin Mode)" if is_vi else "🔐 管理者モード"):
            st.info("Nhập mật khẩu để bật quyền dọn dẹp bài đăng và xem người đăng ẩn danh." if is_vi else "パスワードを入力して、投稿の削除権限と匿名投稿者の確認を有効にします。")
            admin_pwd = st.text_input("Mật khẩu Quản lý" if is_vi else "管理者パスワード", type="password", key="social_admin_pwd")
            if admin_pwd == "admin123":
                st.session_state['is_social_admin'] = True
                st.success("Đã bật Chế độ Quản lý!" if is_vi else "管理者モードを有効にしました！")
            else:
                st.session_state['is_social_admin'] = False
                if admin_pwd:
                    st.error("Mật khẩu không chính xác!" if is_vi else "パスワードが間違っています！")

        is_admin = st.session_state.get('is_social_admin', False)

        # Load posts from database
        try:
            conn = sqlite3.connect(DB_FILE)
            c = conn.cursor()
            try:
                c.execute("SELECT id, author, content, timestamp, likes, real_author, is_anonymous FROM newsfeed ORDER BY timestamp DESC")
            except sqlite3.OperationalError:
                c.execute("SELECT id, author, content, timestamp, likes, author, 0 FROM newsfeed ORDER BY timestamp DESC")
            posts = c.fetchall()
            
            # Pre-fetch all comments to avoid N+1 queries
            c.execute("SELECT post_id, author, content, timestamp FROM comments ORDER BY timestamp ASC")
            all_cmts = c.fetchall()
            cmts_dict = {}
            for cmt in all_cmts:
                post_id = cmt[0]
                if post_id not in cmts_dict:
                    cmts_dict[post_id] = []
                cmts_dict[post_id].append((cmt[1], cmt[2], cmt[3]))
        except Exception as e:
            st.error(f"Lỗi tải Newsfeed: {e}")
            posts = []
            cmts_dict = {}
            
        if not posts:
            st.info("Chưa có bài viết nào trên bảng tin." if is_vi else "投稿がまだありません。")
        else:
            for post in posts:
                post_id = post[0]
                post_cmts = cmts_dict.get(post_id, [])
                render_post_item(post, is_vi, is_admin, T, post_cmts)
                
        try:
            if 'conn' in locals():
                conn.close()
        except:
            pass
        
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(f"<h3 style='color:{T['text_primary']}; font-size:18px;'>{'✍️ Đăng bài viết mới' if is_vi else '✍️ 新しい投稿'}</h3>", unsafe_allow_html=True)
        
        emps = get_company_emp_options('vi' if is_vi else 'ja')
        if not emps:
            st.warning("Bạn cần tải lên dữ liệu chấm công để có danh sách nhân viên trước." if is_vi else "社員リストを取得するために、まず勤怠データをアップロードしてください。")
        else:
            with st.form("new_post_form"):
                post_author = st.selectbox("Chọn tên của bạn" if is_vi else "あなたの名前を選択", emps)
                is_anon_post = False
                new_post = st.text_area("Nhập nội dung..." if is_vi else "内容を入力...")
                
                submitted_post = st.form_submit_button("Đăng bài" if is_vi else "投稿する", type="primary")
                if submitted_post:
                    if new_post.strip():
                        try:
                            conn = sqlite3.connect(DB_FILE)
                            c = conn.cursor()
                            now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            
                            author_db = "Người Dùng Ẩn Danh" if is_anon_post else post_author
                            is_anon_db = 1 if is_anon_post else 0
                            
                            c.execute("INSERT INTO newsfeed (author, content, timestamp, likes, real_author, is_anonymous) VALUES (?, ?, ?, ?, ?, ?)", 
                                      (author_db, new_post.strip(), now, 0, post_author, is_anon_db))
                            conn.commit()
                            conn.close()
                            st.success("Đã đăng thành công!" if is_vi else "投稿しました！")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Lỗi lưu bài viết: {e}")
                    else:
                        st.warning("Vui lòng nhập nội dung!" if is_vi else "内容を入力してください！")

    with tab2:
        col1, col2 = st.columns([1, 1])
        with col1:
            st.markdown(f"<h3 style='color:{T['text_primary']}; font-size:18px; font-weight:700; margin-bottom: 15px;'>{'🏆 Bảng xếp hạng Tháng này' if is_vi else '🏆 今月のランキング'}</h3>", unsafe_allow_html=True)
            
            try:
                conn = sqlite3.connect(DB_FILE)
                c = conn.cursor()
                c.execute("SELECT recipient, COUNT(*) as score FROM vinh_danh GROUP BY recipient ORDER BY score DESC LIMIT 10")
                leaderboard = c.fetchall()
                conn.close()
            except Exception as e:
                leaderboard = []
            
            if not leaderboard:
                st.info("Chưa có dữ liệu vinh danh." if is_vi else "データがまだありません。")
            else:
                colors = ["#F59E0B", "#94A3B8", "#B45309"]
                for i, row in enumerate(leaderboard):
                    rank = i + 1
                    recipient, score = row
                    rank_color = colors[i] if i < 3 else T['text_secondary']
                    
                    st.markdown(f"""
                    <div style="background:{T['bg_card']}; padding: 15px; border-radius: 12px; border: 1px solid {T['border']}; margin-bottom: 10px; display: flex; align-items: center; justify-content: space-between;">
                        <div style="display: flex; align-items: center;">
                            <div style="font-size: {24 if i==0 else 20 if i==1 else 18 if i==2 else 16}px; font-weight: 800; color: {rank_color}; margin-right: 15px; width: 30px; text-align: center;">{rank}</div>
                            <div style="font-weight: 600; color: {T['text_primary']};">{recipient}</div>
                        </div>
                        <div style="font-weight: 700; color: {T['primary']};">{score} ⭐</div>
                    </div>
                    """, unsafe_allow_html=True)
            
        with col2:
            st.markdown(f"<h3 style='color:{T['text_primary']}; font-size:18px; font-weight:700; margin-bottom: 15px;'>{'🎁 Gửi tặng cho đồng nghiệp' if is_vi else '🎁 同僚に贈る'}</h3>", unsafe_allow_html=True)
            emps = get_company_emp_options('vi' if is_vi else 'ja')
            
            if emps:
                recipient = st.selectbox("Chọn người nhận" if is_vi else "受取人を選択", emps)
                kudos_type = st.selectbox("Loại huy hiệu" if is_vi else "バッジの種類", ["🌟 Đồng đội xuất sắc", "💡 Sáng tạo đột phá", "🔥 Tinh thần thép", "🤝 Người kết nối"])
                message = st.text_area("Lời nhắn" if is_vi else "メッセージ")
                if st.button("Gửi tặng" if is_vi else "贈る", type="primary", key="btn_send_kudos"):
                    if recipient and kudos_type:
                        try:
                            conn = sqlite3.connect(DB_FILE)
                            c = conn.cursor()
                            now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            c.execute("INSERT INTO vinh_danh (sender, recipient, kudos_type, message, timestamp) VALUES (?, ?, ?, ?, ?)", 
                                      ("Người Dùng Ẩn Danh", recipient, kudos_type, message, now))
                            conn.commit()
                            conn.close()
                            st.success(f"Đã gửi tặng huy hiệu {kudos_type} tới {recipient} thành công!" if is_vi else "送信成功！")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Lỗi: {e}")
            else:
                st.warning("Vui lòng tải lên file chấm công để tạo danh sách nhân viên trước." if is_vi else "社員リストを作成するために、まず勤怠ファイルをアップロードしてください。")
