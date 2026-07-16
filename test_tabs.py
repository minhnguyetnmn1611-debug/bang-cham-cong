import streamlit as st

if 'lang' not in st.session_state: st.session_state.lang = 'vi'

if st.button('Switch Lang'):
    st.session_state.lang = 'ja' if st.session_state.lang == 'vi' else 'vi'
    st.rerun()

is_vi = st.session_state.lang == 'vi'
tab_names = [
    "📝 Dữ liệu Dự án" if is_vi else "📝 プロジェクトデータ",
    "📊 Thống kê" if is_vi else "📊 統計",
    "📈 Báo cáo" if is_vi else "📈 レポート"
]

st.markdown(f"""
<style>
[data-testid="stTabs"] button[role="tab"] p {{
    display: none;
}}
[data-testid="stTabs"] button[role="tab"]:nth-child(1)::after {{
    content: "{tab_names[0]}";
    font-size: 16px;
}}
[data-testid="stTabs"] button[role="tab"]:nth-child(2)::after {{
    content: "{tab_names[1]}";
    font-size: 16px;
}}
[data-testid="stTabs"] button[role="tab"]:nth-child(3)::after {{
    content: "{tab_names[2]}";
    font-size: 16px;
}}
</style>
""", unsafe_allow_html=True)

t1, t2, t3 = st.tabs(['\u200B', '\u200B\u200B', '\u200B\u200B\u200B'])
with t1: st.write('Content 1')
with t2: st.write('Content 2')
with t3: st.write('Content 3')
