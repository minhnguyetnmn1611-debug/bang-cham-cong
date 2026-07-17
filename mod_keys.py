import re

with open('app.py', 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Update load_saved_api_key
old_load = """def load_saved_api_key():
    if os.path.exists(SECRETS_FILE):
        try:
            with open(SECRETS_FILE, "r", encoding="utf-8") as f:
                secrets = toml.load(f)
                return secrets.get("GEMINI_API_KEY", "")
        except:
            return ""
    return \"\""""

new_load = """def load_saved_api_key():
    import streamlit as st
    if "GEMINI_API_KEY" in st.secrets:
        return st.secrets["GEMINI_API_KEY"]
    if os.path.exists(SECRETS_FILE):
        try:
            with open(SECRETS_FILE, "r", encoding="utf-8") as f:
                secrets = toml.load(f)
                return secrets.get("GEMINI_API_KEY", "")
        except:
            return ""
    return \"\""""

content = content.replace(old_load, new_load)

# 2. Update summarize_tasks_with_ai (lines 1083-1090)
old_summarize = """        api_key = ""
        try:
            import toml
            with open(".streamlit/secrets.toml", "r", encoding="utf-8") as sec_f:
                secrets = toml.load(sec_f)
                # Dùng Gemini API có sẵn của user thay vì Anthropic
                api_key = secrets.get("GEMINI_API_KEY", "")
        except: pass"""

new_summarize = """        api_key = load_saved_api_key()"""

content = content.replace(old_summarize, new_summarize)

# 3. Update tong hop MOS (lines 1705-1714)
old_mos = """        else:
            import toml
            has_key = False
            try:
                with open(".streamlit/secrets.toml", "r", encoding="utf-8") as sec_f:
                    has_key = bool(toml.load(sec_f).get("GEMINI_API_KEY"))
            except: pass
            
            if not has_key:
                st.warning("⚠️ Chưa cấu hình GEMINI_API_KEY trong file .streamlit/secrets.toml. Hệ thống đang dùng chế độ nối chuỗi thủ công thay vì dùng AI!")"""

new_mos = """        else:
            has_key = bool(load_saved_api_key())
            
            if not has_key:
                st.warning("⚠️ Chưa cấu hình GEMINI_API_KEY. Hệ thống đang dùng chế độ nối chuỗi thủ công thay vì dùng AI!")"""

content = content.replace(old_mos, new_mos)

with open('app.py', 'w', encoding='utf-8') as f:
    f.write(content)

print("Updated API key logic.")
