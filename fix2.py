import re

file_path = 'app.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Replace hardcoded blue avatar
content = content.replace(
    '''avt_html = f'<div style="width: 48px; height: 48px; border-radius: 50%; background: linear-gradient(135deg, #E0F2FE, #F0F9FF); color: {T[\'primary\']}; border: 1.5px solid #BAE6FD; display: flex; align-items: center; justify-content: center; font-weight: 800; font-size: 18px; flex-shrink: 0; box-shadow: 0 2px 6px rgba(14,165,233,0.1);">{char_avt}</div>\'''',
    '''avt_html = f'<div style="width: 48px; height: 48px; border-radius: 50%; background: {T[\'bg_card\']}; color: {T[\'primary\']}; border: 1.5px solid {T[\'border\']}; display: flex; align-items: center; justify-content: center; font-weight: 800; font-size: 18px; flex-shrink: 0; box-shadow: 0 2px 6px rgba(0,0,0,0.05);">{char_avt}</div>\''''
)

content = content.replace(
    '''background: #F1F5F9; padding: 2px 8px; border-radius: 6px; margin-left: 6px; vertical-align: middle;">[{ma}]</span>\'''',
    '''background: {T[\'bg_card_hover\']}; padding: 2px 8px; border-radius: 6px; margin-left: 6px; vertical-align: middle;">[{ma}]</span>\''''
)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
print('Fixed avatar!')
