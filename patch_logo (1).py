import base64
with open('assets/logo_header.png', 'rb') as f:
    b64 = base64.b64encode(f.read()).decode('utf-8')

with open('app.py', 'r', encoding='utf-8') as f:
    code = f.read()

# Replace LOGO_HEADER_B64 = load_logo_base64() with the hardcoded base64
target = 'LOGO_HEADER_B64 = load_logo_base64()'
replacement = f'LOGO_HEADER_B64 = "{b64}"'

if target in code:
    code = code.replace(target, replacement)
    with open('app.py', 'w', encoding='utf-8') as f:
        f.write(code)
    
    # Also write to Desktop directly
    with open(r'C:\Users\kifukouza05\Desktop\app.py', 'w', encoding='utf-8') as f:
        f.write(code)
    print('Thanh cong')
else:
    print('Khong tim thay target')
