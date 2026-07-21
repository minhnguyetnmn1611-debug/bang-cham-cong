import re
content = """
◇MOS業務
・開発/山崎様 P012004 MOS様 VMOS技術開発業務委託費 【イニシャル（開発）】 Unity版タンシダイ装置のスマートデバッグ仕上げ 10h
・開発/山崎様 P012004 MOS様 VMOS技術開発業務委託費 【イニシャル（開発）】 展示会のスマートデバッグの安定性確認 20h
・開発/山崎様 P012004 MOS様 VMOS技術開発業務委託費 【イニシャル（開発）】 展示会に参加する 7.5h
"""

content_flat = re.sub(r'[\r\n]+', ' ', content)
content_flat = re.sub(r'[\(\（][^\)\）]*(?:[KJPVM]\d{4,}|\b\d{6}\b)[^\)\）]*[\)\）]', '()', content_flat, flags=re.IGNORECASE)
content_norm = re.sub(r'(?i)([・•◇■◆▼▲]|■\s*MOS|■\s*JMOS|■\s*社内|【\s*MOS\s*】|【\s*JMOS\s*】|【\s*社内\s*】|Dự án:|Project:|\b[KJPVM]\d{4,}\b|\b\d{6}\b|\b[KJPV]\d+\b)', r'\n\1', content_flat)
lines = content_norm.split('\n')
projects = {}
current_p_code = None
task_counter = 1

for line in lines:
    line_clean = line.strip()
    if not line_clean or line_clean.startswith('>'): continue
    line_upper = line_clean.upper()
    print(f"LINE: {line_clean}")
    
    matches = list(re.finditer(r'([KJPVM]\d{4,}|\b\d{6}\b|[KJPV]\d+)|\b(MOS|JMOS|VMOS|社内|有給|休暇|管理|応援|一般|支援)\b', line_upper))
    if matches:
        new_p_code = None
        for m in matches:
            found = m.group(1) or m.group(2)
            if found: new_p_code = found
        if new_p_code: current_p_code = new_p_code
    print(f"  current_p_code: {current_p_code}")
    
    line_working = line_clean.lower()
    if current_p_code and current_p_code.upper().startswith('P012004'):
        if any(k in line_working for k in ['triển lãm', 'trien lam', '展示会']):
            current_p_code = 'SHANAI'
    print(f"  current_p_code after exh logic: {current_p_code}")

    line_working = re.sub(r'(?<!\d:)(\d{1,2})\s*(?:h|hr|giờ|tiếng)\s*(\d{1,2})\b', '', line_working)
    for m_hrs in re.finditer(r'(\d+(?:[.,]\d+)?)\s*(?:h|hr|hrs|giờ|tiếng|h/ngày|h/day)\b|\b(?:giờ|tiếng|h|hours?|thời gian|time)\s*[:=]?\s*(\d+(?:[.,]\d+)?)\b', line_working):
        val_str = m_hrs.group(1) or m_hrs.group(2)
        val_float = float(val_str.replace(',', '.'))
        if val_float > 0 and val_float <= 24:
            p_code_use = current_p_code or "MOS"
            if p_code_use in ['MOS', 'JMOS', 'VMOS', 'SHANAI']:
                p_code_use = f"{p_code_use}_NOCODE_{task_counter}"
                task_counter += 1
            projects[p_code_use] = projects.get(p_code_use, 0.0) + val_float
            print(f"  ADDED {val_float} to {p_code_use}")

print("FINAL PROJECTS:", projects)
