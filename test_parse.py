import sys
sys.path.insert(0, r"c:\Users\kifukouza05\Desktop\Nguyet\Web test bảng chấm công")
from page_mos import parse_single_email_report

email_text = """
◇MOS業務
・開発/山崎様 P012004 MOS様 VMOS技術開発業務委託費 【イニシャル（開発）】 Unity版タンシダイ装置のスマートデバッグ仕上げ 10h
・開発/山崎様 P012004 MOS様 VMOS技術開発業務委託費 【イニシャル（開発）】 展示会のスマートデバッグの安定性確認 20h
・開発/山崎様 P012004 MOS様 VMOS技術開発業務委託費 【イニシャル（開発）】 展示会に参加する 7.5h
"""

res = parse_single_email_report(email_text, "01/01/2026")
print("Result:")
for k, v in res.get('projects', {}).items():
    print(f"{k}: {v} hours")
    print(f"Name: {res.get('project_names', {}).get(k)}")
