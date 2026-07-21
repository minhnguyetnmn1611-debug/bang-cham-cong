import ast
import json
import os

def get_function_lines(filename):
    with open(filename, 'r', encoding='utf-8') as f:
        source = f.read()
    
    tree = ast.parse(source)
    functions = {}
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            functions[node.name] = [node.lineno, node.end_lineno]
    return functions

def extract_funcs(lines, funcs_dict, func_names):
    extracted = []
    for name in func_names:
        if name in funcs_dict:
            start, end = funcs_dict[name]
            # lineno is 1-indexed
            extracted.extend(lines[start-1:end])
            extracted.append("\n\n")
    return "".join(extracted)

def main():
    with open('app.py', 'r', encoding='utf-8') as f:
        lines = f.readlines()
        
    funcs_dict = get_function_lines('app.py')
    
    # 1. utils.py
    utils_names = ['to_time_obj', 'time_to_float', 'format_gio_lam', 'format_gio_lam_str', 
                   'calculate_working_hours', 'clean_date', 'get_fixed_holidays_for_years', 
                   'is_workday_func', 'is_last_saturday_of_month', 'calculate_working_days', 
                   'auto_detect_columns']
    utils_code = "import pandas as pd\nimport datetime\nimport calendar\n\n"
    utils_code += extract_funcs(lines, funcs_dict, utils_names)
    with open('utils.py', 'w', encoding='utf-8') as f:
        f.write(utils_code)
        
    # 2. excel_export.py
    excel_names = ['find_header_row', 'parse_excel_file', 'export_excel_tong_hop']
    excel_code = "import pandas as pd\nimport openpyxl\nfrom openpyxl.styles import Font, Alignment, Border, Side, PatternFill\nfrom openpyxl.utils import get_column_letter\nimport streamlit as st\nfrom utils import *\nfrom translations import get_t, get_data_t\n\n"
    excel_code += extract_funcs(lines, funcs_dict, excel_names)
    with open('excel_export.py', 'w', encoding='utf-8') as f:
        f.write(excel_code)
        
    # 3. ai_chat.py
    ai_names = ['render_chatbot', 'summarize_tasks_with_ai']
    ai_code = "import streamlit as st\nimport httpx\nimport time\nimport json\nfrom db import load_saved_api_key\n\n"
    ai_code += extract_funcs(lines, funcs_dict, ai_names)
    with open('ai_chat.py', 'w', encoding='utf-8') as f:
        f.write(ai_code)

    # 4. pages/2_MOS.py
    mos_names = ['extract_ma_nv_from_filename', 'extract_ten_nv_from_filename', 'parse_mos_file', 'tong_hop_mos', 'render_mos_page']
    mos_code = "import streamlit as st\nimport pandas as pd\nimport re\nimport io\nimport base64\nfrom excel_export import *\nfrom ai_chat import summarize_tasks_with_ai\nfrom translations import get_t, translate_name\nfrom db import get_company_emp_options, get_company_emp_dict\n\n"
    mos_code += extract_funcs(lines, funcs_dict, mos_names)
    # Automatically call the render function for the page
    mos_code += "\nif __name__ == '__main__':\n    render_mos_page()\n"
    with open('pages/2_MOS.py', 'w', encoding='utf-8') as f:
        f.write(mos_code)

    # 5. pages/3_Checkin.py
    checkin_names = ['render_checkin_page']
    checkin_code = "import streamlit as st\nimport pandas as pd\nimport datetime\nfrom db import save_field_checkin, get_field_checkins\nfrom translations import get_t\n\n"
    checkin_code += extract_funcs(lines, funcs_dict, checkin_names)
    checkin_code += "\nif __name__ == '__main__':\n    render_checkin_page()\n"
    with open('pages/3_Checkin.py', 'w', encoding='utf-8') as f:
        f.write(checkin_code)
        
    # 6. pages/4_Gamification.py
    gamification_names = ['render_gamification_dashboard']
    gamification_code = "import streamlit as st\nimport pandas as pd\nfrom translations import get_t\n\n"
    gamification_code += extract_funcs(lines, funcs_dict, gamification_names)
    gamification_code += "\nif __name__ == '__main__':\n    render_gamification_dashboard()\n"
    with open('pages/4_Gamification.py', 'w', encoding='utf-8') as f:
        f.write(gamification_code)
        
    print("Files created successfully.")

if __name__ == '__main__':
    main()
