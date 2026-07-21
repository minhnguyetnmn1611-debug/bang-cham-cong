import ast

def get_function_lines(filename):
    with open(filename, 'r', encoding='utf-8') as f:
        source = f.read()
    tree = ast.parse(source)
    functions = {}
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            functions[node.name] = [node.lineno, node.end_lineno]
    return functions

def main():
    funcs_dict = get_function_lines('app.py')
    
    to_remove = [
        'to_time_obj', 'time_to_float', 'format_gio_lam', 'format_gio_lam_str', 
        'calculate_working_hours', 'clean_date', 'get_fixed_holidays_for_years', 
        'is_workday_func', 'is_last_saturday_of_month', 'calculate_working_days', 
        'auto_detect_columns', 'find_header_row', 'parse_excel_file', 
        'export_excel_tong_hop', 'render_chatbot', 'summarize_tasks_with_ai', 
        'extract_ma_nv_from_filename', 'extract_ten_nv_from_filename', 'parse_mos_file', 
        'tong_hop_mos', 'render_mos_page', 'render_checkin_page', 'render_gamification_dashboard',
        'init_db', 'get_company_emp_options', 'get_company_emp_dict', 
        'save_field_checkin', 'get_field_checkins', 'save_to_db', 'load_saved_api_key'
    ]
    
    ranges = []
    for name in to_remove:
        if name in funcs_dict:
            ranges.append(funcs_dict[name])
            
    ranges.sort(key=lambda x: x[0], reverse=True)
    
    with open('app.py', 'r', encoding='utf-8') as f:
        lines = f.readlines()
        
    for r in ranges:
        start, end = r
        del lines[start-1:end]
        
    imports = """from utils import *
from excel_export import *
from ai_chat import render_chatbot
from db import init_db, get_company_emp_options, get_company_emp_dict, save_field_checkin, get_field_checkins, save_to_db

"""
    # Insert after standard imports, around line 18
    lines.insert(17, imports)
    
    with open('app.py', 'w', encoding='utf-8') as f:
        f.writelines(lines)
        
    print("app.py cleaned.")

if __name__ == '__main__':
    main()
