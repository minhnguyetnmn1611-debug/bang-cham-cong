import pandas as pd
from log_config import logger
import datetime
import calendar
import math

def to_time_obj(val):
    try:
        is_na = pd.isna(val)
    except (TypeError, ValueError):
        is_na = False
    if is_na or val == "":
        return None
    if isinstance(val, datetime.time):
        return val
    if isinstance(val, datetime.datetime):
        return val.time()
    if isinstance(val, (int, float)):
        if 0 <= val < 1:
            total_seconds = int(val * 86400)
            return datetime.time(hour=total_seconds//3600, minute=(total_seconds%3600)//60, second=total_seconds%60)
    if isinstance(val, str):
        val = val.strip()
        if not val or val.lower() in ('nan', 'nat', 'none'):
            return None
        try:
            return pd.to_datetime(val).time()
        except Exception as e: logger.warning(f"Lỗi: {e}", exc_info=True); return None
    return None


def time_to_float(t):
    if t is None:
        return 0.0
    return t.hour + t.minute / 60.0 + t.second / 3600.0


def custom_round(val):
    if val is None or (isinstance(val, float) and math.isnan(val)) or str(val).lower() in ['nan', 'nat', 'none', '']:
        return 0.0
    try:
        val = float(val)
    except (ValueError, TypeError):
        return 0.0
        
    integer_part = math.floor(val)
    decimal_part = val - integer_part
    
    decimal_100 = round(decimal_part * 100)
    tens = decimal_100 // 10
    ones = decimal_100 % 10
    
    if ones < 5:
        final_decimal = tens / 10.0
    elif ones == 5:
        final_decimal = (tens * 10 + 5) / 100.0
    else:
        final_decimal = (tens + 1) / 10.0
        
    result = integer_part + final_decimal
    # Ensure it's clean (e.g. 2.4 instead of 2.40000000001)
    return round(result, 2)


def format_gio_lam(val):
    if val is None or (isinstance(val, float) and math.isnan(val)) or str(val).lower() in ['nan', 'nat', 'none', '']:
        return 0
    val = custom_round(val)
    return int(val) if val == int(val) else val


def format_gio_lam_str(val):
    return str(format_gio_lam(val))


def calculate_working_hours(time_in, time_out, start_chuan=8.0, end_chuan=17.0, lunch_start=12.0, lunch_end=13.0, max_hours=8.0):
    vao_trong = time_in is None or str(time_in).strip().lower() in ['', 'nan', 'none', 'nat']
    ra_trong = time_out is None or str(time_out).strip().lower() in ['', 'nan', 'none', 'nat']
    result = {'admin_hours': 0.0, 'tong_gio': 0.0, 'di_tre': 0, 've_som': 0, 'ot': 0.0, 'is_chieu': False}
    if vao_trong or ra_trong:
        return result
    t_in = to_time_obj(time_in)
    t_out = to_time_obj(time_out)
    if t_in is None or t_out is None:
        return result
    in_f = time_to_float(t_in)
    out_f = time_to_float(t_out)
    if out_f < in_f:
        result['tong_gio'] = -1
        return result
        
    is_chieu = in_f >= lunch_start
    work_start = max(in_f, start_chuan)
    work_end = min(out_f, end_chuan)
    
    if work_end <= work_start:
        admin_hours = 0.0
    else:
        # Trừ đi khoảng thời gian nghềEtrưa nếu có giao thoa
        overlap_lunch = max(0.0, min(work_end, lunch_end) - max(work_start, lunch_start))
        admin_hours = (work_end - work_start) - overlap_lunch
            
    admin_hours = min(admin_hours, max_hours)
    admin_hours = custom_round(admin_hours)
    
    result = {'is_chieu': is_chieu}
    result['admin_hours'] = admin_hours
    result['di_tre'] = int(max(0, (in_f - start_chuan) * 60)) if in_f > start_chuan else 0
    result['ve_som'] = int(max(0, (end_chuan - out_f) * 60)) if out_f < end_chuan and out_f > start_chuan else 0
    # Tăng ca (OT): Cứ sau giờ về chuẩn (17:00) đều tính là OT
    raw_ot = max(0.0, out_f - end_chuan)
    result['ot'] = custom_round(raw_ot)
    result['tong_gio'] = admin_hours
    return result


def clean_date(val):
    if pd.isna(val):
        return None
    try:
        if isinstance(val, datetime.datetime):
            return val
        return pd.to_datetime(str(val).strip(), dayfirst=True)
    except Exception as e: logger.warning(f"Lỗi: {e}", exc_info=True); return None


def get_fixed_holidays_for_years(years):
    dates = []
    for y in sorted(years):
        for (mo, day) in [(1,1), (4,30), (5,1), (9,2)]:
            dates.append(datetime.date(y, mo, day))
    return sorted(dates)


def is_workday_func(d_obj):
    """Trả vềETrue nếu là ngày làm việc.
    - Thứ 7 cuối cùng của tháng = bắt buộc làm việc.
    - Các Thứ 7 khác và Chủ nhật = nghềE
    """
    if pd.isna(d_obj): return True
    try:
        d = d_obj.date() if hasattr(d_obj, 'date') else d_obj
        wd = d.weekday()
        if wd == 6: return False  # Chủ nhật: nghỉ
        if wd == 5:
            if is_last_saturday_of_month(d_obj):
                return True
            return False  # Các Thứ 7 khác: nghỉ
        return True  # Thứ 2 - Thứ 6: luôn làm việc
    except Exception as e: logger.warning(f"Lỗi: {e}", exc_info=True); return True


def is_last_saturday_of_month(d_obj):
    """
    Trả về True nếu ngày này là Thứ 7 cuối cùng của tháng.
    """
    if pd.isna(d_obj): return False
    try:
        import datetime
        d = d_obj.date() if hasattr(d_obj, 'date') else d_obj
        if d.weekday() != 5:  # Không phải Thứ 7
            return False
        
        # Thứ 7 cuối cùng của tháng: Nếu cộng thêm 7 ngày mà sang tháng mới, thì đây là Thứ 7 cuối cùng
        next_week = d + datetime.timedelta(days=7)
        if next_week.month != d.month:
            return True
        return False
    except:
        return False


def calculate_working_days(start_date, end_date, holidays=None, makeups=None):
    if holidays is None:
        holidays = set()
    if makeups is None:
        makeups = set()
    total = 0
    holiday_info = []
    makeup_info = []
    current = start_date
    while current <= end_date:
        is_wd = is_workday_func(current)
        if current in makeups:
            total += 1
            makeup_info.append({
                "date": current,
                "weekday_label": ["Thứ 2","Thứ 3","Thứ 4","Thứ 5","Thứ 6","Thứ 7","Chủ nhật"][current.weekday()]
            })
        elif current in holidays:
            holiday_info.append({
                "date": current,
                "weekday_label": ["Thứ 2","Thứ 3","Thứ 4","Thứ 5","Thứ 6","Thứ 7","Chủ nhật"][current.weekday()],
                "is_workday": is_wd
            })
        elif is_wd:
            total += 1
        current += datetime.timedelta(days=1)
    return total, holiday_info, makeup_info


def auto_detect_columns(df):
    import re
    mapping = {}
    for col in df.columns:
        col_lower = str(col).strip().lower().replace('\n', ' ').replace('\r', ' ')
        col_lower = re.sub(' +', ' ', col_lower)
        if 'mã' in col_lower and 'nv' not in col_lower.replace('mã',''):
            mapping['ma_nv'] = col
        elif col_lower in ['tên nhân viên', 'tên nv', 'hềEtên', 'tên']:
            mapping['ten_nv'] = col
        elif 'ngày' in col_lower and 'ca' not in col_lower:
            mapping['ngay'] = col
        elif col_lower in ['vào', 'vao', 'vào 1', 'vao 1', 'vào1', 'vao1', 'thời gian vào', 'thoi gian vao', 'vào ca']:
            mapping['gio_vao'] = col
        elif col_lower in ['ra', 'ra 1', 'ra1', 'thời gian ra', 'thoi gian ra', 'ra ca']:
            mapping['gio_ra'] = col
        elif col_lower in ['chức vụ', 'chuc vu', 'vềEtrí', 'title']:
            mapping['chuc_vu'] = col
        elif col_lower in ['phòng ban', 'phong ban', 'bềEphận', 'bo phan', 'department']:
            mapping['phong_ban'] = col
        elif col_lower in ['ot', 'giềEot', 'gio ot', 'tăng ca', 'tang ca', 'overtime', 'tc1', 'tc2', 'tc3', 'tổng tăng ca']:
            mapping['ot'] = col
        elif col_lower in ['vào trễ', 'vao tre', 'đi trễ', 'di tre', 'late in', 'late']:
            mapping['di_tre'] = col
        elif col_lower in ['ra sớm', 'ra som', 'về sớm', 've som', 'early out', 'early']:
            mapping['ve_som'] = col
        elif col_lower in ['tổng giờ', 'tong gio', 'tổng số giờ', 'total hours', 'tong_gio', 'tổng']:
            mapping['tong_gio'] = col
    return mapping
