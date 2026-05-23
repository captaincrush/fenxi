# utils.py
import os
import re
import datetime
from decimal import Decimal
from config import *

# def parse_date_from_filename(filename):
#     """从文件名的开始处提取 月-日 (M-D 或 MM-DD)，返回 datetime.date（使用当前年份）"""
#     m = re.match(r"^(\d{1,2})-(\d{1,2})", filename)
#     if not m:
#         return None
#     month, day = int(m.group(1)), int(m.group(2))
#     year = datetime.date.today().year
#     try:
#         return datetime.date(year, month, day)
#     except ValueError:
#         return None
    
def parse_date_from_filename(filename):
    m = re.match(r"^(\d{1,2})-(\d{1,2})", filename)
    if not m: return None
    
    month, day = int(m.group(1)), int(m.group(2))
    today = datetime.date.today()
    year = today.year
    
    # 如果今天是年初（如1月），而文件名是年尾（如12月），判定为去年
    if today.month < 3 and month > 10:
        year -= 1
        
    try:
        return datetime.date(year, month, day)
    except ValueError:
        return None


def find_xlsx_files_with_dates(dirpath):
    """查找目录中的xlsx文件并解析日期"""
    files = []
    for f in os.listdir(dirpath):
        if not f.lower().endswith(".xlsx"):
            continue
        d = parse_date_from_filename(f)
        if d:
            files.append((f, d))
    files.sort(key=lambda x: x[1])
    return files

def days_between(d1, d2):
    """计算两个日期之间的天数"""
    return (d2 - d1).days

def adjust_to_7_days(diff_value, actual_days):
    """将差值调整到7天基准"""
    if actual_days <= 0:
        return None

    diff = Decimal(str(diff_value))
    days = Decimal(str(actual_days))
    if days == 7:
        return diff
    else:
        return diff / days * Decimal('7')

def calc_diff_or_cumulative(val_curr, date_curr, val_prev=None, date_prev=None):
    """
    如果 date_curr 和 date_prev 在同一个月，则返回差分值 (val_curr - val_prev)，并返回天数。
    如果跨月，则返回累计值 val_curr，天数为 date_curr.day。
    """
    if date_prev and date_curr.month == date_prev.month:
        # 同月：差分
        days = (date_curr - date_prev).days
        return val_curr - val_prev, days
    else:
        # 跨月：直接用累计
        return val_curr, date_curr.day

def safe_round_units(x):
    """销量（订单量）不保留小数，四舍五入为整数"""
    if x is None:
        return 0
    return int(round(x))

def format_percent_with_sign(value):
    """保留两位小数，加正负号百分号"""
    if value is None:
        return "0.00%"
    sign = "+" if value > 0 else ""
    return f"{sign}{value:.2f}%"

def parse_brand_country(filename):
    """提取品牌名、国家"""
    lower_name = filename.lower()

    # 提取品牌名
    brand = None
    if "goiteia" in lower_name:
        brand = "Goiteia"
    elif "bloom" in lower_name:
        brand = "Bloom"
    elif "awode" in lower_name:
        brand = "Awode"
    elif "maltgoods" in lower_name:
        brand = "MaltGoods"

    # 匹配国家代码
    match = re.findall(r"-([A-Z]{2})", filename)
    country = match[-1] if match else None

    print(f"对应国家和品牌：{country, brand}")
    return brand, country

def parse_comment_week_date(text):
    """解析批注，返回 (week_num, date) tuple"""
    if not text:
        return None, None
    week_match = re.search(r"新上架[（(]第\s*(\d+)\s*周[）)]", text)
    date_match = re.search(r"@(\d{4}-\d{2}-\d{2})", text)
    week_num = int(week_match.group(1)) if week_match else None
    dt = datetime.datetime.strptime(date_match.group(1), "%Y-%m-%d").date() if date_match else None
    return week_num, dt
