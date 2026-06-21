"""
HZX-采购数据管家 配置文件
"""
import os, sys

# ===== 应用信息 =====
APP_NAME = "HZX-采购数据管家"
VERSION = "V1.0.0"
DEVELOPER = "IT-钟"
PHONE = "18072740843"

# ===== 路径 =====
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SOP_DOCS_DIR = os.path.join(BASE_DIR, "SOP_docs")
FILE_LIST_PATH = os.path.join(BASE_DIR, "file_list.json")

# ===== 默认值 =====
DEFAULT_YEAR1 = 25
DEFAULT_YEAR2 = 26
DEFAULT_ENABLE_BORDER = False

# ===== 工作表名称模板 =====
SHEET_CHANNEL_AMAZON = "{year2}年出货渠道对比-Amazon"
SHEET_CHANNEL_COMPARE = "{year1}年、{year2}年出货渠道对比"
SHEET_THREE_PLATFORMS = "{year1}年、{year2}年采购量和采购总额对比-三平台"
SHEET_CUSTOMS_AMAZON = "{year1}年、{year2}年报关占比表--Amazon"

# ===== 保留的工作表 =====
def get_sheets_keep(year1, year2, month1, month2):
    return [
        f"{year1}.{month1}", f"{year2}.{month2}",
        SHEET_CHANNEL_AMAZON.format(year2=year2),
        SHEET_CHANNEL_COMPARE.format(year1=year1, year2=year2),
        SHEET_THREE_PLATFORMS.format(year1=year1, year2=year2),
        SHEET_CUSTOMS_AMAZON.format(year1=year1, year2=year2),
    ]
