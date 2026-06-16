"""
配置文件 - 月度采购数据对比SOP 项目变量
"""
# ===== 年份变量 =====
YEAR1 = 25  # 基准年份（较早年份）
YEAR2 = 26  # 对比年份（较新年份）

# ===== 文件路径 =====
SOURCE_FILE = r'e:\ZY2026\采购需求-月度采购数据对比SOP\（IT试验版初版）26.6.11--25.5和26.5采购数据对比.xlsx'
OUTPUT_FILE = r'e:\ZY2026\采购需求-月度采购数据对比SOP\（IT试验版初版）26.6.11--25.5和26.5采购数据对比-整合输出-v6.xlsx'

# ===== 格式设置 =====
ENABLE_BORDER = False  # 数据分析表边框开关，True=有边框，False=无边框

# ===== 工作表名称 =====
# 源数据工作表（格式: {year}.{month}）
TEMPLATE_SHEET_PATTERN = "{year}.{month}"

# 分析工作表
SHEET_CHANNEL_AMAZON = "{year2}年出货渠道对比-Amazon"           # e.g. "26年出货渠道对比-Amazon"
SHEET_CHANNEL_COMPARE = "{year1}年、{year2}年出货渠道对比"       # e.g. "25年、26年出货渠道对比"
SHEET_THREE_PLATFORMS = "{year1}年、{year2}年采购量和采购总额对比-三平台"
SHEET_CUSTOMS_AMAZON = "{year1}年、{year2}年报关占比表--Amazon"
