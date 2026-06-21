"""
HZX-采购数据管家 轻量客户端 配置
"""
import os, sys, json

# ===== 应用信息 =====
APP_NAME = "HZX-采购数据管家"
VERSION = "V1.0.0-轻量版"
DEVELOPER = "IT-钟"
PHONE = "18072740843"

# ===== 路径 =====
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SOP_DOCS_DIR = os.path.join(BASE_DIR, "SOP_docs")
FILE_LIST_PATH = os.path.join(BASE_DIR, "file_list.json")
SETTINGS_PATH = os.path.join(BASE_DIR, "settings.json")

# ===== NAS 配置 =====
# 使用 SMB 网络共享路径
NAS_PATH = r"\\192.168.40.3\钟正洋"
NAS_USER = "HZXzhengyang"
NAS_PASS = "hzx001456@"

# NAS 子目录（按需调整）
NAS_SOURCE_DIR = NAS_PATH  # 源文件目录（根目录）
NAS_OUTPUT_DIR = NAS_PATH  # 输出文件目录（根目录）

# ===== 默认值 =====
DEFAULT_YEAR1 = 25
DEFAULT_YEAR2 = 26
DEFAULT_ENABLE_BORDER = False
DEFAULT_THEME = "light"

# ===== 设置持久化 =====
def load_settings():
    defaults = {
        "year1": DEFAULT_YEAR1, "year2": DEFAULT_YEAR2,
        "enable_border": DEFAULT_ENABLE_BORDER, "theme": DEFAULT_THEME,
        "nas_path": NAS_PATH, "nas_user": NAS_USER, "nas_pass": NAS_PASS,
    }
    if os.path.exists(SETTINGS_PATH):
        try:
            with open(SETTINGS_PATH, 'r', encoding='utf-8') as f:
                saved = json.load(f); defaults.update(saved)
        except: pass
    return defaults

def save_settings(settings: dict):
    try:
        with open(SETTINGS_PATH, 'w', encoding='utf-8') as f:
            json.dump(settings, f, ensure_ascii=False, indent=2)
    except: pass

# ===== NAS 映射工具 =====
def mount_nas(nas_path, user, password):
    """连接 NAS 共享目录（Windows 自动凭证管理）"""
    import subprocess
    try:
        # 尝试先删除旧连接
        subprocess.run(f'net use "{nas_path}" /delete /y',
                       shell=True, capture_output=True, timeout=5)
    except: pass
    try:
        result = subprocess.run(
            f'net use "{nas_path}" "{password}" /user:"{user}" /persistent:no',
            shell=True, capture_output=True, timeout=10, encoding='utf-8')
        return result.returncode == 0
    except:
        return False

def ensure_nas_connected(settings):
    """确保 NAS 已连接，未连接则自动映射"""
    nas_path = settings.get("nas_path", NAS_PATH)
    if os.path.exists(nas_path):
        return True  # 已连接
    user = settings.get("nas_user", NAS_USER)
    password = settings.get("nas_pass", NAS_PASS)
    return mount_nas(nas_path, user, password)
