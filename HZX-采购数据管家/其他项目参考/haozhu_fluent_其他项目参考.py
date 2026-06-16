"""
豪猪接码平台 - PyQt5 + Fluent Widgets 精致版
"""
import sys, os, json, time, requests, warnings
from urllib3.exceptions import InsecureRequestWarning
warnings.filterwarnings("ignore", category=InsecureRequestWarning)

CONFIG_FILE = "config.json"
CONFIG = {
    "server": "https://api.haozhuyun.com", "user": "", "pass": "",
    "sid": "", "token": "", "skip_login": True,
    "poll_interval": 15, "max_wait": 180, "app_user": "", "app_pass": "",
    "gh_repo": "", "gh_token": "", "proxy_url": "",
}

def load_config():
    global CONFIG
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            CONFIG.update(json.load(f))
    except: pass

def save_config():
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(CONFIG, f, ensure_ascii=False, indent=2)
    except: pass

def load_env():
    env_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    if not os.path.exists(env_file): return
    try:
        with open(env_file, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line: continue
                k, v = line.split("=", 1)
                k, v = k.strip(), v.strip()
                if k == "API_USER" and v: CONFIG.setdefault("user", v)
                elif k == "API_PASSWORD" and v: CONFIG.setdefault("pass", v)
    except: pass

load_env()
load_config()

def api_request(url, params=None):
    try:
        r = requests.get(url, params=params, timeout=15, verify=False)
        r.raise_for_status(); return r.json()
    except: return None

def api_login(user, passwd):
    url = f"{CONFIG['server']}/sms/"
    d = api_request(url, {"api": "login", "user": user, "pass": passwd})
    if d and (d.get("code") == 0 or d.get("code") == "0"):
        return d.get("token", "")
    return None

def api_get_phone(token, sid):
    d = api_request(f"{CONFIG['server']}/sms/", {"api":"getPhone","token":token,"sid":sid})
    if d and (d.get("code") == "0" or d.get("code") == 0): return d
    return None

def api_get_message(token, sid, phone):
    d = api_request(f"{CONFIG['server']}/sms/", {"api":"getMessage","token":token,"sid":sid,"phone":phone})
    if d and (d.get("code") == "0" or d.get("code") == 0): return d
    return None

def main():
    from PyQt5.QtCore import Qt
    from PyQt5.QtWidgets import QApplication
    from PyQt5.QtNetwork import QLocalServer, QLocalSocket
    import ctypes, subprocess
    try: ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except: pass
    app = QApplication(sys.argv)

    from PyQt5 import QtWidgets
    from PyQt5.QtCore import pyqtSignal, QThread
    from PyQt5.QtWidgets import QVBoxLayout, QHBoxLayout, QLabel, QWidget, QStackedWidget, QDialog, QScrollArea
    from PyQt5.QtGui import QIcon
    from qfluentwidgets import (
        NavigationItemPosition, FluentIcon, FluentWindow, IndeterminateProgressRing,
        TitleLabel, SubtitleLabel, LineEdit, PushButton, CheckBox,
        PrimaryPushButton, SimpleCardWidget, InfoBar,
        PillPushButton, TransparentToolButton, StrongBodyLabel,
    )
    import hashlib, secrets, base64
    from cryptography.fernet import Fernet
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

    CONFIGS_DIR = "configs"

    # ==================== 加密工具 ====================
    class CryptoUtil:
        @staticmethod
        def hash_password(password):
            salt = secrets.token_hex(16)
            dk = hashlib.pbkdf2_hmac("sha256", password.encode(), salt.encode(), 600_000)
            return f"pbkdf2$600000${salt}${dk.hex()}"
        @staticmethod
        def verify_password(password, stored):
            parts = stored.split("$")
            if len(parts) != 4 or parts[0] != "pbkdf2": return False
            _, it, salt, h = parts
            dk = hashlib.pbkdf2_hmac("sha256", password.encode(), salt.encode(), int(it))
            return dk.hex() == h
        @staticmethod
        def _derive_key(password, salt=None):
            if salt is None: salt = secrets.token_bytes(16)
            kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=600_000)
            key = base64.urlsafe_b64encode(kdf.derive(password.encode()))
            return key, salt
        @staticmethod
        def encrypt_data(data, password):
            key, salt = CryptoUtil._derive_key(password)
            f = Fernet(key)
            ct = f.encrypt(json.dumps(data, ensure_ascii=False).encode())
            return salt.hex() + "$" + ct.decode()
        @staticmethod
        def decrypt_data(encrypted, password):
            try:
                sh, cb = encrypted.split("$", 1)
                key, _ = CryptoUtil._derive_key(password, bytes.fromhex(sh))
                return json.loads(Fernet(key).decrypt(cb.encode()).decode())
            except: return None

    class LoginState:
        logged_in = False; current_user = ""; _login_password = ""
        api_server = api_user = api_pass = api_sid = api_token = app_user = app_pass = ""
        @classmethod
        def login(cls, user, password, s, u, p, sid, token, au, ap):
            cls.logged_in = True; cls.current_user = user; cls._login_password = password
            cls.api_server = s; cls.api_user = u; cls.api_pass = p
            cls.api_sid = sid; cls.api_token = token; cls.app_user = au; cls.app_pass = ap
        @classmethod
        def logout(cls):
            for a in ["logged_in","current_user","_login_password","api_server","api_user",
                      "api_pass","api_sid","api_token","app_user","app_pass"]:
                setattr(cls, a, "" if a != "logged_in" else False)
        @classmethod
        def to_config(cls):
            return {k: getattr(cls, k) for k in ["api_server","api_user","api_pass",
                    "api_sid","api_token","app_user","app_pass"]}

    def get_user_path(username): return os.path.join(CONFIGS_DIR, f"{username}.json")
    def save_user_config(username, password, data):
        os.makedirs(CONFIGS_DIR, exist_ok=True)
        with open(get_user_path(username), "w", encoding="utf-8") as f:
            json.dump({"password_hash": CryptoUtil.hash_password(password),
                       "encrypted_data": CryptoUtil.encrypt_data(data, password)}, f)
    def load_user_config(username, password):
        path = get_user_path(username)
        if not os.path.exists(path): return None
        with open(path, "r", encoding="utf-8") as f:
            m = json.load(f)
        if not CryptoUtil.verify_password(password, m["password_hash"]): return None
        return CryptoUtil.decrypt_data(m["encrypted_data"], password)
    def get_root_config():
        defaults = {"current_user":"","users":[],"poll_interval":15,"max_wait":180,"skip_login":True,
                    "gh_repo":"","gh_token":"","proxy_url":""}
        if os.path.exists(CONFIG_FILE):
            try:
                d = json.load(open(CONFIG_FILE, "r", encoding="utf-8"))
                for k in defaults: d.setdefault(k, defaults[k])
                return d
            except: pass
        return defaults
    def save_root_config(data):
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        # 同步到 CONFIG
        for k in ["gh_repo","gh_token","proxy_url","poll_interval","max_wait","skip_login"]:
            if k in data: CONFIG[k] = data[k]

    # ---------- 5天免登录 ----------
    SESSION_BASE = os.path.dirname(sys.executable) if getattr(sys,'frozen',False) else os.path.dirname(os.path.abspath(__file__))
    SESSION_FILE = os.path.join(SESSION_BASE, ".session_cache")
    SESSION_KEY = base64.urlsafe_b64encode(hashlib.sha256(b"HZX-Session-Key-2024!!").digest())
    def save_session(username, password):
        try:
            data = json.dumps({"user": username, "password": password, "ts": time.time()})
            f = Fernet(SESSION_KEY)
            with open(SESSION_FILE, "w") as fh: fh.write(f.encrypt(data.encode()).decode())
        except: pass
    def load_session():
        try:
            if not os.path.exists(SESSION_FILE): return None, None
            with open(SESSION_FILE, "r") as fh: cipher = fh.read()
            data = json.loads(Fernet(SESSION_KEY).decrypt(cipher.encode()).decode())
            if time.time() - data["ts"] > 432000: os.remove(SESSION_FILE); return None, None
            return data["user"], data["password"]
        except: return None, None
    def clear_session():
        try:
            if os.path.exists(SESSION_FILE): os.remove(SESSION_FILE)
        except: pass

    def creds(key):
        if LoginState.logged_in:
            return getattr(LoginState, key, CONFIG.get(key, ""))
        return CONFIG.get(key, "")

    def _make_scrollable(widget, layout):
        container = QWidget(); container.setLayout(layout)
        sa = QScrollArea(); sa.setWidget(container); sa.setWidgetResizable(True); sa.setFrameShape(0)
        outer = QVBoxLayout(widget); outer.setContentsMargins(0,0,0,0); outer.addWidget(sa)
        return container

    # ==================== 登录对话框 ====================
    class LoginDialog(QDialog):
        def __init__(self, parent=None):
            super().__init__(parent)
            self.setWindowTitle("登录"); self.setFixedSize(380, 330)
            self.result_data = None; self.remember = False; self._setup_ui()
        def _setup_ui(self):
            l = QVBoxLayout(self); l.setSpacing(14); l.setContentsMargins(32,24,32,24)
            l.addWidget(TitleLabel("登录", self)); l.addSpacing(6)
            self.user_input = LineEdit(self); self.user_input.setPlaceholderText("用户名"); l.addWidget(self.user_input)
            self.pass_input = LineEdit(self); self.pass_input.setPlaceholderText("密码")
            self.pass_input.setEchoMode(LineEdit.Password); l.addWidget(self.pass_input)
            root = get_root_config(); users = root.get("users", [])
            if users:
                h = StrongBodyLabel("已有账号: "+", ".join(users), self)
                h.setStyleSheet("color:#888;font-size:12px"); l.addWidget(h)
            br = QHBoxLayout()
            self.login_btn = PrimaryPushButton("登录", self); self.login_btn.clicked.connect(self._do_login); br.addWidget(self.login_btn)
            self.reg_btn = PushButton("注册", self); self.reg_btn.clicked.connect(self._do_register); br.addWidget(self.reg_btn)
            l.addLayout(br)
            self.remember_cb = CheckBox("5天免登录", self); self.remember_cb.setChecked(True); l.addWidget(self.remember_cb)
            l.addStretch()
        def _do_login(self):
            u = self.user_input.text().strip(); p = self.pass_input.text().strip()
            if not u or not p: InfoBar.warning("","请输入用户名和密码",duration=2000,parent=self); return
            data = load_user_config(u, p)
            if data is None:
                if os.path.exists(get_user_path(u)): InfoBar.error("","密码错误",duration=2000,parent=self)
                else: InfoBar.warning("","用户不存在，请先注册",duration=2000,parent=self)
                return
            self.result_data = {"user":u,"password":p,"data":data}; self.remember = self.remember_cb.isChecked(); self.accept()
        def _do_register(self):
            d = RegisterDialog(self)
            if d.exec() == QDialog.Accepted:
                self.user_input.setText(d.result_user); self.pass_input.setText(d.result_pass)
                InfoBar.success("","注册成功，请点击登录",duration=2000,parent=self)

    class RegisterDialog(QDialog):
        def __init__(self, parent=None):
            super().__init__(parent)
            self.setWindowTitle("注册"); self.setFixedSize(380,300)
            self.result_user=""; self.result_pass=""; self._setup_ui()
        def _setup_ui(self):
            l = QVBoxLayout(self); l.setSpacing(14); l.setContentsMargins(32,24,32,24)
            l.addWidget(TitleLabel("注册新账号",self)); l.addSpacing(6)
            self.user_input = LineEdit(self); self.user_input.setPlaceholderText("用户名"); l.addWidget(self.user_input)
            self.pass_input = LineEdit(self); self.pass_input.setPlaceholderText("密码")
            self.pass_input.setEchoMode(LineEdit.Password); l.addWidget(self.pass_input)
            self.confirm_input = LineEdit(self); self.confirm_input.setPlaceholderText("确认密码")
            self.confirm_input.setEchoMode(LineEdit.Password); l.addWidget(self.confirm_input)
            br = QHBoxLayout()
            reg_btn = PrimaryPushButton("注册",self); reg_btn.clicked.connect(self._do_register); br.addWidget(reg_btn)
            cancel_btn = PushButton("取消",self); cancel_btn.clicked.connect(self.reject); br.addWidget(cancel_btn)
            l.addLayout(br); l.addStretch()
        def _do_register(self):
            u = self.user_input.text().strip(); p = self.pass_input.text().strip(); c = self.confirm_input.text().strip()
            if not u or not p: InfoBar.warning("","请填写用户名和密码",duration=2000,parent=self); return
            if p != c: InfoBar.warning("","两次密码不一致",duration=2000,parent=self); return
            if os.path.exists(get_user_path(u)): InfoBar.warning("","用户已存在",duration=2000,parent=self); return
            data = {"api_server":CONFIG.get("server",""),"api_user":"","api_pass":"","api_sid":"","api_token":"","app_user":u,"app_pass":p}
            save_user_config(u, p, data)
            root = get_root_config()
            if u not in root["users"]: root["users"].append(u)
            save_root_config(root)
            self.result_user=u; self.result_pass=p; self.accept()

    class LoginPromptWidget(QWidget):
        def __init__(self, title="", message="", parent=None):
            super().__init__(parent)
            l = QVBoxLayout(self); l.setAlignment(Qt.AlignCenter); l.setSpacing(16)
            l.addWidget(TitleLabel(title, self))
            m = StrongBodyLabel(message, self); m.setStyleSheet("color:#888;"); m.setAlignment(Qt.AlignCenter); m.setWordWrap(True)
            l.addWidget(m)
            self.login_btn = PrimaryPushButton("登录 / 注册", self); self.login_btn.setFixedWidth(200); l.addWidget(self.login_btn)
            l.addStretch()

    # ==================== CopyableCard ====================
    class CopyableCard(SimpleCardWidget):
        def __init__(self, title="", value="", parent=None):
            super().__init__(parent); self._value=value; self.setBorderRadius(12)
            l=QVBoxLayout(self); l.setContentsMargins(20,14,20,14); l.setSpacing(6)
            tr=QHBoxLayout(); self.title_label=StrongBodyLabel(title,self); tr.addWidget(self.title_label); tr.addStretch(); l.addLayout(tr)
            vr=QHBoxLayout(); self.value_label=TitleLabel(value or "—",self); self.value_label.setWordWrap(True); vr.addWidget(self.value_label,1)
            self.copy_btn=TransparentToolButton(FluentIcon.COPY,self); self.copy_btn.setToolTip("复制")
            self.copy_btn.clicked.connect(self._copy); self.copy_btn.setFixedSize(32,32); vr.addWidget(self.copy_btn); l.addLayout(vr)
        def set_value(self,t): self._value=t; self.value_label.setText(t or "—")
        def _copy(self):
            if self._value: QtWidgets.QApplication.clipboard().setText(self._value); InfoBar.success("","已复制到剪贴板",duration=1500,parent=self)

    # ==================== Workers ====================
    class AutoFlowWorker(QThread):
        log_signal=pyqtSignal(str); result_signal=pyqtSignal(dict); error_signal=pyqtSignal(str)
        progress_signal=pyqtSignal(str,str); finished_signal=pyqtSignal()
        def __init__(self,parent=None): super().__init__(parent); self._stop_flag=False
        def stop(self): self._stop_flag=True
        def run(self):
            try:
                if self._stop_flag: return
                token = creds("api_token") or CONFIG.get("token","")
                if not token:
                    u=creds("api_user") or CONFIG.get("user","")
                    p=creds("api_pass") or CONFIG.get("pass","")
                    token=api_login(u,p)
                    if not token: self.error_signal.emit("登录失败"); return
                    if LoginState.logged_in: LoginState.api_token=token
                    else: CONFIG["token"]=token; save_config()
                elif not CONFIG.get("skip_login",True):
                    u=creds("api_user") or CONFIG.get("user","")
                    p=creds("api_pass") or CONFIG.get("pass","")
                    token=api_login(u,p)
                    if not token: self.error_signal.emit("登录失败"); return
                    if LoginState.logged_in: LoginState.api_token=token
                    else: CONFIG["token"]=token; save_config()
                if self._stop_flag: return
                sid=creds("api_sid") or CONFIG.get("sid","")
                phone_data=api_get_phone(token,sid)
                if not phone_data: self.error_signal.emit("获取号码失败"); return
                phone=phone_data.get("phone","")
                self.progress_signal.emit(phone,"等待验证码...")
                result=self._poll(token,sid,phone)
                if self._stop_flag: return
                if result:
                    self.progress_signal.emit(phone,result.get("yzm",""))
                    self.result_signal.emit({"phone":phone,"yzm":result.get("yzm",""),"sms":result.get("sms","")})
                else: self.error_signal.emit(f"超时：{phone} 未收到验证码")
            except Exception as e: self.error_signal.emit(str(e))
            finally: self.finished_signal.emit()
        def _poll(self,token,sid,phone):
            iv,max_w=CONFIG["poll_interval"],CONFIG["max_wait"]; start=time.time(); n=0
            while time.time()-start<max_w:
                if self._stop_flag: return None
                n+=1
                r=api_get_message(token,sid,phone)
                if r: return r
                for _ in range(iv):
                    if self._stop_flag: return None
                    time.sleep(1)
            return None

    class QueryCodeWorker(QThread):
        log_signal=pyqtSignal(str); result_signal=pyqtSignal(dict); error_signal=pyqtSignal(str)
        progress_signal=pyqtSignal(str); finished_signal=pyqtSignal()
        def __init__(self,phone,parent=None): super().__init__(parent); self.phone=phone; self._stop_flag=False
        def stop(self): self._stop_flag=True
        def run(self):
            try:
                if self._stop_flag: return
                token = creds("api_token") or CONFIG.get("token","")
                if not token:
                    u=creds("api_user") or CONFIG.get("user","")
                    p=creds("api_pass") or CONFIG.get("pass","")
                    token=api_login(u,p)
                    if not token: self.error_signal.emit("登录失败"); return
                    if LoginState.logged_in: LoginState.api_token=token
                    else: CONFIG["token"]=token; save_config()
                elif not CONFIG.get("skip_login",True):
                    u=creds("api_user") or CONFIG.get("user","")
                    p=creds("api_pass") or CONFIG.get("pass","")
                    token=api_login(u,p)
                    if not token: self.error_signal.emit("登录失败"); return
                    if LoginState.logged_in: LoginState.api_token=token
                    else: CONFIG["token"]=token; save_config()
                if self._stop_flag: return
                sid=creds("api_sid") or CONFIG.get("sid","")
                result=self._poll(token,sid,self.phone)
                if self._stop_flag: return
                if result:
                    self.progress_signal.emit(result.get("yzm",""))
                    self.result_signal.emit({"phone":self.phone,"yzm":result.get("yzm",""),"sms":result.get("sms","")})
                else: self.error_signal.emit("超时：未收到验证码")
            except Exception as e: self.error_signal.emit(str(e))
            finally: self.finished_signal.emit()
        def _poll(self,token,sid,phone):
            iv,max_w=CONFIG["poll_interval"],CONFIG["max_wait"]; start=time.time(); n=0
            while time.time()-start<max_w:
                if self._stop_flag: return None
                n+=1
                r=api_get_message(token,sid,phone)
                if r: return r
                for _ in range(iv):
                    if self._stop_flag: return None
                    time.sleep(1)
            return None

    # ==================== 全自动流程页面 ====================
    class AutoFlowPage(QWidget):
        def __init__(self,parent=None):
            super().__init__(parent); self.setObjectName("autoPage"); self.worker=None; self._setup_ui()
        def _setup_ui(self):
            l=QVBoxLayout(); l.setAlignment(Qt.AlignCenter); l.setSpacing(20)
            _make_scrollable(self,l)
            self.btn=PushButton("▶ 一键获取",self); self.btn.setCursor(Qt.PointingHandCursor)
            self.btn.setMinimumSize(120,120); self.btn.setMaximumSize(280,280)
            self.btn.setStyleSheet("PushButton{border-radius:60px;font-size:20px;font-weight:bold;border:none;color:white;background:qlineargradient(x1:0,y1:0,x2:1,y2:1,stop:0 #2196F3,stop:1 #0D47A1)}PushButton:hover{background:qlineargradient(x1:0,y1:0,x2:1,y2:1,stop:0 #42A5F5,stop:1 #1565C0)}PushButton:pressed{padding:2px 0 0 2px;background:qlineargradient(x1:0,y1:0,x2:1,y2:1,stop:0 #1565C0,stop:1 #0D47A1)}PushButton:disabled{background:qlineargradient(x1:0,y1:0,x2:1,y2:1,stop:0 #90CAF9,stop:1 #64B5F6);color:#E3F2FD}")
            self.btn.clicked.connect(self._start)
            bc=QVBoxLayout(); bc.setAlignment(Qt.AlignCenter); bc.addWidget(self.btn); l.addLayout(bc,3)
            self.status=StrongBodyLabel("",self); self.status.setAlignment(Qt.AlignCenter); l.addWidget(self.status)
            cl=QHBoxLayout(); cl.setSpacing(30); cl.setAlignment(Qt.AlignCenter)
            self.phone_card=CopyableCard("手机号","",self); cl.addWidget(self.phone_card)
            self.code_card=CopyableCard("验证码","",self); cl.addWidget(self.code_card)
            l.addLayout(cl,1)
            bl=QVBoxLayout(); bl.setAlignment(Qt.AlignCenter); bl.setSpacing(8)
            self.spinner=IndeterminateProgressRing(self); self.spinner.setFixedSize(40,40); self.spinner.setVisible(False); bl.addWidget(self.spinner)
            self.succ=StrongBodyLabel("",self); self.succ.setAlignment(Qt.AlignCenter); self.succ.setStyleSheet("font-size:16px;color:#4CAF50"); self.succ.setVisible(False); bl.addWidget(self.succ)
            l.addLayout(bl,1)
        def _start(self):
            if self.worker and self.worker.isRunning():
                self.worker.stop(); self.btn.setEnabled(False); self.btn.setText("正在停止..."); return
            self.phone_card.set_value(""); self.code_card.set_value(""); self.succ.setVisible(False); self.status.setText("")
            self.spinner.setVisible(True)
            self.btn.setText("■ 停止"); self.btn.setStyleSheet("PushButton{border-radius:60px;font-size:20px;font-weight:bold;border:none;color:white;background:qlineargradient(x1:0,y1:0,x2:1,y2:1,stop:0 #e53935,stop:1 #b71c1c)}PushButton:hover{background:qlineargradient(x1:0,y1:0,x2:1,y2:1,stop:0 #ef5350,stop:1 #c62828)}PushButton:pressed{padding:2px 0 0 2px;background:qlineargradient(x1:0,y1:0,x2:1,y2:1,stop:0 #c62828,stop:1 #b71c1c)}")
            self.worker=AutoFlowWorker()
            self.worker.log_signal.connect(lambda m:None)
            self.worker.result_signal.connect(self._on_result); self.worker.error_signal.connect(self._on_error)
            self.worker.progress_signal.connect(self._on_progress); self.worker.finished_signal.connect(self._on_fin)
            self.worker.start()
        def _restore_btn(self):
            self.btn.setText("▶ 一键获取"); self.btn.setStyleSheet("PushButton{border-radius:60px;font-size:20px;font-weight:bold;border:none;color:white;background:qlineargradient(x1:0,y1:0,x2:1,y2:1,stop:0 #2196F3,stop:1 #0D47A1)}PushButton:hover{background:qlineargradient(x1:0,y1:0,x2:1,y2:1,stop:0 #42A5F5,stop:1 #1565C0)}PushButton:pressed{padding:2px 0 0 2px;background:qlineargradient(x1:0,y1:0,x2:1,y2:1,stop:0 #1565C0,stop:1 #0D47A1)}")
        def _on_fin(self):
            self.btn.setEnabled(True); self._restore_btn()
            if self.worker and self.worker._stop_flag:
                self.spinner.setVisible(False); self.succ.setVisible(False); self.status.setText("已停止")
        def _on_result(self,d):
            self.phone_card.set_value(d.get("phone","")); self.code_card.set_value(d.get("yzm",""))
            self.spinner.setVisible(False); self.succ.setText("✓ 已获取成功"); self.succ.setVisible(True); self.status.setText("")
        def _on_error(self,m): self.spinner.setVisible(False); self.succ.setVisible(False); self.btn.setEnabled(True); self._restore_btn(); self.status.setText(f"失败: {m}"); InfoBar.error("",m,duration=3000,parent=self)
        def _on_progress(self,phone,status):
            if phone: self.phone_card.set_value(phone)
            if status=="等待验证码...": self.status.setText("号码已获取，等待验证码...")
            elif status and status!="等待验证码...": self.code_card.set_value(status)

    # ==================== 查询验证码页面 ====================
    class QueryPage(QWidget):
        def __init__(self,parent=None):
            super().__init__(parent); self.setObjectName("queryPage"); self.worker=None; self._setup_ui()
        def _setup_ui(self):
            l=QVBoxLayout(); l.setSpacing(24); l.addWidget(TitleLabel("查询验证码")); l.addSpacing(8)
            _make_scrollable(self,l)
            sc=SimpleCardWidget(); sc.setBorderRadius(12); sl=QHBoxLayout(sc); sl.setContentsMargins(16,12,16,12)
            il=QLabel(); il.setFixedSize(24,24); il.setText("🔍"); il.setStyleSheet("font-size:18px"); sl.addWidget(il)
            self.phone_input=LineEdit(); self.phone_input.setPlaceholderText("输入手机号..."); self.phone_input.setClearButtonEnabled(True); sl.addWidget(self.phone_input,1)
            self.query_btn=PushButton(); self.query_btn.setFixedWidth(120); self.query_btn.setText("查询")
            self.query_btn.setStyleSheet("PushButton{border-radius:8px;color:white;font-size:16px;font-weight:bold;border:none;padding:8px 20px;background:qlineargradient(x1:0,y1:0,x2:1,y2:1,stop:0 #2196F3,stop:1 #0D47A1)}PushButton:hover{background:qlineargradient(x1:0,y1:0,x2:1,y2:1,stop:0 #42A5F5,stop:1 #1565C0)}PushButton:pressed{background:qlineargradient(x1:0,y1:0,x2:1,y2:1,stop:0 #1565C0,stop:1 #0D47A1)}")
            self.query_btn.clicked.connect(self._start); sl.addWidget(self.query_btn); l.addWidget(sc)
            rc=SimpleCardWidget(); rc.setBorderRadius(12); rl=QVBoxLayout(rc); rl.setContentsMargins(24,20,24,20); rl.setSpacing(12)
            rl.addWidget(StrongBodyLabel("验证码结果",rc))
            cr=QHBoxLayout(); self.code_card=CopyableCard("验证码","",rc); cr.addWidget(self.code_card,1)
            self.spinner=IndeterminateProgressRing(rc); self.spinner.setFixedSize(32,32); self.spinner.setVisible(False); cr.addWidget(self.spinner)
            rl.addLayout(cr); l.addWidget(rc); l.addStretch()
        def _start(self):
            if self.worker and self.worker.isRunning():
                self.worker.stop(); self.query_btn.setEnabled(False); self.query_btn.setText("正在停止..."); return
            phone=self.phone_input.text().strip()
            if not phone: InfoBar.warning("","请输入手机号",duration=2000,parent=self); return
            self.code_card.set_value(""); self.spinner.setVisible(True)
            self.query_btn.setText("■ 停止"); self.query_btn.setStyleSheet("PushButton{border-radius:8px;color:white;font-size:16px;font-weight:bold;border:none;padding:8px 20px;background:qlineargradient(x1:0,y1:0,x2:1,y2:1,stop:0 #e53935,stop:1 #b71c1c)}PushButton:hover{background:qlineargradient(x1:0,y1:0,x2:1,y2:1,stop:0 #ef5350,stop:1 #c62828)}PushButton:pressed{background:qlineargradient(x1:0,y1:0,x2:1,y2:1,stop:0 #c62828,stop:1 #b71c1c)}")
            self.worker=QueryCodeWorker(phone)
            self.worker.log_signal.connect(lambda m:None)
            self.worker.result_signal.connect(self._on_result); self.worker.error_signal.connect(self._on_error)
            self.worker.progress_signal.connect(self._on_progress); self.worker.finished_signal.connect(self._on_fin)
            self.worker.start()
        def _on_fin(self):
            self.query_btn.setEnabled(True); self.query_btn.setText("查询")
            self.query_btn.setStyleSheet("PushButton{border-radius:8px;color:white;font-size:16px;font-weight:bold;border:none;padding:8px 20px;background:qlineargradient(x1:0,y1:0,x2:1,y2:1,stop:0 #2196F3,stop:1 #0D47A1)}PushButton:hover{background:qlineargradient(x1:0,y1:0,x2:1,y2:1,stop:0 #42A5F5,stop:1 #1565C0)}PushButton:pressed{background:qlineargradient(x1:0,y1:0,x2:1,y2:1,stop:0 #1565C0,stop:1 #0D47A1)}")
            if self.worker and self.worker._stop_flag: self.spinner.setVisible(False)
        def _on_result(self,d): self.code_card.set_value(d.get("yzm","")); self.spinner.setVisible(False); InfoBar.success("",f"验证码: {d.get('yzm','')}",duration=3000,parent=self)
        def _on_error(self,m): self.spinner.setVisible(False); InfoBar.error("",m,duration=3000,parent=self)
        def _on_progress(self,y): self.code_card.set_value(y)

    # ==================== 个人中心页面 ====================
    class ProfilePage(QWidget):
        def __init__(self,parent=None):
            super().__init__(parent); self.setObjectName("profilePage"); self._setup_ui()
        def _setup_ui(self):
            self.stack=QStackedWidget(self); l=QVBoxLayout(self); l.addWidget(self.stack)
            self.prompt=LoginPromptWidget("个人中心","登录后可查看个人信息",self); self.prompt.login_btn.clicked.connect(self._do_login)
            # 已登录内容
            iw=QWidget(); il=QVBoxLayout(iw); il.setSpacing(20); il.addWidget(TitleLabel("个人中心")); il.addSpacing(8)
            card=SimpleCardWidget(); card.setBorderRadius(12); cl=QVBoxLayout(card); cl.setContentsMargins(24,20,24,20); cl.setSpacing(12)
            cl.addWidget(SubtitleLabel("登录信息",card))
            self.info_labels={}
            for lb,k in [("用户名","user"),("应用账号","app_user"),("应用密码","app_pass")]:
                r=QHBoxLayout(); r.addWidget(StrongBodyLabel(lb+":",card))
                self.info_labels[k]=StrongBodyLabel("",card); self.info_labels[k].setStyleSheet("color:#666")
                r.addWidget(self.info_labels[k]); r.addStretch(); cl.addLayout(r)
            il.addWidget(card)
            lcard=SimpleCardWidget(); lcard.setBorderRadius(12); ll=QVBoxLayout(lcard); ll.setContentsMargins(24,20,24,20); ll.setSpacing(12)
            ll.addWidget(SubtitleLabel("快捷链接",lcard))
            tg=PushButton("🌐 豪猪 Telegram 频道",lcard); tg.setStyleSheet("PushButton{background:#0088cc;color:white;border-radius:8px;padding:10px 24px;font-size:15px;font-weight:bold;border:none}PushButton:hover{background:#0099dd}PushButton:pressed{background:#0077bb}")
            tg.setCursor(Qt.PointingHandCursor); tg.clicked.connect(lambda:__import__('webbrowser').open("https://telegramchannels.me/zh/groups/haozhuma")); ll.addWidget(tg)
            il.addWidget(lcard)
            acard=SimpleCardWidget(); acard.setBorderRadius(12); al=QVBoxLayout(acard); al.setContentsMargins(24,20,24,20); al.setSpacing(8)
            al.addWidget(SubtitleLabel("关于",acard))
            al.addWidget(StrongBodyLabel("豪猪信客 — 归属 @广州聚星",acard))
            al.addWidget(StrongBodyLabel("联系方式: 2787326121@qq.com",acard))
            il.addWidget(acard)
            # 滚动
            sa=QScrollArea(); sa.setWidget(iw); sa.setWidgetResizable(True); sa.setFrameShape(0)
            self.stack.addWidget(self.prompt)
            self.stack.addWidget(sa)
            self.stack.setCurrentIndex(0)
        def showEvent(self,e): super().showEvent(e); self._refresh()
        def _refresh(self):
            if LoginState.logged_in:
                self.info_labels["user"].setText(LoginState.current_user)
                self.info_labels["app_user"].setText(LoginState.app_user or "未设置")
                self.info_labels["app_pass"].setText("••••••" if LoginState.app_pass else "未设置")
                self.stack.setCurrentIndex(1)
            else: self.stack.setCurrentIndex(0)
        def _do_login(self):
            d=LoginDialog(self)
            if d.exec()==QDialog.Accepted:
                r=d.result_data
                LoginState.login(r["user"],r["password"],r["data"].get("api_server",""),r["data"].get("api_user",""),r["data"].get("api_pass",""),r["data"].get("api_sid",""),r["data"].get("api_token",""),r["data"].get("app_user",""),r["data"].get("app_pass",""))
                if d.remember: save_session(r["user"], r["password"])
                else: clear_session()
                self._refresh(); InfoBar.success("",f"欢迎，{r['user']}！",duration=2000,parent=self)

    # ==================== 设置页面 ====================
    class SettingsPage(QWidget):
        config_saved=pyqtSignal()
        def __init__(self,parent=None):
            super().__init__(parent); self.setObjectName("settingsPage")
            self.stack=QStackedWidget(self); l=QVBoxLayout(self); l.addWidget(self.stack)
            self._build_prompt(); self._build_content(); self.stack.setCurrentIndex(0)
        def _build_prompt(self):
            w=LoginPromptWidget("设置","登录后可查看和修改 API 配置",self); w.login_btn.clicked.connect(self._do_login); self.stack.addWidget(w)
        def _build_content(self):
            w=QWidget(); l=QVBoxLayout(w); l.setSpacing(16); l.addWidget(TitleLabel("设置")); l.addSpacing(8)
            ac=SimpleCardWidget(); ac.setBorderRadius(12); al=QVBoxLayout(ac); al.setContentsMargins(24,20,24,20); al.setSpacing(12)
            al.addWidget(SubtitleLabel("API 配置",ac))
            self.entries={}
            for lb,k in [("服务器地址","api_server"),("API 账号","api_user"),("API 密码","api_pass"),("项目 ID","api_sid"),("Token","api_token")]:
                r=QHBoxLayout(); r.addWidget(StrongBodyLabel(lb,ac))
                e=LineEdit(); 
                if k=="api_pass": e.setEchoMode(LineEdit.Password)
                e.setFixedWidth(400 if k=="api_token" else 350); r.addWidget(e); r.addStretch(); al.addLayout(r); self.entries[k]=e
            l.addWidget(ac)
            apc=SimpleCardWidget(); apc.setBorderRadius(12); apl=QVBoxLayout(apc); apl.setContentsMargins(24,20,24,20); apl.setSpacing(12)
            apl.addWidget(SubtitleLabel("应用配置（个人中心）",apc))
            for lb,k in [("应用账号","app_user"),("应用密码","app_pass")]:
                r=QHBoxLayout(); r.addWidget(StrongBodyLabel(lb,apc)); e=LineEdit()
                if k=="app_pass": e.setEchoMode(LineEdit.Password)
                e.setFixedWidth(350); r.addWidget(e); r.addStretch(); apl.addLayout(r); self.entries[k]=e
            l.addWidget(apc)
            avc=SimpleCardWidget(); avc.setBorderRadius(12); avl=QVBoxLayout(avc); avl.setContentsMargins(24,20,24,20); avl.setSpacing(12)
            avl.addWidget(SubtitleLabel("高级选项",avc))
            self.skip_cb=CheckBox("不登录获取 token（复用已有 token）",avc); self.skip_cb.setChecked(CONFIG.get("skip_login",True)); avl.addWidget(self.skip_cb)
            pr=QHBoxLayout(); pr.addWidget(StrongBodyLabel("轮询间隔(秒)",avc)); self.poll_entry=LineEdit(); self.poll_entry.setText(str(CONFIG.get("poll_interval",15))); self.poll_entry.setFixedWidth(120); pr.addWidget(self.poll_entry); pr.addStretch(); avl.addLayout(pr)
            wr=QHBoxLayout(); wr.addWidget(StrongBodyLabel("最大等待(秒)",avc)); self.wait_entry=LineEdit(); self.wait_entry.setText(str(CONFIG.get("max_wait",180))); self.wait_entry.setFixedWidth(120); wr.addWidget(self.wait_entry); wr.addStretch(); avl.addLayout(wr)
            l.addWidget(avc)
            uc=SimpleCardWidget(); uc.setBorderRadius(12); ul=QVBoxLayout(uc); ul.setContentsMargins(24,20,24,20); ul.setSpacing(12)
            ul.addWidget(SubtitleLabel("自动更新配置",uc))
            ul.addWidget(StrongBodyLabel("更新源: GitHub Releases（私有仓库需 Token）",uc))
            self.update_entries={}
            for lb,k,p in [("GitHub 仓库","gh_repo","如: zenyth02012/haozhu-xinke"),("Token","gh_token","GitHub Personal Access Token"),("代理地址","proxy_url","可选, 如 http://127.0.0.1:7890")]:
                r=QHBoxLayout(); r.addWidget(StrongBodyLabel(lb,uc)); e=LineEdit()
                if k=="gh_token": e.setEchoMode(LineEdit.Password)
                e.setPlaceholderText(p); e.setFixedWidth(480 if k=="proxy_url" else 400); e.setText(CONFIG.get(k,""))
                r.addWidget(e); r.addStretch(); ul.addLayout(r); self.update_entries[k]=e
            l.addWidget(uc)
            bc=SimpleCardWidget(); bc.setBorderRadius(12); bl=QHBoxLayout(bc); bl.setContentsMargins(24,16,24,16)
            save_btn=PrimaryPushButton("保存配置",bc); save_btn.clicked.connect(self._save); bl.addWidget(save_btn)
            logout_btn=PushButton("退出登录",bc); logout_btn.clicked.connect(self._logout); bl.addWidget(logout_btn); bl.addStretch()
            l.addWidget(bc)
            # 滚动
            sa=QScrollArea(); sa.setWidget(w); sa.setWidgetResizable(True); sa.setFrameShape(0)
            self.stack.addWidget(sa)
        def showEvent(self,e):
            super().showEvent(e)
            if LoginState.logged_in:
                for k in ["api_server","api_user","api_pass","api_sid","api_token","app_user","app_pass"]:
                    self.entries[k].setText(getattr(LoginState,k,""))
                self.stack.setCurrentIndex(1)
            else: self.stack.setCurrentIndex(0)
        def _do_login(self):
            d=LoginDialog(self)
            if d.exec()==QDialog.Accepted:
                r=d.result_data
                LoginState.login(r["user"],r["password"],r["data"].get("api_server",""),r["data"].get("api_user",""),r["data"].get("api_pass",""),r["data"].get("api_sid",""),r["data"].get("api_token",""),r["data"].get("app_user",""),r["data"].get("app_pass",""))
                if d.remember: save_session(r["user"], r["password"])
                else: clear_session()
                for k in ["api_server","api_user","api_pass","api_sid","api_token","app_user","app_pass"]:
                    self.entries[k].setText(getattr(LoginState,k,""))
                self.stack.setCurrentIndex(1); InfoBar.success("",f"欢迎，{r['user']}！",duration=2000,parent=self)
        def _logout(self):
            LoginState.logout(); clear_session(); self.stack.setCurrentIndex(0); InfoBar.info("","已退出登录",duration=2000,parent=self)
        def _save(self):
            try:
                for k in ["api_server","api_user","api_pass","api_sid","api_token","app_user","app_pass"]:
                    setattr(LoginState,k,self.entries[k].text().strip())
                CONFIG["skip_login"]=self.skip_cb.isChecked()
                CONFIG["poll_interval"]=int(self.poll_entry.text().strip())
                CONFIG["max_wait"]=int(self.wait_entry.text().strip())
                rc=get_root_config()
                rc.update({"current_user":LoginState.current_user,"users":rc.get("users",[]),
                    "poll_interval":CONFIG["poll_interval"],"max_wait":CONFIG["max_wait"],"skip_login":CONFIG["skip_login"]})
                for k in ["gh_repo","gh_token","proxy_url"]:
                    rc[k]=self.update_entries[k].text().strip(); CONFIG[k]=rc[k]
                save_root_config(rc)
                if LoginState._login_password:
                    save_user_config(LoginState.current_user,LoginState._login_password,LoginState.to_config())
                self.config_saved.emit(); InfoBar.success("","配置已保存",duration=2000,parent=self)
            except ValueError: InfoBar.error("","轮询间隔和最大等待必须为数字",duration=3000,parent=self)

    # ==================== 自动更新 ====================
    APP_VERSION = "4.0"
    def _gh_api(url, token, proxy):
        headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}
        proxies = {"http": proxy, "https": proxy} if proxy else None
        try:
            r = requests.get(url, headers=headers, proxies=proxies, timeout=15, verify=False)
            if r.status_code == 200: return r.json()
        except: pass
        return None
    def check_update():
        repo = CONFIG.get("gh_repo","")
        token = CONFIG.get("gh_token","")
        proxy = CONFIG.get("proxy_url","")
        if not repo: return None
        api = f"https://api.github.com/repos/{repo}/releases/latest"
        data = _gh_api(api, token, proxy)
        if not data: return None
        tag = data.get("tag_name","").lstrip("v")
        if tag <= APP_VERSION: return None
        assets = data.get("assets", [])
        dl = ""
        for a in assets:
            if a.get("name","").endswith(".exe"):
                dl = a.get("browser_download_url",""); break
        if not dl: dl = data.get("zipball_url","")
        return tag, dl, data.get("body","")

    class UpdateDialog(QDialog):
        def __init__(self, new_ver, notes, download_url, parent=None):
            super().__init__(parent); self.dl=download_url
            self.setWindowTitle("发现新版本"); self.setFixedSize(440, 260)
            l=QVBoxLayout(self); l.setContentsMargins(24,20,24,20); l.setSpacing(12)
            l.addWidget(TitleLabel(f"新版本 v{new_ver} 可用",self))
            if notes: n=StrongBodyLabel(notes[:200],self); n.setWordWrap(True); l.addWidget(n)
            br=QHBoxLayout()
            skip_btn=PushButton("稍后提醒",self); skip_btn.clicked.connect(self.reject); br.addWidget(skip_btn)
            update_btn=PrimaryPushButton("立即更新",self); update_btn.clicked.connect(self._update); br.addWidget(update_btn)
            l.addLayout(br)
        def _update(self):
            import subprocess, shutil
            self.setEnabled(False)
            try:
                proxy = CONFIG.get("proxy_url","")
                proxies = {"http": proxy, "https": proxy} if proxy else None
                token = CONFIG.get("gh_token","")
                headers = {"Authorization": f"token {token}"} if token else {}
                InfoBar.info("","正在下载更新...",duration=3000,parent=self)
                r = requests.get(self.dl, headers=headers, proxies=proxies, timeout=120, verify=False, stream=True)
                dst = os.path.join(BASE_DIR, "豪猪信客_new.exe")
                with open(dst, "wb") as f:
                    for chunk in r.iter_content(8192): f.write(chunk)
                exe_name = "豪猪信客.exe"
                updater = f'''@echo off\r\ntimeout /t 1 /nobreak >nul\r\ndel "{os.path.join(BASE_DIR, exe_name)}"\r\nmove "{dst}" "{os.path.join(BASE_DIR, exe_name)}"\r\nstart "" "{os.path.join(BASE_DIR, exe_name)}"\r\nif exist "%~f0" del "%~f0"\r\n'''
                bat = os.path.join(BASE_DIR, "_update.bat")
                with open(bat, "w") as f: f.write(updater)
                subprocess.Popen([bat], shell=True, creationflags=subprocess.DETACHED_PROCESS)
                QApplication.quit()
            except Exception as e:
                InfoBar.error("",f"更新失败: {e}",duration=5000,parent=self); self.setEnabled(True)

    # ==================== 主窗口 ====================
    class MainWindow(FluentWindow):
        def __init__(self):
            super().__init__(); self.setWindowTitle("豪猪信客")
            p=os.path.join(os.path.dirname(os.path.abspath(__file__)),"AI智能管家功能与技术栈 (2).png")
            if os.path.exists(p): self.setWindowIcon(QIcon(p))
            self.setMinimumSize(200, 200); self.resize(820, 600)
            self.auto_page=AutoFlowPage(); self.query_page=QueryPage(); self.profile_page=ProfilePage(); self.settings_page=SettingsPage()
            self.addSubInterface(self.auto_page,FluentIcon.PLAY,"全自动流程")
            self.addSubInterface(self.query_page,FluentIcon.SEARCH,"查询验证码")
            self.addSubInterface(self.profile_page,FluentIcon.PEOPLE,"个人中心",position=NavigationItemPosition.BOTTOM)
            self.addSubInterface(self.settings_page,FluentIcon.SETTING,"设置",position=NavigationItemPosition.BOTTOM)
            self._start_local_server()
            # 启动后延迟检测更新
            from PyQt5.QtCore import QTimer
            QTimer.singleShot(2000, self._check_update)
        def _check_update(self):
            try:
                info = check_update()
                if info:
                    d = UpdateDialog(*info, self)
                    if d.exec() == QDialog.Accepted: pass
            except: pass
        def _start_local_server(self):
            try:
                self._server=QLocalServer(self); self._server.setSocketOptions(QLocalServer.WorldAccessOption)
                self._server.newConnection.connect(self._on_client)
                self._server.listen("HZX-智客3.0-LocalServer")
            except: pass
        def _on_client(self):
            try:
                c=self._server.nextPendingConnection()
                if c:
                    c.readyRead.connect(lambda: self._bring_to_front(c))
            except: pass
        def _bring_to_front(self, c):
            try:
                c.readAll()
                self.showNormal(); self.raise_(); self.activateWindow()
                c.disconnectFromServer(); c.deleteLater()
            except: pass

    # ==================== 单例检测 ====================
    SHARED_KEY = "HZX-智客3.0-SingleInstance"
    LOCAL_SERVER = "HZX-智客3.0-LocalServer"
    from PyQt5.QtCore import QSharedMemory
    shared_mem = QSharedMemory(SHARED_KEY)
    if shared_mem.attach():
        # 已有实例运行，通知它显示窗口
        sock = QLocalSocket()
        sock.connectToServer(LOCAL_SERVER)
        if sock.waitForConnected(1000):
            sock.write(b"show"); sock.waitForBytesWritten(1000); sock.disconnectFromServer()
        sys.exit(0)
    shared_mem.create(1)

    window=MainWindow()
    # 自动登录（5天免登录）
    user, password = load_session()
    if user and password:
        data = load_user_config(user, password)
        if data:
            LoginState.login(user, password, data.get("api_server",""), data.get("api_user",""),
                             data.get("api_pass",""), data.get("api_sid",""), data.get("api_token",""),
                             data.get("app_user",""), data.get("app_pass",""))
    p=os.path.join(os.path.dirname(os.path.abspath(__file__)),"AI智能管家功能与技术栈 (2).png")
    if os.path.exists(p): app.setWindowIcon(QIcon(p))
    window.show(); sys.exit(app.exec())

if __name__=="__main__": main()
