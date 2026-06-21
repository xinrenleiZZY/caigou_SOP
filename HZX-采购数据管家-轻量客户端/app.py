"""
HZX-采购数据管家 V1.0.0 轻量客户端
本地只保留 UI 框架，处理引擎通过 HTTP 调用 NAS/云端服务器
"""
import sys, os, json, threading, shutil, time, smtplib, email
from PyQt5.QtCore import Qt, QTimer, pyqtSignal, QObject, QUrl, QThread
from PyQt5.QtGui import QFont, QDesktopServices
from PyQt5.QtWidgets import QApplication, QFileDialog, QHBoxLayout, QVBoxLayout, QLabel, QWidget, QRadioButton, QButtonGroup

from qfluentwidgets import (
    FluentWindow, NavigationItemPosition, FluentIcon as FIC,
    SimpleCardWidget, PushButton, PrimaryPushButton, TransparentToolButton,
    LineEdit, TextEdit, InfoBar, InfoBarPosition, IndeterminateProgressRing,
    setTheme, Theme, ComboBox, BodyLabel, TitleLabel,
    CaptionLabel, StrongBodyLabel, SpinBox, ScrollArea,
    setThemeColor, SwitchButton, PasswordLineEdit,
)

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from config import APP_NAME, VERSION, DEVELOPER, PHONE, SOP_DOCS_DIR, FILE_LIST_PATH, load_settings, save_settings, ensure_nas_connected
from main_processing import run_processing

# ========== 信号基类 ==========
class LogSignal(QObject):
    log_received = pyqtSignal(str)

# ========== 处理工作线程（本地处理，文件来自 NAS）==========
class ProcessWorker(QThread):
    finished = pyqtSignal(bool, str)
    log_signal = pyqtSignal(str)

    def __init__(self, source, output, year1, year2):
        super().__init__()
        self.source = source; self.output = output
        self.year1 = year1; self.year2 = year2

    def run(self):
        def log(m):
            self.log_signal.emit(m)
        try:
            log(f"� 源文件: {self.source}")
            log(f"📥 输出到: {self.output}")
            ok = run_processing(self.source, self.output, self.year1, self.year2,
                                enable_border=False, log=log)
            self.finished.emit(ok, self.output)
        except Exception as e:
            import traceback
            self.log_signal.emit(f"❌ 异常: {e}\n{traceback.format_exc()}")
            self.finished.emit(False, self.output)

# ========== 邮件工作线程 ==========
class EmailWorker(QThread):
    done = pyqtSignal(bool, str)
    log_signal = pyqtSignal(str)

    def __init__(self, to_addr, subject, body):
        super().__init__()
        self.to_addr = to_addr; self.subject = subject; self.body = body
    def run(self):
        try:
            msg = email.message.EmailMessage()
            msg["From"] = "2787326121@qq.com"; msg["To"] = self.to_addr
            msg["Subject"] = self.subject or f"[{APP_NAME}] 用户反馈"
            msg.set_content(self.body)
            with smtplib.SMTP_SSL("smtp.qq.com", 465, timeout=30) as s:
                s.login("2787326121@qq.com", "nfjrsfsqiorkdega"); s.send_message(msg)
            self.done.emit(True, "邮件已发送")
        except Exception as e:
            import traceback
            self.log_signal.emit(f"[邮箱] 发送失败: {e}\n{traceback.format_exc()}")
            self.done.emit(False, str(e))


# ========== 主窗口 ==========
class MainWindow(FluentWindow):
    def __init__(self):
        super().__init__()
        self.settings = load_settings()
        is_dark = self.settings.get("theme", "light") == "dark"
        if is_dark:
            setTheme(Theme.DARK)
        else:
            setTheme(Theme.LIGHT)
        setThemeColor("#0078D4")
        from PyQt5.QtGui import QColor
        self.setCustomBackgroundColor(
            QColor(243, 243, 243),
            QColor(28, 28, 28)
        )
        self.setWindowTitle(f"{APP_NAME} {VERSION}")
        self.resize(1200, 800)
        os.makedirs(SOP_DOCS_DIR, exist_ok=True)

        # 自动连接 NAS
        if ensure_nas_connected(self.settings):
            print(f"[NAS] 已连接到: {self.settings.get('nas_path')}")
        else:
            print(f"[NAS] 无法连接，请检查网络和配置")

        self.file_list = self.load_file_list()
        self.current_source_file = ""
        self.current_output_file = ""
        self.year1_val = self.settings.get("year1", 25)
        self.year2_val = self.settings.get("year2", 26)

        self.home_page = HomePage(self)
        self.sop_page = SOPPage(self)
        self.file_list_page = FileListPage(self)
        self.log_page = LogPage(self)
        self.settings_page = SettingsPage(self)
        self.email_page = EmailPage(self)
        self.tutorial_page = TutorialPage(self)
        self.version_page = VersionPage(self)

        self.home_page.setObjectName("home")
        self.sop_page.setObjectName("sopProcess")
        self.file_list_page.setObjectName("fileList")
        self.log_page.setObjectName("processLog")
        self.settings_page.setObjectName("settings")
        self.email_page.setObjectName("emailFeedback")
        self.tutorial_page.setObjectName("tutorial")
        self.version_page.setObjectName("versionInfo")

        # 所有页面容器设置透明背景，适配深色模式
        transparent_style = "QWidget{background:transparent;} QScrollArea{background:transparent;}"
        for page in [self.home_page, self.sop_page, self.file_list_page, self.log_page,
                     self.settings_page, self.email_page, self.tutorial_page, self.version_page]:
            page.setStyleSheet(transparent_style)
            if hasattr(page, 'viewport'):
                page.viewport().setStyleSheet("background:transparent;")

        ni = self.navigationInterface
        self.addSubInterface(self.home_page, FIC.HOME, "首页", isTransparent=True)

        from PyQt5.QtWidgets import QFrame, QVBoxLayout
        self._tools_group = QFrame(); self._tools_group.setObjectName("toolsGroup"); self._tools_group.setStyleSheet("background:transparent;")
        gl = QVBoxLayout(self._tools_group)
        gl.addWidget(TitleLabel("工具管理")); gl.addWidget(BodyLabel("选择下方功能"))
        self.addSubInterface(self._tools_group, FIC.DEVELOPER_TOOLS, "工具管理", NavigationItemPosition.SCROLL, isTransparent=True)
        self.addSubInterface(self.sop_page, FIC.ALBUM, "月度数据对比SOP",
                             NavigationItemPosition.SCROLL, parent=self._tools_group, isTransparent=True)

        self.addSubInterface(self.file_list_page, FIC.DOCUMENT, "文件列表", NavigationItemPosition.SCROLL, isTransparent=True)
        self.addSubInterface(self.log_page, FIC.LIBRARY, "处理日志", NavigationItemPosition.SCROLL, isTransparent=True)
        self.addSubInterface(self.settings_page, FIC.SETTING, "设置", NavigationItemPosition.SCROLL, isTransparent=True)
        self.addSubInterface(self.email_page, FIC.MAIL, "邮箱反馈", NavigationItemPosition.SCROLL, isTransparent=True)

        self._help_group = QFrame(); self._help_group.setObjectName("helpGroup"); self._help_group.setStyleSheet("background:transparent;")
        gl2 = QVBoxLayout(self._help_group)
        gl2.addWidget(TitleLabel("帮助")); gl2.addWidget(BodyLabel("选择下方功能"))
        self.addSubInterface(self._help_group, FIC.HELP, "帮助", NavigationItemPosition.SCROLL, isTransparent=True)
        self.addSubInterface(self.tutorial_page, FIC.EDUCATION, "使用教程",
                             NavigationItemPosition.SCROLL, parent=self._help_group, isTransparent=True)
        self.addSubInterface(self.version_page, FIC.INFO, "版本信息",
                             NavigationItemPosition.SCROLL, parent=self._help_group, isTransparent=True)

        ni.addSeparator()
        ver_btn = PushButton(f"{APP_NAME}  {VERSION}")
        ver_btn.setEnabled(False); ver_btn.setFixedHeight(36)
        ni.addWidget("version", ver_btn, lambda: None, NavigationItemPosition.BOTTOM)

    def load_file_list(self):
        if os.path.exists(FILE_LIST_PATH):
            try: return json.load(open(FILE_LIST_PATH, 'r', encoding='utf-8'))
            except: pass
        return []
    def save_file_list(self):
        json.dump(self.file_list, open(FILE_LIST_PATH, 'w', encoding='utf-8'), ensure_ascii=False, indent=2)
    def add_to_file_list(self, path):
        if path and path not in self.file_list:
            self.file_list.append(path); self.save_file_list()


# ========== 首页 ==========
class HomePage(ScrollArea):
    def __init__(self, parent: MainWindow):
        super().__init__(); self.main = parent; self.setObjectName("homePage")
        w = QWidget(); w.setStyleSheet("background:transparent;"); l = QVBoxLayout(w); l.setSpacing(16); l.setContentsMargins(32, 24, 32, 24)
        l.addWidget(TitleLabel(f"欢迎使用 {APP_NAME}"))
        l.addWidget(CaptionLabel(f"版本 {VERSION} | 开发人: {DEVELOPER} | 电话: {PHONE}"))
        l.addSpacing(8)
        r1 = QHBoxLayout(); r1.setSpacing(16)
        for t, v, _ in [("📂 已上传文件", str(len([f for f in self.main.file_list if os.path.exists(f)])) + " 个", None),
                        ("📊 输出文件", str(len([f for f in os.listdir(SOP_DOCS_DIR) if f.endswith('.xlsx')])) + " 个" if os.path.exists(SOP_DOCS_DIR) else "0 个", None),
                        ("⚙️ 任务配置", "云端处理引擎", None)]:
            c = SimpleCardWidget(); c.setBorderRadius(12)
            cl = QVBoxLayout(c); cl.setContentsMargins(20, 16, 20, 16)
            cl.addWidget(StrongBodyLabel(t)); cl.addSpacing(4); cl.addWidget(TitleLabel(v)); r1.addWidget(c)
        l.addLayout(r1)
        l.addWidget(StrongBodyLabel("快速跳转"))
        r2 = QHBoxLayout(); r2.setSpacing(12)
        route_map = {"sopProcess": "sop_page", "fileList": "file_list_page",
                     "processLog": "log_page", "emailFeedback": "email_page"}
        for title, desc, route, _ in [("🚀 月度数据对比SOP","打开处理工具","sopProcess",None),
                                       ("📄 文件列表","查看上传和输出文件","fileList",None),
                                       ("📋 处理日志","查看处理记录","processLog",None),
                                       ("✉ 邮箱反馈","反馈问题","emailFeedback",None)]:
            c = SimpleCardWidget(); c.setBorderRadius(12); c.setCursor(Qt.PointingHandCursor)
            cl = QVBoxLayout(c); cl.setContentsMargins(20, 16, 20, 16)
            cl.addWidget(StrongBodyLabel(title)); cl.addWidget(CaptionLabel(desc))
            hr = route_map.get(route, route + "_page")
            c.mousePressEvent = lambda e, r=hr: self.main.switchTo(getattr(self.main, r))
            r2.addWidget(c)
        l.addLayout(r2); l.addStretch()
        self.setWidget(w); self.setWidgetResizable(True)


# ========== SOP 处理页面（轻量版，HTTP 调用）===========
class SOPPage(ScrollArea):
    def __init__(self, parent: MainWindow):
        super().__init__(); self.main = parent; self.setObjectName("sopPage")
        self.processing = False; self._worker = None; self._setup_ui()
    def _add_log(self, msg):
        if hasattr(self.main, 'log_page') and self.main.log_page:
            self.main.log_page.add_log(msg)
    def _setup_ui(self):
        w = QWidget(); w.setStyleSheet("background:transparent;"); l = QVBoxLayout(w); l.setSpacing(12); l.setContentsMargins(32, 20, 32, 20)
        l.addWidget(TitleLabel("月度采购数据对比SOP"))
        l.addWidget(CaptionLabel("上传源文件至 NAS 服务器，一键完成全部处理"))
        l.addSpacing(4)
        br = QHBoxLayout(); br.setAlignment(Qt.AlignCenter)
        self.btn_process = PushButton("▶  一键处理")
        self.btn_process.setCursor(Qt.PointingHandCursor)
        self.btn_process.setMinimumSize(180, 180); self.btn_process.setMaximumSize(180, 180)
        self.btn_process.setStyleSheet("""PushButton{border-radius:90px;font-size:20px;font-weight:bold;border:none;color:white;background:qlineargradient(x1:0,y1:0,x2:1,y2:1,stop:0 #0078D4,stop:1 #005A9E);}PushButton:hover{background:qlineargradient(x1:0,y1:0,x2:1,y2:1,stop:0 #1E88E5,stop:1 #006DB3);}PushButton:pressed{padding:2px 0 0 2px;background:qlineargradient(x1:0,y1:0,x2:1,y2:1,stop:0 #005A9E,stop:1 #00457E);}""")
        self.btn_process.clicked.connect(self.start_processing)
        br.addWidget(self.btn_process); l.addLayout(br)
        sr = QHBoxLayout(); sr.setAlignment(Qt.AlignCenter)
        self.progress_ring = IndeterminateProgressRing(); self.progress_ring.setFixedSize(36, 36); self.progress_ring.setVisible(False)
        sr.addWidget(self.progress_ring)
        self.status_label = BodyLabel(""); self.status_label.setAlignment(Qt.AlignCenter)
        sr.addWidget(self.status_label); l.addLayout(sr)

        card1 = SimpleCardWidget(); card1.setBorderRadius(12)
        c1 = QVBoxLayout(card1); c1.setContentsMargins(24, 16, 24, 16)
        c1.addWidget(StrongBodyLabel("源文件"))
        r1 = QHBoxLayout()
        self.file_path_edit = LineEdit(); self.file_path_edit.setPlaceholderText("选择Excel文件..."); self.file_path_edit.setReadOnly(True)
        r1.addWidget(self.file_path_edit, 1)
        self.btn_browse = PushButton(FIC.FOLDER, "浏览文件"); self.btn_browse.clicked.connect(self.browse_file); r1.addWidget(self.btn_browse)
        self.btn_nas = PushButton(FIC.SYNC, "浏览NAS"); self.btn_nas.clicked.connect(self.browse_nas); r1.addWidget(self.btn_nas)
        self.file_combo = ComboBox(); self.file_combo.setMinimumWidth(200); self.file_combo.setPlaceholderText("最近使用...")
        self.file_combo.currentTextChanged.connect(self._oc); r1.addWidget(self.file_combo)
        c1.addLayout(r1); l.addWidget(card1)

        card_t = SimpleCardWidget(); card_t.setBorderRadius(12)
        ct = QVBoxLayout(card_t); ct.setContentsMargins(24, 16, 24, 16)
        ct.addWidget(StrongBodyLabel("目标文件"))
        mr = QHBoxLayout()
        self.rb_default = QRadioButton("默认"); self.rb_custom = QRadioButton("自定义"); self.rb_default.setChecked(True)
        self.mode_group = QButtonGroup(); self.mode_group.addButton(self.rb_default, 0); self.mode_group.addButton(self.rb_custom, 1)
        self.mode_group.buttonClicked.connect(lambda b: self.output_path_edit.setEnabled(self.mode_group.id(b) == 1))
        mr.addWidget(self.rb_default); mr.addWidget(self.rb_custom); mr.addStretch()
        ct.addLayout(mr)
        self.output_path_edit = LineEdit(); self.output_path_edit.setPlaceholderText("自定义保存路径..."); self.output_path_edit.setEnabled(False)
        ct.addWidget(self.output_path_edit)
        self.btn_download = PrimaryPushButton("📥 下载结果文件"); self.btn_download.clicked.connect(self.download_result); self.btn_download.setEnabled(False)
        ct.addWidget(self.btn_download); l.addWidget(card_t)

        card2 = SimpleCardWidget(); card2.setBorderRadius(12)
        c2 = QVBoxLayout(card2); c2.setContentsMargins(24, 16, 24, 16)
        c2.addWidget(StrongBodyLabel("处理参数"))
        pr = QHBoxLayout()
        pr.addWidget(BodyLabel("年份1:"))
        self.year1_spin = SpinBox(); self.year1_spin.setValue(25); self.year1_spin.setRange(20, 99); pr.addWidget(self.year1_spin)
        pr.addSpacing(30)
        pr.addWidget(BodyLabel("年份2:"))
        self.year2_spin = SpinBox(); self.year2_spin.setValue(26); self.year2_spin.setRange(20, 99); pr.addWidget(self.year2_spin); pr.addStretch()
        c2.addLayout(pr); l.addWidget(card2); l.addStretch()
        self.setWidget(w); self.setWidgetResizable(True); self._refresh()

    def _refresh(self):
        self.file_combo.clear(); self.file_combo.addItem("")
        for fp in self.main.file_list:
            if os.path.exists(fp): self.file_combo.addItem(fp)
    def _oc(self, text):
        if text and os.path.exists(text): self.file_path_edit.setText(text); self.main.current_source_file = text
    def browse_file(self):
        p, _ = QFileDialog.getOpenFileName(self, "选择Excel文件", "", "Excel文件 (*.xlsx *.xlsm);;所有文件 (*.*)")
        if p: self.file_path_edit.setText(p); self.main.current_source_file = p; self.main.add_to_file_list(p); self._refresh()
    def browse_nas(self):
        """从 NAS 共享目录选择文件"""
        nas_path = self.main.settings.get("nas_path", r"\\192.168.40.3\钟正洋")
        if not os.path.exists(nas_path):
            ensure_nas_connected(self.main.settings)
        p, _ = QFileDialog.getOpenFileName(self, "从NAS选择Excel文件", nas_path, "Excel文件 (*.xlsx *.xlsm);;所有文件 (*.*)")
        if p: self.file_path_edit.setText(p); self.main.current_source_file = p; self.main.add_to_file_list(p); self._refresh()
    def _get_output(self, source):
        base = os.path.basename(source); name = os.path.splitext(base)[0]
        ts = time.strftime("_%Y%m%d_%H%M%S")
        default = os.path.join(SOP_DOCS_DIR, f"{name}{ts}-整合输出.xlsx")
        if self.rb_custom.isChecked() and self.output_path_edit.text(): return self.output_path_edit.text()
        return default
    def start_processing(self):
        if self.processing: return
        source = self.file_path_edit.text()
        if not source or not os.path.exists(source):
            InfoBar.error("错误", "请选择有效的Excel文件", orient=Qt.Horizontal, isClosable=True, position=InfoBarPosition.TOP, parent=self); return
        output = self._get_output(source)
        if self.rb_custom.isChecked() and not self.output_path_edit.text():
            InfoBar.warning("提示", "请填写自定义路径", orient=Qt.Horizontal, isClosable=True, position=InfoBarPosition.TOP, parent=self); return
        self.main.current_source_file = source; self.main.current_output_file = output
        self.processing = True
        self.btn_process.setEnabled(False); self.btn_process.setText("⏳  处理中...")
        self.btn_process.setStyleSheet("PushButton{border-radius:90px;font-size:20px;font-weight:bold;border:none;color:white;background:qlineargradient(x1:0,y1:0,x2:1,y2:1,stop:0 #FF9800,stop:1 #E65100);}")
        self.progress_ring.setVisible(True); self.status_label.setText("正在处理...")
        self.main.log_page.log_text.setText("等待处理...")
        y1 = self.year1_spin.value(); y2 = self.year2_spin.value()
        self.main.year1_val = y1; self.main.year2_val = y2
        self._worker = ProcessWorker(source, output, y1, y2)
        self._worker.log_signal.connect(self._add_log)
        self._worker.finished.connect(self._done)
        self._worker.start()
    def _restore_btn(self):
        self.btn_process.setEnabled(True); self.btn_process.setText("▶  一键处理")
        self.btn_process.setStyleSheet("PushButton{border-radius:90px;font-size:20px;font-weight:bold;border:none;color:white;background:qlineargradient(x1:0,y1:0,x2:1,y2:1,stop:0 #0078D4,stop:1 #005A9E);}PushButton:hover{background:qlineargradient(x1:0,y1:0,x2:1,y2:1,stop:0 #1E88E5,stop:1 #006DB3);}PushButton:pressed{padding:2px 0 0 2px;background:qlineargradient(x1:0,y1:0,x2:1,y2:1,stop:0 #005A9E,stop:1 #00457E);}")
        self.progress_ring.setVisible(False)
    def _done(self, ok, output):
        self.processing = False; self._restore_btn()
        if ok:
            self.btn_download.setEnabled(True); self.status_label.setText("✅ 处理完成！")
            self.status_label.setStyleSheet("color:#4CAF50;font-size:16px;font-weight:bold;")
            self._add_log("✅ 处理完成！"); InfoBar.success("完成", "结果文件已生成", orient=Qt.Horizontal, isClosable=True, position=InfoBarPosition.TOP, parent=self)
        else:
            self.status_label.setText("❌ 处理失败"); self.status_label.setStyleSheet("color:#E53935;font-size:16px;")
            self._add_log("❌ 处理失败，请检查日志"); InfoBar.error("失败", "请检查日志", orient=Qt.Horizontal, isClosable=True, position=InfoBarPosition.TOP, parent=self)
    def download_result(self):
        if not self.main.current_output_file or not os.path.exists(self.main.current_output_file):
            InfoBar.warning("提示", "请先完成处理", orient=Qt.Horizontal, isClosable=True, position=InfoBarPosition.TOP, parent=self); return
        p, _ = QFileDialog.getSaveFileName(self, "保存结果文件", os.path.basename(self.main.current_output_file), "Excel文件 (*.xlsx)")
        if p: shutil.copy2(self.main.current_output_file, p); InfoBar.success("成功", f"已保存到: {p}", orient=Qt.Horizontal, isClosable=True, position=InfoBarPosition.TOP, parent=self)


# ========== 处理日志 ==========
class LogPage(ScrollArea):
    def __init__(self, parent: MainWindow):
        super().__init__(); self.main = parent; self.setObjectName("logPage")
        w = QWidget(); w.setStyleSheet("background:transparent;"); l = QVBoxLayout(w); l.setSpacing(12); l.setContentsMargins(32, 24, 32, 24)
        l.addWidget(TitleLabel("处理日志")); l.addSpacing(4)
        c = SimpleCardWidget(); c.setBorderRadius(12)
        cl = QVBoxLayout(c); cl.setContentsMargins(20, 14, 20, 14)
        self.log_text = BodyLabel("暂无处理记录"); self.log_text.setWordWrap(True); self.log_text.setTextInteractionFlags(Qt.TextSelectableByMouse)
        cl.addWidget(self.log_text); l.addWidget(c); l.addStretch()
        self.setWidget(w); self.setWidgetResizable(True)
    def add_log(self, msg):
        t = self.log_text.text()
        if t == "暂无处理记录": t = ""
        lines = t.split('\n'); lines.append(f"[{time.strftime('%H:%M:%S')}] {msg}")
        if len(lines) > 200: lines = lines[-200:]
        self.log_text.setText('\n'.join(lines))
        self.verticalScrollBar().setValue(self.verticalScrollBar().maximum())


# ========== 文件列表 ==========
class FileListPage(ScrollArea):
    def __init__(self, parent: MainWindow):
        super().__init__(); self.main = parent; self.setObjectName("fileListPage")
        w = QWidget(); w.setStyleSheet("background:transparent;"); l = QVBoxLayout(w); l.setSpacing(12); l.setContentsMargins(32, 24, 32, 24)
        l.addWidget(TitleLabel("文件列表")); l.addSpacing(4)
        c1 = SimpleCardWidget(); c1.setBorderRadius(12)
        c1s = QVBoxLayout(c1); c1s.setContentsMargins(24, 16, 24, 16)
        c1s.addWidget(StrongBodyLabel("已上传的文件"))
        self.files_label = BodyLabel("暂无文件记录"); self.files_label.setStyleSheet("color:#888;"); c1s.addWidget(self.files_label)
        l.addWidget(c1)
        br = QHBoxLayout()
        self._br = PushButton(FIC.UPDATE, "刷新列表"); self._br.clicked.connect(self.refresh_list); br.addWidget(self._br)
        self._bc = PushButton(FIC.DELETE, "清空列表"); self._bc.clicked.connect(self.clear_list); br.addWidget(self._bc)
        self._bo = PrimaryPushButton(FIC.FOLDER, "📂 打开输出目录")
        self._bo.clicked.connect(lambda: self._opendir()); br.addWidget(self._bo)
        br.addStretch(); l.addLayout(br)
        l.addWidget(StrongBodyLabel("输出文件目录"))
        c2 = SimpleCardWidget(); c2.setBorderRadius(12)
        c2s = QVBoxLayout(c2); c2s.setContentsMargins(24, 16, 24, 16)
        self.output_label = BodyLabel("暂无输出文件"); self.output_label.setStyleSheet("color:#888;"); c2s.addWidget(self.output_label)
        l.addWidget(c2); l.addStretch()
        self.setWidget(w); self.setWidgetResizable(True); self.refresh_list()
    def refresh_list(self):
        files = [f for f in self.main.file_list if os.path.exists(f)]
        self.files_label.setText('\n'.join([f"  📄 {i+1}. {f}" for i,f in enumerate(files)]) if files else "暂无文件记录")
        self.files_label.setStyleSheet("" if files else "color:#888;")
        if os.path.exists(SOP_DOCS_DIR):
            out = [f for f in os.listdir(SOP_DOCS_DIR) if f.endswith('.xlsx')]
            self.output_label.setText('\n'.join([f"  📊 {f}" for f in out]) if out else "暂无输出文件")
            self.output_label.setStyleSheet("" if out else "color:#888;")
    def _opendir(self):
        if os.path.exists(SOP_DOCS_DIR):
            QDesktopServices.openUrl(QUrl.fromLocalFile(SOP_DOCS_DIR))
            InfoBar.success("已打开", f"目录: {SOP_DOCS_DIR}", orient=Qt.Horizontal, isClosable=True, position=InfoBarPosition.TOP, parent=self)
        else:
            InfoBar.warning("提示", "输出目录不存在", orient=Qt.Horizontal, isClosable=True, position=InfoBarPosition.TOP, parent=self)
    def clear_list(self):
        self.main.file_list = []; self.main.save_file_list(); self.refresh_list()
        InfoBar.info("已清空", "文件列表已清空", orient=Qt.Horizontal, isClosable=True, position=InfoBarPosition.TOP, parent=self)


# ========== 设置 ==========
class SettingsPage(ScrollArea):
    def __init__(self, parent: MainWindow):
        super().__init__(); self.main = parent; self.setObjectName("settingsPage"); self._setup_ui()
    def _setup_ui(self):
        w = QWidget(); w.setStyleSheet("background:transparent;"); l = QVBoxLayout(w); l.setSpacing(12); l.setContentsMargins(32, 24, 32, 24)
        l.addWidget(TitleLabel("设置")); l.addSpacing(4)

        # NAS 配置
        ct = SimpleCardWidget(); ct.setBorderRadius(12)
        ctl = QVBoxLayout(ct); ctl.setContentsMargins(24, 16, 24, 16)
        ctl.addWidget(StrongBodyLabel("NAS 配置"))
        nr = QHBoxLayout(); nr.addWidget(BodyLabel("NAS路径:"))
        self.nas_path_edit = LineEdit(); self.nas_path_edit.setText(self.main.settings.get("nas_path", ""))
        nr.addWidget(self.nas_path_edit, 1); ctl.addLayout(nr)
        url = QHBoxLayout(); url.addWidget(BodyLabel("用户名:"))
        self.nas_user_edit = LineEdit(); self.nas_user_edit.setText(self.main.settings.get("nas_user", ""))
        url.addWidget(self.nas_user_edit, 1); ctl.addLayout(url)
        pwr = QHBoxLayout(); pwr.addWidget(BodyLabel("密码:"))
        self.nas_pass_edit = PasswordLineEdit(); self.nas_pass_edit.setText(self.main.settings.get("nas_pass", ""))
        pwr.addWidget(self.nas_pass_edit, 1); ctl.addLayout(pwr)
        btn_save_nas = PrimaryPushButton("保存NAS配置")
        btn_save_nas.clicked.connect(self._save_nas)
        ctl.addWidget(btn_save_nas)
        btn_test_nas = PushButton("测试连接")
        btn_test_nas.clicked.connect(self._test_nas)
        ctl.addWidget(btn_test_nas)
        l.addWidget(ct)

        ct2 = SimpleCardWidget(); ct2.setBorderRadius(12)
        ct2l = QVBoxLayout(ct2); ct2l.setContentsMargins(24, 16, 24, 16)
        ct2l.addWidget(StrongBodyLabel("外观"))
        tr = QHBoxLayout(); tr.addWidget(BodyLabel("深色模式"))
        self.theme_switch = SwitchButton(); self.theme_switch.setChecked(self.main.settings.get("theme", "light") == "dark")
        self.theme_switch.checkedChanged.connect(self._on_theme)
        tr.addWidget(self.theme_switch); tr.addStretch(); ct2l.addLayout(tr)
        l.addWidget(ct2)

        c1 = SimpleCardWidget(); c1.setBorderRadius(12)
        c1l = QVBoxLayout(c1); c1l.setContentsMargins(24, 16, 24, 16)
        c1l.addWidget(StrongBodyLabel("应用信息"))
        for k,v in [("应用名称", APP_NAME),("版本", VERSION),("开发人", DEVELOPER),("联系电话", PHONE)]:
            r = QHBoxLayout(); r.addWidget(BodyLabel(f"{k}:")); lb = BodyLabel(v); lb.setStyleSheet("color:#666;"); r.addWidget(lb); r.addStretch(); c1l.addLayout(r)
        l.addWidget(c1)

        c3 = SimpleCardWidget(); c3.setBorderRadius(12)
        c3l = QVBoxLayout(c3); c3l.setContentsMargins(24, 16, 24, 16)
        c3l.addWidget(StrongBodyLabel("关于"))
        c3l.addWidget(BodyLabel(f"{APP_NAME} - 轻量客户端版，处理引擎部署在NAS"))
        lb = BodyLabel(f"开发人: {DEVELOPER} | 电话: {PHONE}"); lb.setStyleSheet("color:#888;"); c3l.addWidget(lb)
        l.addWidget(c3); l.addStretch(); self.setWidget(w); self.setWidgetResizable(True)

    def _on_theme(self, is_dark):
        setTheme(Theme.DARK if is_dark else Theme.LIGHT)
        self.main.settings["theme"] = "dark" if is_dark else "light"
        save_settings(self.main.settings)
    def _save_nas(self):
        self.main.settings["nas_path"] = self.nas_path_edit.text().strip()
        self.main.settings["nas_user"] = self.nas_user_edit.text().strip()
        self.main.settings["nas_pass"] = self.nas_pass_edit.text().strip()
        save_settings(self.main.settings)
        InfoBar.success("已保存", "NAS配置已保存", orient=Qt.Horizontal, isClosable=True, position=InfoBarPosition.TOP, parent=self)

    def _test_nas(self):
        """测试 NAS 连接"""
        from config import mount_nas
        path = self.nas_path_edit.text().strip()
        user = self.nas_user_edit.text().strip()
        pwd = self.nas_pass_edit.text().strip()
        ok = mount_nas(path, user, pwd)
        if ok:
            InfoBar.success("连接成功", f"已连接到: {path}", orient=Qt.Horizontal, isClosable=True, position=InfoBarPosition.TOP, parent=self)
        else:
            InfoBar.error("连接失败", f"无法连接到: {path}", orient=Qt.Horizontal, isClosable=True, position=InfoBarPosition.TOP, parent=self)


# ========== 邮箱反馈 ==========
class EmailPage(ScrollArea):
    def __init__(self, parent: MainWindow):
        super().__init__(); self.main = parent; self.setObjectName("emailPage"); self._setup_ui()
    def _setup_ui(self):
        w = QWidget(); w.setStyleSheet("background:transparent;"); l = QVBoxLayout(w); l.setSpacing(12); l.setContentsMargins(32, 24, 32, 24)
        l.addWidget(TitleLabel("邮箱反馈")); l.addSpacing(4)
        c = SimpleCardWidget(); c.setBorderRadius(12)
        cl = QVBoxLayout(c); cl.setContentsMargins(24, 16, 24, 16)
        cl.addWidget(StrongBodyLabel("发送邮件反馈"))
        r1 = QHBoxLayout(); r1.addWidget(BodyLabel("收件人:"))
        self.to_edit = LineEdit(); self.to_edit.setPlaceholderText("收件人邮箱"); self.to_edit.setText("2787326121@qq.com"); r1.addWidget(self.to_edit, 1); cl.addLayout(r1)
        r2 = QHBoxLayout(); r2.addWidget(BodyLabel("主题:"))
        self.subject_edit = LineEdit(); self.subject_edit.setPlaceholderText("反馈主题"); r2.addWidget(self.subject_edit, 1); cl.addLayout(r2)
        cl.addWidget(BodyLabel("内容:"))
        self.body_edit = TextEdit(); self.body_edit.setPlaceholderText("请描述您的问题或建议..."); self.body_edit.setMinimumHeight(150); cl.addWidget(self.body_edit)
        br = QHBoxLayout()
        self.btn_send = PrimaryPushButton(FIC.SEND, "发送反馈"); self.btn_send.clicked.connect(self.send); br.addWidget(self.btn_send); br.addStretch(); cl.addLayout(br)
        l.addWidget(c)
        ic = SimpleCardWidget(); ic.setBorderRadius(12)
        il = QVBoxLayout(ic); il.setContentsMargins(24, 16, 24, 16)
        il.addWidget(StrongBodyLabel("邮箱配置"))
        il.addWidget(BodyLabel("服务器: smtp.qq.com (SSL: 465)"))
        il.addWidget(BodyLabel("发件人: 2787326121@qq.com"))
        il.addWidget(BodyLabel("授权码: nfjrsfsqiorkdega"))
        l.addWidget(ic); l.addStretch(); self.setWidget(w); self.setWidgetResizable(True)
    def send(self):
        to = self.to_edit.text().strip(); sub = self.subject_edit.text().strip(); body = self.body_edit.toPlainText().strip()
        if not to or not body:
            InfoBar.warning("提示", "请填写收件人和内容", orient=Qt.Horizontal, isClosable=True, position=InfoBarPosition.TOP, parent=self); return
        self.btn_send.setEnabled(False); self.btn_send.setText("发送中...")
        self._ew = EmailWorker(to, sub, body)
        self._ew.log_signal.connect(lambda m: self.main.log_page.add_log(m))
        self._ew.done.connect(self._od)
        self._ew.start()
    def _od(self, ok, msg):
        self.btn_send.setEnabled(True); self.btn_send.setText("发送反馈")
        if ok: InfoBar.success("成功", msg, orient=Qt.Horizontal, isClosable=True, position=InfoBarPosition.TOP, parent=self); self.body_edit.clear()
        else: InfoBar.error("发送失败", msg, orient=Qt.Horizontal, isClosable=True, position=InfoBarPosition.TOP, parent=self)


# ========== 使用教程 ==========
class TutorialPage(ScrollArea):
    def __init__(self, parent: MainWindow):
        super().__init__(); self.main = parent; self.setObjectName("tutorialPage")
        w = QWidget(); w.setStyleSheet("background:transparent;"); l = QVBoxLayout(w); l.setSpacing(10); l.setContentsMargins(32, 24, 32, 24)
        l.addWidget(TitleLabel("使用教程")); l.addSpacing(4)
        steps = [("① 上传文件", "点击「浏览文件」选择Excel文件，或从下拉列表选最近使用的。"),
                 ("② 设置参数", "确认年份1/年份2，默认为25/26。"),
                 ("③ 一键处理", "文件上传至NAS服务器，自动完成全部任务。"),
                 ("④ 查看日志", "处理日志可在「处理日志」中查看。"),
                 ("⑤ 下载结果", "处理完成后点击「📥 下载结果文件」。"),
                 ("⑥ 输出目录", "结果自动保存在应用目录下 SOP_docs 文件夹。")]
        for t,d in steps:
            c = SimpleCardWidget(); c.setBorderRadius(12); cl = QVBoxLayout(c); cl.setContentsMargins(24, 12, 24, 12)
            cl.addWidget(StrongBodyLabel(t)); cl.addWidget(BodyLabel(d)); l.addWidget(c)
        l.addStretch(); self.setWidget(w); self.setWidgetResizable(True)


# ========== 版本信息 ==========
class VersionPage(ScrollArea):
    def __init__(self, parent: MainWindow):
        super().__init__(); self.main = parent; self.setObjectName("versionPage")
        w = QWidget(); w.setStyleSheet("background:transparent;"); l = QVBoxLayout(w); l.setSpacing(12); l.setContentsMargins(32, 24, 32, 24)
        l.addWidget(TitleLabel("版本信息")); l.addSpacing(4)
        c = SimpleCardWidget(); c.setBorderRadius(12)
        cl = QVBoxLayout(c); cl.setContentsMargins(24, 20, 24, 20)
        for k,v in [("应用", f"{APP_NAME} {VERSION}"),("开发人", DEVELOPER),("联系电话", PHONE)]:
            r = QHBoxLayout(); r.addWidget(BodyLabel(f"{k}:")); lb = BodyLabel(v); lb.setStyleSheet("color:#666;"); r.addWidget(lb); r.addStretch(); cl.addLayout(r)
        l.addWidget(c)
        c2 = SimpleCardWidget(); c2.setBorderRadius(12)
        c2l = QVBoxLayout(c2); c2l.setContentsMargins(24, 20, 24, 20)
        c2l.addWidget(StrongBodyLabel("版权信息"))
        c2l.addWidget(BodyLabel(f"{APP_NAME} - 轻量客户端版，处理引擎部署在NAS"))
        lb = BodyLabel(f"© {DEVELOPER} | 电话: {PHONE}"); lb.setStyleSheet("color:#888;"); c2l.addWidget(lb)
        l.addWidget(c2); l.addStretch()
        self.setWidget(w); self.setWidgetResizable(True)


# ========== 启动 ==========
if __name__ == "__main__":
    QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    app = QApplication(sys.argv); app.setFont(QFont("微软雅黑", 9))
    w = MainWindow(); w.show(); sys.exit(app.exec_())
