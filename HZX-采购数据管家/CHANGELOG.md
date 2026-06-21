# HZX-采购数据管家 更新日志

## V2.0.0 (2026-06-21)

### 🎯 新增功能

#### 深色模式
- 设置页面新增「深色模式」开关按钮
- 实时切换 Theme.DARK / Theme.LIGHT
- 右侧内容区同步深色背景 (#1C1C1C)，修复白色背景文字不可见问题
- 所有页面容器与 ScrollArea 视口自适应透明背景

#### 参数配置持久化
- 年份 1/年份 2 参数保存至 `settings.json`
- 深色模式状态持久化
- 启动时自动加载上次保存的配置
- 设置页面新增「保存参数」按钮

#### NAS 网络共享支持（客户端增强）
- 启动时自动映射 NAS 共享目录
- 支持 SMB 协议直连 (`\\192.168.40.3\钟正洋`)
- 设置页面可配置 NAS 路径/账号/密码
- 「测试连接」按钮检测连通性
- SOP 页面新增「浏览NAS」按钮，直接打开 NAS 目录选文件

### 🔄 变更

#### 数据分析表排序调整
- **customs（是否报关）列排序改为降序**
  - 旧：`(空白)` → `是`（升序）
  - **新：`是` → `(空白)`（降序）**
  - 涉及：`build_amazon_hierarchy()` 中 `sorted(cs_map.items(), reverse=True)`
- **reason_m（不在义乌出的原因）按拼音字母排序**
  - 新增 `_pinyin_key()` 排序函数，引入 `pypinyin` 库
  - 排序：**中文原因（A→Z 拼音）→ `(空白)`**
  - `reason_n` 和 `reason_o` 保持降序（中文 → `(空白)`）

### 🐛 问题修复

| 问题 | 修复 |
|---|---|
| 一键处理按钮卡在「处理中」状态 | 改用 `QThread` + `pyqtSignal` 替代 `threading.Thread` + `QTimer.singleShot` |
| 邮箱反馈「发送中」状态不变回 | 使用 `EmailWorker(QThread)` 方式，`finally` 块恢复按钮 |
| 文件列表按钮点击无反应 | `PushButton` 变量引用修复 |
| 处理中按钮变形 | 固定正方形尺寸 `180x180`，圆角 `90px` |
| 加载动画导致按钮偏移 | 分离为独立 `status_row` 布局 |
| 深色模式右侧白底 | `setCustomBackgroundColor(QColor(28,28,28))` + `isTransparent=True` |
| 输出文件路径在临时目录 | `sys.frozen` 判断，`BASE_DIR` 指向 exe 所在目录 |
| 输出文件名覆盖 | 默认名添加 `_YYYYMMDD_HHMMSS` 时间戳 |

### 📦 文件变更

```
config.py       → VERSION "V1.0.0" → "V2.0.0", 新增 load_settings/save_settings
app.py          → V2.0.0, 深色模式, 设置持久化, QThread 重写
main_processing.py → V2.0.0, customs 排序 reverse=True
settings.json   → 新增（持久化配置文件，自动生成）
HZX-采购数据管家.spec → V2.0.0
```

---

## V1.0.0 (2026-06-16)

### 初始版本

- 基础 PyQt5 + qfluentwidgets 应用框架
- 月度采购数据对比 SOP 五大处理任务
- 文件列表管理（上传/输出文件）
- 处理日志实时展示
- QQ 邮箱反馈功能
- 使用教程与版本信息页面
- 应用仪表盘（数据概览 + 快速跳转）
- 导航栏分组（工具管理 / 帮助）
- 一键处理圆形大按钮 Fluent Design
