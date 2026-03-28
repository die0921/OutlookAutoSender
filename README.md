# Outlook 自动发送助手

基于 Python + PyQt5 的 Windows 桌面应用，通过 Outlook COM 接口自动定时发送邮件，支持 Excel 任务管理、Jinja2 模板渲染和 5 种定时方式。

## 功能特性

- **5 种定时方式**：按天重复、按周重复、按月固定日期、按月工作日、按日期列表
- **Excel 任务管理**：从 Excel 读取发送任务，自动回写发送状态和时间
- **Jinja2 模板系统**：支持变量替换、Excel 区域嵌入为 HTML 表格
- **Outlook 联系人组**：自动展开通讯组列表为邮箱地址
- **图形界面**：任务列表、运行日志、配置编辑、模板管理、历史记录查看
- **工作日计算**：支持自定义节假日，跳过非工作日

## 截图

> 主窗口：任务列表 + 运行日志 + 启动/停止控制

## 环境要求

- Windows 10/11
- Microsoft Outlook（已安装并配置账户）
- Python 3.8+（仅开发时需要）

## 快速开始

### 方式一：直接运行 exe（推荐）

1. 从 [Releases](../../releases) 下载 `OutlookAutoSender.exe`
2. 将 `config.yaml` 和 `templates.yaml` 放在同一目录
3. 双击运行

### 方式二：从源码运行

```bash
# 克隆项目
git clone https://github.com/die0921/OutlookAutoSender.git
cd OutlookAutoSender

# 创建虚拟环境
python -m venv venv
venv\Scripts\activate

# 安装依赖
pip install -r requirements.txt

# 初始化示例配置（首次运行）
python setup_configs.py

# 启动应用
python main.py
```

## 配置说明

### config.yaml

```yaml
excel:
  main_file: data/emails.xlsx   # 任务 Excel 文件路径
  sheet_name: Sheet1

outlook:
  account: your@email.com       # Outlook 账户
  save_to_sent: true

workdays:
  weekdays: [0, 1, 2, 3, 4]     # 0=周一 ... 4=周五
  holidays: []                   # 节假日列表，格式：YYYY-MM-DD
```

### templates.yaml

```yaml
templates:
  - name: 默认模板
    subject: "{{ title }}"
    body:
      - type: text
        content: "亲爱的 {{ name }}，\n\n{{ content }}"
```

### Excel 任务表格式

| 列名 | 说明 |
|------|------|
| 收件人 | 邮箱地址，多个用分号分隔 |
| 抄送 | 抄送邮箱（可选） |
| 主题 | 邮件主题（支持模板变量） |
| 模板 | templates.yaml 中的模板名称 |
| 发送方式 | daily / weekly / monthly_fixed / monthly_workday / date_list |
| 发送时间 | HH:MM 格式 |
| 状态 | 待发送 / 已发送 / 发送失败 |

## 打包为 exe

```bash
venv\Scripts\activate
pip install pyinstaller
# 双击 build.bat 或运行：
pyinstaller OutlookAutoSender.spec --clean
# 输出：dist\OutlookAutoSender.exe
```

## 开发

### 项目结构

```
OutlookAutoSender/
├── main.py                  # 入口
├── config.yaml              # 应用配置
├── templates.yaml           # 邮件模板
├── src/
│   ├── models/              # 数据模型（frozen dataclass）
│   ├── data/                # 数据访问层（Excel/配置/日志）
│   ├── services/            # 业务逻辑（调度/模板/发送）
│   ├── ui/                  # PyQt5 界面
│   └── utils/               # 工具函数
├── tests/                   # pytest 测试（含集成测试）
└── build.bat                # 打包脚本
```

### 运行测试

```bash
pytest --tb=short
```

### 技术栈

| 库 | 用途 |
|----|------|
| PyQt5 5.15 | 桌面 GUI |
| pywin32 | Outlook COM 接口 |
| openpyxl | Excel 读写 |
| APScheduler | 定时任务 |
| Jinja2 | 模板渲染 |
| PyYAML | 配置文件解析 |

## License

MIT
