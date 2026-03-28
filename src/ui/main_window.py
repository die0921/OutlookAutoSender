from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTableWidget, QTableWidgetItem, QTextEdit, QPushButton,
    QLabel, QStatusBar, QAction, QHeaderView,
    QSplitter, QGroupBox
)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QColor
from typing import List, Optional
from src.models.email_task import EmailTask


class MainWindow(QMainWindow):
    """Outlook 自动发送助手主窗口"""

    # 信号
    start_requested = pyqtSignal()
    stop_requested = pyqtSignal()
    manual_send_requested = pyqtSignal()
    refresh_requested = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Outlook 自动发送助手")
        self.setMinimumSize(900, 600)
        self._setup_ui()
        self._connect_signals()

    def _setup_ui(self) -> None:
        """初始化 UI 组件"""
        self._setup_menu()
        self._setup_central_widget()
        self._setup_status_bar()

    def _setup_menu(self) -> None:
        """设置菜单栏"""
        menubar = self.menuBar()

        # 配置菜单
        config_menu = menubar.addMenu("配置")
        self.action_config = QAction("打开配置", self)
        config_menu.addAction(self.action_config)

        # 模板菜单
        template_menu = menubar.addMenu("模板管理")
        self.action_templates = QAction("管理模板", self)
        template_menu.addAction(self.action_templates)

        # 历史记录菜单
        history_menu = menubar.addMenu("历史记录")
        self.action_history = QAction("查看历史", self)
        history_menu.addAction(self.action_history)

        # 关于菜单
        about_menu = menubar.addMenu("关于")
        self.action_about = QAction("关于本软件", self)
        about_menu.addAction(self.action_about)

    def _setup_central_widget(self) -> None:
        """设置主区域"""
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        # 系统状态区
        info_group = QGroupBox("系统状态")
        info_layout = QHBoxLayout(info_group)
        self.status_label = QLabel("状态: 已停止")
        self.file_label = QLabel("文件: 未选择")
        info_layout.addWidget(self.status_label)
        info_layout.addStretch()
        info_layout.addWidget(self.file_label)
        layout.addWidget(info_group)

        # 控制按钮
        btn_layout = QHBoxLayout()
        self.start_btn = QPushButton("启动")
        self.stop_btn = QPushButton("停止")
        self.manual_btn = QPushButton("立即发送")
        self.refresh_btn = QPushButton("刷新")
        self.stop_btn.setEnabled(False)
        btn_layout.addWidget(self.start_btn)
        btn_layout.addWidget(self.stop_btn)
        btn_layout.addWidget(self.manual_btn)
        btn_layout.addWidget(self.refresh_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        # 分割器：上方表格 + 下方日志
        splitter = QSplitter(Qt.Vertical)

        # 待发送列表
        task_group = QGroupBox("待发送邮件")
        task_layout = QVBoxLayout(task_group)
        self.task_table = QTableWidget()
        self.task_table.setColumnCount(5)
        self.task_table.setHorizontalHeaderLabels(
            ["收件人", "主题", "发送方式", "计划时间", "状态"]
        )
        self.task_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.task_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.task_table.setSelectionBehavior(QTableWidget.SelectRows)
        task_layout.addWidget(self.task_table)
        splitter.addWidget(task_group)

        # 日志区
        log_group = QGroupBox("运行日志")
        log_layout = QVBoxLayout(log_group)
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(180)
        log_layout.addWidget(self.log_text)
        splitter.addWidget(log_group)

        layout.addWidget(splitter)

    def _setup_status_bar(self) -> None:
        """设置状态栏"""
        self.statusBar().showMessage("就绪")

    def _connect_signals(self) -> None:
        """连接按钮信号"""
        self.start_btn.clicked.connect(self.start_requested)
        self.stop_btn.clicked.connect(self.stop_requested)
        self.manual_btn.clicked.connect(self.manual_send_requested)
        self.refresh_btn.clicked.connect(self.refresh_requested)

    # --- Public API ---

    def update_tasks(self, tasks: List[EmailTask]) -> None:
        """更新待发送任务列表"""
        self.task_table.setRowCount(len(tasks))
        for row, task in enumerate(tasks):
            self.task_table.setItem(row, 0, QTableWidgetItem(task.recipient))
            self.task_table.setItem(row, 1, QTableWidgetItem(task.subject))
            self.task_table.setItem(row, 2, QTableWidgetItem(task.schedule_type))
            planned = (task.planned_time.strftime('%Y-%m-%d %H:%M')
                       if task.planned_time else '—')
            self.task_table.setItem(row, 3, QTableWidgetItem(planned))
            self.task_table.setItem(row, 4, QTableWidgetItem(task.status))

    def set_running_state(self, running: bool) -> None:
        """更新运行状态显示"""
        if running:
            self.status_label.setText("状态: 运行中")
            self.start_btn.setEnabled(False)
            self.stop_btn.setEnabled(True)
            self.statusBar().showMessage("调度器运行中")
        else:
            self.status_label.setText("状态: 已停止")
            self.start_btn.setEnabled(True)
            self.stop_btn.setEnabled(False)
            self.statusBar().showMessage("调度器已停止")

    def set_file_path(self, path: str) -> None:
        """显示当前 Excel 文件路径"""
        self.file_label.setText(f"文件: {path}")

    def append_log(self, message: str) -> None:
        """追加日志消息"""
        self.log_text.append(message)
        scrollbar = self.log_text.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def show_error(self, message: str) -> None:
        """在状态栏显示错误"""
        self.statusBar().showMessage(f"错误: {message}")
