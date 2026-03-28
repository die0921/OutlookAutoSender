from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QFormLayout,
    QLineEdit, QSpinBox, QCheckBox, QPushButton,
    QTabWidget, QWidget, QLabel, QFileDialog,
    QDialogButtonBox, QGroupBox
)
from PyQt5.QtCore import pyqtSignal
from src.models.config import Config


class ConfigDialog(QDialog):
    """配置编辑对话框"""
    config_saved = pyqtSignal(dict)

    def __init__(self, config: Config = None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("配置")
        self.setMinimumWidth(500)
        self._setup_ui()
        if config:
            self.load_config(config)

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        tabs = QTabWidget()

        # Excel 配置 tab
        excel_tab = QWidget()
        form = QFormLayout(excel_tab)
        self.file_edit = QLineEdit()
        browse_btn = QPushButton("浏览...")
        browse_btn.clicked.connect(self._browse_file)
        file_layout = QHBoxLayout()
        file_layout.addWidget(self.file_edit)
        file_layout.addWidget(browse_btn)
        form.addRow("Excel 文件:", file_layout)
        self.sheet_edit = QLineEdit()
        form.addRow("工作表名称:", self.sheet_edit)
        tabs.addTab(excel_tab, "Excel")

        # 定时配置 tab
        schedule_tab = QWidget()
        sform = QFormLayout(schedule_tab)
        self.interval_spin = QSpinBox()
        self.interval_spin.setRange(10, 3600)
        self.interval_spin.setSuffix(" 秒")
        sform.addRow("检查间隔:", self.interval_spin)
        self.expiry_spin = QSpinBox()
        self.expiry_spin.setRange(1, 120)
        self.expiry_spin.setSuffix(" 分钟")
        sform.addRow("过期提醒:", self.expiry_spin)
        tabs.addTab(schedule_tab, "定时")

        # Outlook 配置 tab
        outlook_tab = QWidget()
        oform = QFormLayout(outlook_tab)
        self.account_edit = QLineEdit()
        self.account_edit.setPlaceholderText("留空使用默认账户")
        oform.addRow("Outlook 账户:", self.account_edit)
        self.save_sent_check = QCheckBox("保存到已发送")
        self.save_sent_check.setChecked(True)
        oform.addRow("", self.save_sent_check)
        tabs.addTab(outlook_tab, "Outlook")

        layout.addWidget(tabs)

        # 按钮
        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btn_box.accepted.connect(self._on_accept)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)

    def _browse_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择 Excel 文件", "", "Excel 文件 (*.xlsx *.xls)")
        if path:
            self.file_edit.setText(path)

    def load_config(self, config: Config):
        self.file_edit.setText(config.excel.main_file)
        self.sheet_edit.setText(config.excel.sheet_name)
        self.interval_spin.setValue(config.app.check_interval)
        self.expiry_spin.setValue(config.app.expiry_notice_minutes)
        self.account_edit.setText(config.outlook.account)
        self.save_sent_check.setChecked(config.outlook.save_to_sent)

    def get_values(self) -> dict:
        return {
            'excel_file': self.file_edit.text(),
            'sheet_name': self.sheet_edit.text(),
            'check_interval': self.interval_spin.value(),
            'expiry_notice_minutes': self.expiry_spin.value(),
            'outlook_account': self.account_edit.text(),
            'save_to_sent': self.save_sent_check.isChecked(),
        }

    def _on_accept(self):
        self.config_saved.emit(self.get_values())
        self.accept()
