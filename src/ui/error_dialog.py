from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QTextEdit, QGroupBox
)
from PyQt5.QtCore import pyqtSignal


class ErrorDialog(QDialog):
    """发送失败错误处理对话框"""
    open_excel_requested = pyqtSignal()
    skip_requested = pyqtSignal()
    stop_requested = pyqtSignal()

    def __init__(self, row_index: int = 0, recipient: str = '',
                 error_message: str = '', parent=None):
        super().__init__(parent)
        self.setWindowTitle("发送失败")
        self.setMinimumWidth(450)
        self._setup_ui(row_index, recipient, error_message)

    def _setup_ui(self, row_index, recipient, error_message):
        layout = QVBoxLayout(self)

        # 错误详情
        detail_group = QGroupBox("错误详情")
        detail_layout = QVBoxLayout(detail_group)
        detail_layout.addWidget(QLabel(f"行号: {row_index}"))
        detail_layout.addWidget(QLabel(f"收件人: {recipient}"))
        self.error_text = QTextEdit()
        self.error_text.setReadOnly(True)
        self.error_text.setPlainText(error_message)
        self.error_text.setMaximumHeight(100)
        detail_layout.addWidget(self.error_text)
        layout.addWidget(detail_group)

        # 处理选项
        layout.addWidget(QLabel("请选择处理方式:"))
        open_btn = QPushButton("打开 Excel 并修复")
        open_btn.clicked.connect(self.open_excel_requested)
        open_btn.clicked.connect(self.accept)
        layout.addWidget(open_btn)

        skip_btn = QPushButton("跳过继续")
        skip_btn.clicked.connect(self.skip_requested)
        skip_btn.clicked.connect(self.accept)
        layout.addWidget(skip_btn)

        stop_btn = QPushButton("暂时停止")
        stop_btn.clicked.connect(self.stop_requested)
        stop_btn.clicked.connect(self.accept)
        layout.addWidget(stop_btn)
