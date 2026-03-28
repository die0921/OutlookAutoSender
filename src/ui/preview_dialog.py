from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QFormLayout,
    QLabel, QTextEdit, QPushButton, QListWidget,
    QDialogButtonBox, QGroupBox
)
from PyQt5.QtCore import pyqtSignal
from src.services.email_service import EmailPreview


class PreviewDialog(QDialog):
    """邮件预览对话框"""
    send_now_requested = pyqtSignal()

    def __init__(self, preview: EmailPreview = None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("邮件预览")
        self.setMinimumSize(600, 500)
        self._setup_ui()
        if preview:
            self.load_preview(preview)

    def _setup_ui(self):
        layout = QVBoxLayout(self)

        # 元数据
        form_group = QGroupBox("邮件信息")
        form = QFormLayout(form_group)
        self.recipients_label = QLabel()
        self.recipients_label.setWordWrap(True)
        form.addRow("收件人:", self.recipients_label)
        self.cc_label = QLabel()
        form.addRow("抄送:", self.cc_label)
        self.subject_label = QLabel()
        form.addRow("主题:", self.subject_label)
        self.time_label = QLabel()
        form.addRow("计划时间:", self.time_label)
        layout.addWidget(form_group)

        # 正文
        body_group = QGroupBox("正文")
        body_layout = QVBoxLayout(body_group)
        self.body_text = QTextEdit()
        self.body_text.setReadOnly(True)
        body_layout.addWidget(self.body_text)
        layout.addWidget(body_group)

        # 附件
        attach_group = QGroupBox("附件")
        attach_layout = QVBoxLayout(attach_group)
        self.attach_list = QListWidget()
        self.attach_list.setMaximumHeight(80)
        attach_layout.addWidget(self.attach_list)
        layout.addWidget(attach_group)

        # 按钮
        btn_layout = QHBoxLayout()
        self.send_now_btn = QPushButton("立即发送")
        self.send_now_btn.clicked.connect(self.send_now_requested)
        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(self.close)
        btn_layout.addWidget(self.send_now_btn)
        btn_layout.addStretch()
        btn_layout.addWidget(close_btn)
        layout.addLayout(btn_layout)

    def load_preview(self, preview: EmailPreview):
        self.recipients_label.setText('; '.join(preview.recipients))
        self.cc_label.setText('; '.join(preview.cc_list))
        self.subject_label.setText(preview.subject)
        self.time_label.setText(preview.planned_time or '—')
        if '<html>' in preview.body.lower():
            self.body_text.setHtml(preview.body)
        else:
            self.body_text.setPlainText(preview.body)
        self.attach_list.clear()
        for attachment in preview.attachments:
            self.attach_list.addItem(attachment)
