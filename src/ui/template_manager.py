from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QListWidget,
    QTextEdit, QLabel, QPushButton, QSplitter,
    QGroupBox, QFormLayout, QLineEdit
)
from PyQt5.QtCore import Qt
from typing import Dict
from src.models.template import Template


class TemplateManager(QDialog):
    """模板管理对话框"""

    def __init__(self, templates: Dict[str, Template] = None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("模板管理")
        self.setMinimumSize(700, 500)
        self._templates = templates or {}
        self._setup_ui()
        self._load_templates()

    def _setup_ui(self):
        layout = QHBoxLayout(self)
        splitter = QSplitter(Qt.Horizontal)

        # 左侧：模板列表
        left = QGroupBox("模板列表")
        left_layout = QVBoxLayout(left)
        self.template_list = QListWidget()
        self.template_list.currentTextChanged.connect(self._on_template_selected)
        left_layout.addWidget(self.template_list)
        btn_layout = QHBoxLayout()
        self.new_btn = QPushButton("新建")
        self.delete_btn = QPushButton("删除")
        btn_layout.addWidget(self.new_btn)
        btn_layout.addWidget(self.delete_btn)
        left_layout.addLayout(btn_layout)
        splitter.addWidget(left)

        # 右侧：模板详情
        right = QGroupBox("模板详情")
        right_layout = QVBoxLayout(right)
        form = QFormLayout()
        self.name_label = QLabel()
        form.addRow("名称:", self.name_label)
        self.subject_label = QLabel()
        form.addRow("主题:", self.subject_label)
        self.type_label = QLabel()
        form.addRow("类型:", self.type_label)
        right_layout.addLayout(form)
        right_layout.addWidget(QLabel("正文部分:"))
        self.body_text = QTextEdit()
        self.body_text.setReadOnly(True)
        right_layout.addWidget(self.body_text)
        right_layout.addWidget(QLabel("变量列表:"))
        self.vars_label = QLabel()
        self.vars_label.setWordWrap(True)
        right_layout.addWidget(self.vars_label)
        splitter.addWidget(right)

        splitter.setSizes([200, 500])
        layout.addWidget(splitter)

    def _load_templates(self):
        self.template_list.clear()
        for name in self._templates:
            self.template_list.addItem(name)

    def _on_template_selected(self, name: str):
        if name in self._templates:
            tmpl = self._templates[name]
            self.name_label.setText(tmpl.name)
            self.subject_label.setText(tmpl.subject)
            self.type_label.setText(tmpl.body_type)
            parts_text = '\n'.join(
                p.content or f'[{p.type}: {p.sheet}/{p.range}]'
                for p in tmpl.body_parts
            )
            self.body_text.setPlainText(parts_text)

    def get_template_count(self) -> int:
        return self.template_list.count()
