from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTableWidget,
    QTableWidgetItem, QPushButton, QLabel, QHeaderView
)
from PyQt5.QtCore import pyqtSignal
from typing import List
from src.models.email_task import EmailTask


class FailedItemsManager(QDialog):
    """失败项管理对话框：批量处理发送失败的邮件"""
    reset_requested = pyqtSignal(list)  # list of row_index

    def __init__(self, tasks: List[EmailTask] = None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("失败项管理")
        self.setMinimumSize(650, 400)
        self._tasks = [t for t in (tasks or []) if t.status == '发送失败']
        self._setup_ui()
        self._load_tasks()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("以下邮件发送失败，请选择处理方式："))

        self.failed_table = QTableWidget()
        self.failed_table.setColumnCount(3)
        self.failed_table.setHorizontalHeaderLabels(["行号", "收件人", "主题"])
        self.failed_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.failed_table.setSelectionBehavior(QTableWidget.SelectRows)
        layout.addWidget(self.failed_table)

        btn_layout = QHBoxLayout()
        self.reset_btn = QPushButton("重置为待发送")
        self.reset_btn.clicked.connect(self._on_reset)
        btn_layout.addWidget(self.reset_btn)
        select_all_btn = QPushButton("全选")
        select_all_btn.clicked.connect(self.failed_table.selectAll)
        btn_layout.addWidget(select_all_btn)
        btn_layout.addStretch()
        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(self.close)
        btn_layout.addWidget(close_btn)
        layout.addLayout(btn_layout)

    def _load_tasks(self):
        self.failed_table.setRowCount(len(self._tasks))
        for row, task in enumerate(self._tasks):
            self.failed_table.setItem(row, 0, QTableWidgetItem(str(task.row_index)))
            self.failed_table.setItem(row, 1, QTableWidgetItem(task.recipient))
            self.failed_table.setItem(row, 2, QTableWidgetItem(task.subject))

    def _on_reset(self):
        selected_rows = {idx.row() for idx in self.failed_table.selectedIndexes()}
        row_indices = [self._tasks[r].row_index for r in selected_rows]
        if row_indices:
            self.reset_requested.emit(row_indices)

    def get_failed_count(self) -> int:
        return self.failed_table.rowCount()
