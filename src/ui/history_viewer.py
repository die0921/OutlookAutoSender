from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTableWidget,
    QTableWidgetItem, QPushButton, QLabel, QComboBox,
    QHeaderView, QFileDialog
)
import csv
from typing import List
from src.models.email_task import EmailTask


class HistoryViewer(QDialog):
    """发送历史记录查看器"""

    def __init__(self, tasks: List[EmailTask] = None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("历史记录")
        self.setMinimumSize(700, 450)
        self._all_tasks = tasks or []
        self._setup_ui()
        self.load_history(self._all_tasks)

    def _setup_ui(self):
        layout = QVBoxLayout(self)

        # 筛选区
        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("状态筛选:"))
        self.status_filter = QComboBox()
        self.status_filter.addItems(["全部", "已发送", "发送失败", "待发送"])
        self.status_filter.currentTextChanged.connect(self._apply_filter)
        filter_layout.addWidget(self.status_filter)
        filter_layout.addStretch()
        export_btn = QPushButton("导出 CSV")
        export_btn.clicked.connect(self._export_csv)
        filter_layout.addWidget(export_btn)
        layout.addLayout(filter_layout)

        # 历史列表
        self.history_table = QTableWidget()
        self.history_table.setColumnCount(4)
        self.history_table.setHorizontalHeaderLabels(["收件人", "主题", "状态", "发送时间"])
        self.history_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.history_table.setEditTriggers(QTableWidget.NoEditTriggers)
        layout.addWidget(self.history_table)

        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(self.close)
        layout.addWidget(close_btn)

    def load_history(self, tasks: List[EmailTask]):
        self._all_tasks = tasks
        self._apply_filter(self.status_filter.currentText())

    def _apply_filter(self, status: str):
        if status == "全部":
            filtered = self._all_tasks
        else:
            filtered = [t for t in self._all_tasks if t.status == status]
        self.history_table.setRowCount(len(filtered))
        for row, task in enumerate(filtered):
            self.history_table.setItem(row, 0, QTableWidgetItem(task.recipient))
            self.history_table.setItem(row, 1, QTableWidgetItem(task.subject))
            self.history_table.setItem(row, 2, QTableWidgetItem(task.status))
            sent = (task.sent_time.strftime('%Y-%m-%d %H:%M')
                    if task.sent_time else '—')
            self.history_table.setItem(row, 3, QTableWidgetItem(sent))

    def _export_csv(self):
        path, _ = QFileDialog.getSaveFileName(self, "导出 CSV", "", "CSV 文件 (*.csv)")
        if path:
            with open(path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                writer.writerow(["收件人", "主题", "状态", "发送时间"])
                for task in self._all_tasks:
                    sent = (task.sent_time.strftime('%Y-%m-%d %H:%M')
                            if task.sent_time else '')
                    writer.writerow([task.recipient, task.subject, task.status, sent])

    def get_row_count(self) -> int:
        return self.history_table.rowCount()
