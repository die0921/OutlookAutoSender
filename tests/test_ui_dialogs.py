import pytest
import sys
from PyQt5.QtWidgets import QApplication
from src.ui.config_dialog import ConfigDialog
from src.ui.preview_dialog import PreviewDialog
from src.ui.error_dialog import ErrorDialog
from src.ui.template_manager import TemplateManager
from src.ui.history_viewer import HistoryViewer
from src.ui.failed_items_manager import FailedItemsManager
from src.services.email_service import EmailPreview
from src.models.email_task import EmailTask
from src.models.template import Template, BodyPart
from src.models.config import (
    Config, AppConfig, ExcelConfig, OutlookConfig,
    WorkdaysConfig, ErrorHandlingConfig, LoggingConfig
)


@pytest.fixture(scope='module')
def app():
    return QApplication.instance() or QApplication(sys.argv)


def make_task(status='已发送', row=1):
    return EmailTask(
        row_index=row, recipient='user@example.com', cc='', subject='Test',
        attachments='', status=status, send_mode='一次性', template_name='',
        schedule_type='按天', schedule_params='9:00',
        planned_time=None, last_sent_date=None, sent_time=None, variables={},
    )


def make_config():
    return Config(
        app=AppConfig(name="Test", check_interval=60, expiry_notice_minutes=30),
        excel=ExcelConfig(main_file="test.xlsx", sheet_name="Sheet1", columns={}),
        outlook=OutlookConfig(account="test@x.com", save_to_sent=True),
        workdays=WorkdaysConfig(weekdays=[1, 2, 3, 4, 5], holidays=[]),
        error_handling=ErrorHandlingConfig(
            stop_scheduler=True, skip_and_continue=False,
            mark_failed_as="发送失败", show_dialog=True, sound_alert=False
        ),
        logging=LoggingConfig(level="INFO", file="logs/app.log", max_size_mb=10, backup_count=5),
        status_values={}, send_mode_values={}, templates_file="templates.yaml"
    )


# --- ConfigDialog ---

def test_config_dialog_creates(app, qtbot):
    dlg = ConfigDialog()
    qtbot.addWidget(dlg)
    assert dlg.windowTitle() == "配置"


def test_config_dialog_load_config(app, qtbot):
    config = make_config()
    dlg = ConfigDialog(config=config)
    qtbot.addWidget(dlg)
    assert dlg.file_edit.text() == "test.xlsx"
    assert dlg.account_edit.text() == "test@x.com"
    assert dlg.interval_spin.value() == 60


def test_config_dialog_get_values(app, qtbot):
    dlg = ConfigDialog()
    qtbot.addWidget(dlg)
    dlg.file_edit.setText("my.xlsx")
    values = dlg.get_values()
    assert values['excel_file'] == "my.xlsx"


# --- PreviewDialog ---

def test_preview_dialog_creates(app, qtbot):
    dlg = PreviewDialog()
    qtbot.addWidget(dlg)
    assert dlg.windowTitle() == "邮件预览"


def test_preview_dialog_load_preview(app, qtbot):
    preview = EmailPreview(
        recipients=['a@b.com'],
        cc_list=[],
        subject='Test Subject',
        body='Hello',
        attachments=['file.pdf'],
        planned_time='2026-03-28 09:00',
    )
    dlg = PreviewDialog(preview=preview)
    qtbot.addWidget(dlg)
    assert 'a@b.com' in dlg.recipients_label.text()
    assert dlg.subject_label.text() == 'Test Subject'
    assert dlg.attach_list.count() == 1


# --- ErrorDialog ---

def test_error_dialog_creates(app, qtbot):
    dlg = ErrorDialog(row_index=2, recipient='x@y.com', error_message='Test error')
    qtbot.addWidget(dlg)
    assert dlg.windowTitle() == "发送失败"
    assert 'Test error' in dlg.error_text.toPlainText()


def test_error_dialog_signals(app, qtbot):
    dlg = ErrorDialog(row_index=1, recipient='a@b.com', error_message='err')
    qtbot.addWidget(dlg)
    assert dlg.open_excel_requested is not None
    assert dlg.skip_requested is not None
    assert dlg.stop_requested is not None


# --- TemplateManager ---

def test_template_manager_creates(app, qtbot):
    templates = {
        '简单模板': Template(
            name='简单模板', subject='{{主题}}', body_type='text',
            body_parts=[BodyPart(type='text', content='Hello')],
            data_sources={}
        )
    }
    dlg = TemplateManager(templates=templates)
    qtbot.addWidget(dlg)
    assert dlg.get_template_count() == 1


# --- HistoryViewer ---

def test_history_viewer_creates(app, qtbot):
    tasks = [make_task('已发送'), make_task('发送失败', row=2)]
    dlg = HistoryViewer(tasks=tasks)
    qtbot.addWidget(dlg)
    assert dlg.get_row_count() == 2


def test_history_viewer_filter(app, qtbot):
    tasks = [make_task('已发送'), make_task('发送失败', row=2)]
    dlg = HistoryViewer(tasks=tasks)
    qtbot.addWidget(dlg)
    dlg.status_filter.setCurrentText("已发送")
    assert dlg.get_row_count() == 1


# --- FailedItemsManager ---

def test_failed_items_manager_creates(app, qtbot):
    tasks = [make_task('发送失败'), make_task('已发送', row=2)]
    dlg = FailedItemsManager(tasks=tasks)
    qtbot.addWidget(dlg)
    # Only failed tasks shown
    assert dlg.get_failed_count() == 1
