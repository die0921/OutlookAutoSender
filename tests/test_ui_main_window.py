import pytest
from datetime import datetime
from PyQt5.QtWidgets import QApplication
from src.ui.main_window import MainWindow
from src.models.email_task import EmailTask


@pytest.fixture(scope='module')
def app():
    import sys
    return QApplication.instance() or QApplication(sys.argv)


@pytest.fixture
def window(app, qtbot):
    w = MainWindow()
    qtbot.addWidget(w)
    return w


def make_task(row=1, recipient='user@example.com', subject='Test',
              planned_time=None):
    return EmailTask(
        row_index=row, recipient=recipient, cc='', subject=subject,
        attachments='', status='待发送', send_mode='一次性',
        template_name='', schedule_type='按天', schedule_params='9:00',
        planned_time=planned_time, last_sent_date=None, sent_time=None,
        variables={},
    )


def test_window_title(window):
    assert window.windowTitle() == "Outlook 自动发送助手"


def test_buttons_exist(window):
    assert window.start_btn is not None
    assert window.stop_btn is not None
    assert window.manual_btn is not None
    assert window.refresh_btn is not None


def test_stop_btn_initially_disabled(window):
    assert not window.stop_btn.isEnabled()


def test_update_tasks(window):
    tasks = [make_task(1, 'a@b.com', 'Subject1'),
             make_task(2, 'c@d.com', 'Subject2')]
    window.update_tasks(tasks)
    assert window.task_table.rowCount() == 2
    assert window.task_table.item(0, 0).text() == 'a@b.com'
    assert window.task_table.item(1, 1).text() == 'Subject2'


def test_set_running_state_true(window):
    window.set_running_state(True)
    assert not window.start_btn.isEnabled()
    assert window.stop_btn.isEnabled()
    assert '运行中' in window.status_label.text()


def test_set_running_state_false(window):
    window.set_running_state(False)
    assert window.start_btn.isEnabled()
    assert not window.stop_btn.isEnabled()
    assert '停止' in window.status_label.text()


def test_append_log(window):
    window.append_log("Test log message")
    assert "Test log message" in window.log_text.toPlainText()


def test_start_signal(window, qtbot):
    with qtbot.waitSignal(window.start_requested, timeout=1000):
        window.start_btn.click()


def test_stop_signal(window, qtbot):
    window.stop_btn.setEnabled(True)
    with qtbot.waitSignal(window.stop_requested, timeout=1000):
        window.stop_btn.click()
