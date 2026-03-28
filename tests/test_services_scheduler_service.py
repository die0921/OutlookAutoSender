import pytest
from datetime import datetime, date, timedelta
from unittest.mock import MagicMock, call, patch
from src.services.scheduler_service import SchedulerService, SchedulerStatus
from src.models.email_task import EmailTask
from src.models.config import (
    Config, AppConfig, ExcelConfig, OutlookConfig,
    WorkdaysConfig, ErrorHandlingConfig, LoggingConfig
)


def make_config():
    return Config(
        app=AppConfig(name="Test", check_interval=60, expiry_notice_minutes=30),
        excel=ExcelConfig(main_file="test.xlsx", sheet_name="Sheet1", columns={}),
        outlook=OutlookConfig(account="", save_to_sent=True),
        workdays=WorkdaysConfig(weekdays=[1,2,3,4,5], holidays=[]),
        error_handling=ErrorHandlingConfig(
            stop_scheduler=True, skip_and_continue=False,
            mark_failed_as="发送失败", show_dialog=True, sound_alert=False
        ),
        logging=LoggingConfig(level="INFO", file="logs/app.log", max_size_mb=10, backup_count=5),
        status_values={"pending": "待发送", "sent": "已发送", "failed": "发送失败"},
        send_mode_values={"repeat": "重复", "once": "一次性"},
        templates_file="templates.yaml"
    )


def make_task(row_index=1, status='待发送', send_mode='一次性',
              planned_time=None, last_sent_date=None, schedule_type='按天',
              schedule_params='9:00'):
    return EmailTask(
        row_index=row_index,
        recipient='user@example.com',
        cc='',
        subject='Test',
        attachments='',
        status=status,
        send_mode=send_mode,
        template_name='',
        schedule_type=schedule_type,
        schedule_params=schedule_params,
        planned_time=planned_time,
        last_sent_date=last_sent_date,
        sent_time=None,
        variables={},
    )


@pytest.fixture
def mock_deps():
    reader = MagicMock()
    writer = MagicMock()
    email_svc = MagicMock()
    calc = MagicMock()
    return reader, writer, email_svc, calc


@pytest.fixture
def scheduler(mock_deps):
    reader, writer, email_svc, calc = mock_deps
    return SchedulerService(
        config=make_config(),
        excel_reader=reader,
        excel_writer=writer,
        email_service=email_svc,
        schedule_calculator=calc,
    )


def test_initial_status(scheduler):
    status = scheduler.get_status()
    assert status.running is False
    assert status.last_check is None
    assert status.error_message is None


def test_start_sets_running(scheduler, mock_deps):
    reader, _, email_svc, calc = mock_deps
    now = datetime.now()
    # No pending tasks to send
    task = make_task(status='待发送', planned_time=None)
    reader.read_tasks.return_value = [task]
    calc.calculate_next_time.return_value = now + timedelta(hours=1)

    scheduler.start()
    assert scheduler.get_status().running is True


def test_stop(scheduler, mock_deps):
    reader, _, _, calc = mock_deps
    reader.read_tasks.return_value = []
    scheduler.start()
    scheduler.stop()
    assert scheduler.get_status().running is False


def test_send_due_task(scheduler, mock_deps):
    reader, writer, email_svc, calc = mock_deps
    now = datetime.now()
    past_time = now - timedelta(minutes=5)

    task = make_task(status='待发送', planned_time=past_time)
    reader.read_tasks.return_value = [task]
    email_svc.send_email.return_value = (True, '')
    calc.should_reset_status.return_value = False
    calc.calculate_next_time.return_value = None

    scheduler.start()

    email_svc.send_email.assert_called_once_with(task)
    writer.update_sent.assert_called_once()


def test_failed_send_stops_scheduler(scheduler, mock_deps):
    reader, writer, email_svc, calc = mock_deps
    now = datetime.now()
    past_time = now - timedelta(minutes=5)

    task = make_task(status='待发送', planned_time=past_time)
    reader.read_tasks.return_value = [task]
    email_svc.send_email.return_value = (False, '网络错误')

    scheduler.start()

    assert scheduler.get_status().running is False
    assert scheduler.get_status().error_message == '网络错误'
    writer.update_status.assert_called_with('test.xlsx', 1, '发送失败')


def test_reset_repeat_task(scheduler, mock_deps):
    reader, writer, email_svc, calc = mock_deps
    yesterday = date.today() - timedelta(days=1)

    # First read: sent repeat task that needs reset
    sent_task = make_task(status='已发送', send_mode='重复',
                          last_sent_date=yesterday, planned_time=None)
    # Second read (after reset): pending task without planned time
    pending_task = make_task(status='待发送', send_mode='重复', planned_time=None)
    # Third read (after calc): still no due tasks
    reader.read_tasks.side_effect = [
        [sent_task],   # first read for reset check
        [pending_task], # second read for time calculation
        [],            # third read for due tasks
    ]
    calc.should_reset_status.return_value = True
    calc.calculate_next_time.return_value = datetime.now() + timedelta(hours=1)

    scheduler.start()

    writer.reset_to_pending.assert_called_once_with('test.xlsx', 1)


def test_on_error_callback(mock_deps):
    reader, writer, email_svc, calc = mock_deps
    error_callback = MagicMock()
    now = datetime.now()

    scheduler = SchedulerService(
        config=make_config(),
        excel_reader=reader,
        excel_writer=writer,
        email_service=email_svc,
        schedule_calculator=calc,
        on_error=error_callback,
    )

    past_time = now - timedelta(minutes=5)
    task = make_task(status='待发送', planned_time=past_time)
    reader.read_tasks.return_value = [task]
    email_svc.send_email.return_value = (False, '连接失败')

    scheduler.start()

    error_callback.assert_called_once()
