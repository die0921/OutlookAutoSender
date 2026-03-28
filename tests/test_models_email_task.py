from src.models.email_task import EmailTask
from datetime import datetime, date


def test_email_task_creation():
    task = EmailTask(
        row_index=1,
        recipient="user@example.com",
        cc="",
        subject="Test Subject",
        attachments="",
        status="待发送",
        send_mode="一次性",
        template_name="",
        schedule_type="按天",
        schedule_params="9:00",
        planned_time=None,
        last_sent_date=None,
        sent_time=None,
        variables={"客户名称": "张三"}
    )

    assert task.row_index == 1
    assert task.recipient == "user@example.com"
    assert task.status == "待发送"
    assert task.variables["客户名称"] == "张三"


def test_email_task_status_checks():
    task = EmailTask(
        row_index=1,
        recipient="user@example.com",
        cc="",
        subject="Test",
        attachments="",
        status="待发送",
        send_mode="重复",
        template_name="",
        schedule_type="按天",
        schedule_params="9:00",
        planned_time=None,
        last_sent_date=None,
        sent_time=None,
        variables={}
    )

    assert task.is_pending() is True
    assert task.is_sent() is False
    assert task.is_failed() is False
    assert task.is_repeat_mode() is True
