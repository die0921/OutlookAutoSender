import pytest
from unittest.mock import MagicMock, patch
from datetime import datetime
from src.services.email_service import EmailService, EmailPreview
from src.models.email_task import EmailTask


def make_task(recipient='user@example.com', cc='', template_name='',
              subject='Test', attachments='', variables=None):
    return EmailTask(
        row_index=1,
        recipient=recipient,
        cc=cc,
        subject=subject,
        attachments=attachments,
        status='待发送',
        send_mode='一次性',
        template_name=template_name,
        schedule_type='按天',
        schedule_params='9:00',
        planned_time=datetime(2026, 3, 28, 9, 0),
        last_sent_date=None,
        sent_time=None,
        variables=variables or {},
    )


@pytest.fixture
def mock_template_engine():
    engine = MagicMock()
    engine.render.return_value = ('Rendered Subject', '<html><body>Rendered Body</body></html>')
    return engine


@pytest.fixture
def mock_contact_service():
    service = MagicMock()
    service.resolve_recipients.return_value = ['user@example.com']
    return service


@pytest.fixture
def mock_outlook():
    mail = MagicMock()
    mail.Recipients = MagicMock()
    mail.Recipients.Add.return_value = MagicMock()
    mail.Attachments = MagicMock()
    outlook = MagicMock()
    outlook.CreateItem.return_value = mail
    return outlook, mail


@pytest.fixture
def service(mock_template_engine, mock_contact_service, mock_outlook):
    outlook_app, _ = mock_outlook
    return EmailService(mock_template_engine, mock_contact_service, outlook_app)


def test_send_email_success(service, mock_outlook):
    _, mail = mock_outlook
    task = make_task(template_name='简单模板')
    success, error = service.send_email(task)
    assert success is True
    assert error == ''
    mail.Send.assert_called_once()


def test_send_email_calls_template_engine(service, mock_template_engine):
    task = make_task(template_name='月报模板', variables={'客户名称': '张三'})
    service.send_email(task)
    mock_template_engine.render.assert_called_once_with('月报模板', {'客户名称': '张三'})


def test_send_email_no_template_uses_subject(service, mock_outlook):
    _, mail = mock_outlook
    task = make_task(subject='直接主题', template_name='')
    service.send_email(task)
    assert mail.Subject == '直接主题'


def test_send_email_with_cc(service, mock_contact_service, mock_outlook):
    _, mail = mock_outlook
    mock_contact_service.resolve_recipients.side_effect = lambda x: [x]
    task = make_task(recipient='a@x.com', cc='b@x.com', template_name='简单模板')
    service.send_email(task)
    # Recipients.Add should be called for both to and cc
    calls = [str(c) for c in mail.Recipients.Add.call_args_list]
    assert len(calls) >= 2


def test_send_email_with_attachments(service, mock_outlook):
    _, mail = mock_outlook
    task = make_task(attachments='file1.pdf;file2.xlsx', template_name='简单模板')
    service.send_email(task)
    assert mail.Attachments.Add.call_count == 2


def test_send_email_failure_returns_false(mock_template_engine, mock_contact_service):
    mock_template_engine.render.side_effect = Exception("Render error")
    svc = EmailService(mock_template_engine, mock_contact_service)
    task = make_task(template_name='模板')
    success, error = svc.send_email(task)
    assert success is False
    assert 'Render error' in error


def test_preview_email(service):
    task = make_task(template_name='简单模板', variables={'客户名称': '张三'})
    preview = service.preview_email(task)
    assert isinstance(preview, EmailPreview)
    assert preview.recipients == ['user@example.com']
    assert preview.subject == 'Rendered Subject'
    assert preview.planned_time == '2026-03-28 09:00'


def test_parse_attachments_empty(service):
    result = service._parse_attachments('')
    assert result == []


def test_parse_attachments_multiple(service):
    result = service._parse_attachments('a.pdf;b.xlsx;c.docx')
    assert result == ['a.pdf', 'b.xlsx', 'c.docx']
