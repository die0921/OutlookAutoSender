from src.models.config import (
    AppConfig,
    ExcelConfig,
    OutlookConfig,
    WorkdaysConfig,
    ErrorHandlingConfig,
    LoggingConfig,
    Config,
)


def test_app_config_creation():
    config = AppConfig(
        name="Test App",
        version="1.0.0",
        check_interval=60,
        expiry_notice_minutes=30,
    )
    assert config.name == "Test App"
    assert config.version == "1.0.0"
    assert config.check_interval == 60
    assert config.expiry_notice_minutes == 30


def test_excel_config_creation():
    columns = {"recipient": "收件人", "status": "发送状态", "subject": "主题"}
    config = ExcelConfig(
        file_path="test.xlsx",
        columns=columns,
    )
    assert config.file_path == "test.xlsx"
    assert config.columns["recipient"] == "收件人"
    assert config.columns["status"] == "发送状态"


def test_outlook_config_creation():
    config = OutlookConfig(
        auto_close_after_send=True,
        signature_enabled=True,
        importance="normal",
        read_receipt=False,
        delivery_receipt=False,
    )
    assert config.auto_close_after_send is True
    assert config.signature_enabled is True
    assert config.importance == "normal"
    assert config.read_receipt is False


def test_workdays_config():
    config = WorkdaysConfig(
        enabled=True,
        days=[1, 2, 3, 4, 5],
        start_time="09:00",
        end_time="18:00",
        skip_holidays=True,
        holidays=["2026-01-01", "2026-01-02"],
    )
    assert config.enabled is True
    assert 1 in config.days
    assert 5 in config.days
    assert "2026-01-01" in config.holidays
    assert config.start_time == "09:00"
    assert config.end_time == "18:00"


def test_error_handling_config():
    config = ErrorHandlingConfig(
        max_retries=3,
        retry_interval_seconds=60,
        log_errors=True,
        notify_on_failure=True,
    )
    assert config.max_retries == 3
    assert config.retry_interval_seconds == 60
    assert config.log_errors is True
    assert config.notify_on_failure is True


def test_logging_config():
    config = LoggingConfig(
        level="INFO",
        file_path="logs/app.log",
        max_bytes=10485760,
        backup_count=5,
    )
    assert config.level == "INFO"
    assert config.file_path == "logs/app.log"
    assert config.max_bytes == 10485760
    assert config.backup_count == 5


def test_full_config_creation():
    app_config = AppConfig(
        name="Outlook 自动发送助手",
        version="1.0.0",
        check_interval=60,
        expiry_notice_minutes=30,
    )
    excel_config = ExcelConfig(
        file_path="data.xlsx",
        columns={
            "status": "状态",
            "recipient": "收件人",
            "subject": "主题",
        },
    )
    outlook_config = OutlookConfig(
        auto_close_after_send=True,
        signature_enabled=True,
        importance="normal",
        read_receipt=False,
        delivery_receipt=False,
    )
    workdays_config = WorkdaysConfig(
        enabled=True,
        days=[1, 2, 3, 4, 5],
        start_time="09:00",
        end_time="18:00",
        skip_holidays=True,
        holidays=[],
    )
    error_handling_config = ErrorHandlingConfig(
        max_retries=3,
        retry_interval_seconds=60,
        log_errors=True,
        notify_on_failure=True,
    )
    logging_config = LoggingConfig(
        level="INFO",
        file_path="logs/app.log",
        max_bytes=10485760,
        backup_count=5,
    )

    status_values = {
        "pending": "待发送",
        "scheduled": "已排程",
        "sent": "已发送",
        "failed": "发送失败",
    }
    send_mode_values = {
        "immediate": "立即发送",
        "scheduled": "定时发送",
    }

    config = Config(
        app=app_config,
        excel=excel_config,
        outlook=outlook_config,
        workdays=workdays_config,
        error_handling=error_handling_config,
        logging=logging_config,
        status_values=status_values,
        send_mode_values=send_mode_values,
        templates_file="templates.yaml",
    )

    assert config.app.name == "Outlook 自动发送助手"
    assert config.excel.file_path == "data.xlsx"
    assert config.outlook.auto_close_after_send is True
    assert config.workdays.enabled is True
    assert config.error_handling.max_retries == 3
    assert config.logging.level == "INFO"
    assert config.status_values["pending"] == "待发送"
    assert config.send_mode_values["immediate"] == "立即发送"
    assert config.templates_file == "templates.yaml"


def test_config_frozen():
    """Test that config dataclasses are frozen (immutable)"""
    config = AppConfig(
        name="Test",
        version="1.0.0",
        check_interval=60,
        expiry_notice_minutes=30,
    )
    # Attempting to modify a frozen dataclass should raise an error
    import pytest

    with pytest.raises(Exception):  # FrozenInstanceError
        config.name = "Modified"  # type: ignore
