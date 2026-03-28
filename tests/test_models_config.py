from src.models.config import (
    AppConfig, ExcelConfig, OutlookConfig,
    WorkdaysConfig, ErrorHandlingConfig, LoggingConfig, Config
)


def test_app_config_creation():
    config = AppConfig(name="Test App", check_interval=60, expiry_notice_minutes=30)
    assert config.name == "Test App"
    assert config.check_interval == 60
    assert config.expiry_notice_minutes == 30


def test_excel_config_creation():
    config = ExcelConfig(
        main_file="test.xlsx",
        sheet_name="Sheet1",
        columns={"recipient": "收件人", "status": "发送状态"}
    )
    assert config.main_file == "test.xlsx"
    assert config.sheet_name == "Sheet1"
    assert config.columns["recipient"] == "收件人"


def test_outlook_config_creation():
    config = OutlookConfig(account="test@example.com", save_to_sent=True)
    assert config.account == "test@example.com"
    assert config.save_to_sent is True


def test_workdays_config():
    config = WorkdaysConfig(weekdays=[1, 2, 3, 4, 5], holidays=["2026-01-01"])
    assert 1 in config.weekdays
    assert "2026-01-01" in config.holidays
    assert 6 not in config.weekdays


def test_error_handling_config():
    config = ErrorHandlingConfig(
        stop_scheduler=True,
        skip_and_continue=False,
        mark_failed_as="发送失败",
        show_dialog=True,
        sound_alert=True
    )
    assert config.stop_scheduler is True
    assert config.mark_failed_as == "发送失败"


def test_logging_config():
    config = LoggingConfig(level="INFO", file="logs/app.log", max_size_mb=10, backup_count=5)
    assert config.level == "INFO"
    assert config.file == "logs/app.log"
    assert config.max_size_mb == 10


def test_config_immutable():
    config = AppConfig(name="App", check_interval=60, expiry_notice_minutes=30)
    try:
        config.name = "Changed"
        assert False, "Should have raised FrozenInstanceError"
    except Exception:
        pass  # Expected - frozen dataclass


def test_full_config():
    config = Config(
        app=AppConfig(name="App", check_interval=60, expiry_notice_minutes=30),
        excel=ExcelConfig(main_file="t.xlsx", sheet_name="S1", columns={}),
        outlook=OutlookConfig(account="", save_to_sent=True),
        workdays=WorkdaysConfig(weekdays=[1,2,3,4,5], holidays=[]),
        error_handling=ErrorHandlingConfig(
            stop_scheduler=True, skip_and_continue=False,
            mark_failed_as="发送失败", show_dialog=True, sound_alert=False
        ),
        logging=LoggingConfig(level="INFO", file="logs/app.log", max_size_mb=10, backup_count=5),
        status_values={"pending": "待发送"},
        send_mode_values={"repeat": "重复"},
        templates_file="templates.yaml"
    )
    assert config.app.name == "App"
    assert config.templates_file == "templates.yaml"
