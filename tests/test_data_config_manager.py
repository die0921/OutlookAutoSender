import pytest
import tempfile
import os
from src.data.config_manager import ConfigManager


CONFIG_YAML = """
app:
  name: "Test App"
  check_interval: 60
  expiry_notice_minutes: 30
excel:
  main_file: "test.xlsx"
  sheet_name: "Sheet1"
  columns:
    recipient: "收件人"
    status: "发送状态"
status_values:
  pending: "待发送"
  sent: "已发送"
  failed: "发送失败"
send_mode_values:
  repeat: "重复"
  once: "一次性"
templates_file: "templates.yaml"
outlook:
  account: ""
  save_to_sent: true
workdays:
  weekdays: [1, 2, 3, 4, 5]
  holidays: []
error_handling:
  on_error:
    stop_scheduler: true
    skip_and_continue: false
  mark_failed_as: "发送失败"
  notification:
    show_dialog: true
    sound_alert: true
logging:
  level: "INFO"
  file: "logs/test.log"
  max_size_mb: 10
  backup_count: 5
"""


@pytest.fixture
def config_file(tmp_path):
    f = tmp_path / "config.yaml"
    f.write_text(CONFIG_YAML, encoding='utf-8')
    return str(f)


def test_load_config_success(config_file):
    manager = ConfigManager(config_file)
    config = manager.load()

    assert config.app.name == "Test App"
    assert config.app.check_interval == 60
    assert config.excel.main_file == "test.xlsx"
    assert config.excel.sheet_name == "Sheet1"
    assert config.status_values["pending"] == "待发送"
    assert config.outlook.save_to_sent is True
    assert config.error_handling.stop_scheduler is True
    assert config.error_handling.mark_failed_as == "发送失败"
    assert config.logging.level == "INFO"
    assert config.logging.max_size_mb == 10


def test_load_config_workdays(config_file):
    manager = ConfigManager(config_file)
    config = manager.load()

    assert config.workdays.weekdays == [1, 2, 3, 4, 5]
    assert config.workdays.holidays == []


def test_save_and_reload(config_file, tmp_path):
    manager = ConfigManager(config_file)
    config = manager.load()

    save_path = str(tmp_path / "saved_config.yaml")
    save_manager = ConfigManager(save_path)
    save_manager.save(config)

    reload_manager = ConfigManager(save_path)
    reloaded = reload_manager.load()

    assert reloaded.app.name == config.app.name
    assert reloaded.excel.main_file == config.excel.main_file
    assert reloaded.status_values == config.status_values


def test_load_missing_file():
    manager = ConfigManager("nonexistent.yaml")
    with pytest.raises(FileNotFoundError):
        manager.load()
