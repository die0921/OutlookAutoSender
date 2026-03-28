from dataclasses import dataclass
from typing import Dict, List


@dataclass(frozen=True)
class AppConfig:
    """应用配置"""

    name: str
    version: str
    check_interval: int
    expiry_notice_minutes: int


@dataclass(frozen=True)
class ExcelConfig:
    """Excel 配置"""

    file_path: str
    columns: Dict[str, str]


@dataclass(frozen=True)
class OutlookConfig:
    """Outlook 配置"""

    auto_close_after_send: bool
    signature_enabled: bool
    importance: str
    read_receipt: bool
    delivery_receipt: bool


@dataclass(frozen=True)
class WorkdaysConfig:
    """工作日配置"""

    enabled: bool
    days: List[int]
    start_time: str
    end_time: str
    skip_holidays: bool
    holidays: List[str]


@dataclass(frozen=True)
class ErrorHandlingConfig:
    """错误处理配置"""

    max_retries: int
    retry_interval_seconds: int
    log_errors: bool
    notify_on_failure: bool


@dataclass(frozen=True)
class LoggingConfig:
    """日志配置"""

    level: str
    file_path: str
    max_bytes: int
    backup_count: int


@dataclass(frozen=True)
class Config:
    """完整配置"""

    app: AppConfig
    excel: ExcelConfig
    outlook: OutlookConfig
    workdays: WorkdaysConfig
    error_handling: ErrorHandlingConfig
    logging: LoggingConfig
    status_values: Dict[str, str]
    send_mode_values: Dict[str, str]
    templates_file: str
