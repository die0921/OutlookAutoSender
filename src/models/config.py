from dataclasses import dataclass
from typing import List, Dict


@dataclass(frozen=True)
class AppConfig:
    """应用配置"""
    name: str
    check_interval: int
    expiry_notice_minutes: int


@dataclass(frozen=True)
class ExcelConfig:
    """Excel 配置"""
    main_file: str
    sheet_name: str
    columns: Dict[str, str]


@dataclass(frozen=True)
class OutlookConfig:
    """Outlook 配置"""
    account: str
    save_to_sent: bool


@dataclass(frozen=True)
class WorkdaysConfig:
    """工作日配置"""
    weekdays: List[int]
    holidays: List[str]


@dataclass(frozen=True)
class ErrorHandlingConfig:
    """错误处理配置"""
    stop_scheduler: bool
    skip_and_continue: bool
    mark_failed_as: str
    show_dialog: bool
    sound_alert: bool


@dataclass(frozen=True)
class LoggingConfig:
    """日志配置"""
    level: str
    file: str
    max_size_mb: int
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
