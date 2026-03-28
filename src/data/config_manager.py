import yaml
from typing import Dict, Any

from src.models.config import (
    Config,
    AppConfig,
    ExcelConfig,
    OutlookConfig,
    WorkdaysConfig,
    ErrorHandlingConfig,
    LoggingConfig,
)


class ConfigManager:
    """配置文件管理器"""

    def __init__(self, config_file: str) -> None:
        self.config_file = config_file

    def load(self) -> Config:
        """加载配置文件"""
        with open(self.config_file, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f)

        return Config(
            app=AppConfig(**data["app"]),
            excel=ExcelConfig(**data["excel"]),
            outlook=OutlookConfig(**data["outlook"]),
            workdays=WorkdaysConfig(**data["workdays"]),
            error_handling=ErrorHandlingConfig(
                stop_scheduler=data["error_handling"]["on_error"]["stop_scheduler"],
                skip_and_continue=data["error_handling"]["on_error"][
                    "skip_and_continue"
                ],
                mark_failed_as=data["error_handling"]["mark_failed_as"],
                show_dialog=data["error_handling"]["notification"]["show_dialog"],
                sound_alert=data["error_handling"]["notification"]["sound_alert"],
            ),
            logging=LoggingConfig(**data["logging"]),
            status_values=data["status_values"],
            send_mode_values=data["send_mode_values"],
            templates_file=data["templates_file"],
        )

    def save(self, config: Config) -> None:
        """保存配置文件"""
        data = {
            "app": {
                "name": config.app.name,
                "check_interval": config.app.check_interval,
                "expiry_notice_minutes": config.app.expiry_notice_minutes,
            },
            "excel": {
                "main_file": config.excel.main_file,
                "sheet_name": config.excel.sheet_name,
                "columns": config.excel.columns,
            },
            "outlook": {
                "account": config.outlook.account,
                "save_to_sent": config.outlook.save_to_sent,
            },
            "workdays": {
                "weekdays": config.workdays.weekdays,
                "holidays": config.workdays.holidays,
            },
            "error_handling": {
                "on_error": {
                    "stop_scheduler": config.error_handling.stop_scheduler,
                    "skip_and_continue": config.error_handling.skip_and_continue,
                },
                "mark_failed_as": config.error_handling.mark_failed_as,
                "notification": {
                    "show_dialog": config.error_handling.show_dialog,
                    "sound_alert": config.error_handling.sound_alert,
                },
            },
            "logging": {
                "level": config.logging.level,
                "file": config.logging.file,
                "max_size_mb": config.logging.max_size_mb,
                "backup_count": config.logging.backup_count,
            },
            "status_values": config.status_values,
            "send_mode_values": config.send_mode_values,
            "templates_file": config.templates_file,
        }

        with open(self.config_file, "w", encoding="utf-8") as f:
            yaml.dump(data, f, allow_unicode=True, default_flow_style=False)
