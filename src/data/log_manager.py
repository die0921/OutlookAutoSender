import logging
import os
from logging.handlers import RotatingFileHandler
from src.models.config import LoggingConfig


class LogManager:
    """日志管理器，支持文件轮转和双输出（文件+控制台）"""

    def __init__(self, config: LoggingConfig):
        self.config = config
        self._logger = None

    def get_logger(self, name: str = 'outlook_sender') -> logging.Logger:
        """获取配置好的 Logger 实例"""
        if self._logger is not None:
            return self._logger

        logger = logging.getLogger(name)
        logger.setLevel(getattr(logging, self.config.level.upper(), logging.INFO))

        # 避免重复添加 handler
        if not logger.handlers:
            formatter = logging.Formatter(
                '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
            )

            # 文件 handler（带轮转）
            log_dir = os.path.dirname(self.config.file)
            if log_dir:
                os.makedirs(log_dir, exist_ok=True)

            file_handler = RotatingFileHandler(
                self.config.file,
                maxBytes=self.config.max_size_mb * 1024 * 1024,
                backupCount=self.config.backup_count,
                encoding='utf-8',
            )
            file_handler.setFormatter(formatter)
            logger.addHandler(file_handler)

            # 控制台 handler
            console_handler = logging.StreamHandler()
            console_handler.setFormatter(formatter)
            logger.addHandler(console_handler)

        self._logger = logger
        return logger

    def close(self) -> None:
        """关闭所有 handler，释放文件句柄"""
        if self._logger:
            for handler in self._logger.handlers[:]:
                handler.close()
                self._logger.removeHandler(handler)
            self._logger = None
