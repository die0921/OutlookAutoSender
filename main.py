import sys
import os
import logging
from PyQt5.QtWidgets import QApplication, QMessageBox

# Ensure src is importable
sys.path.insert(0, os.path.dirname(__file__))

from src.data.config_manager import ConfigManager
from src.data.log_manager import LogManager
from src.data.excel_reader import ExcelReader
from src.data.excel_writer import ExcelWriter
from src.data.template_loader import TemplateLoader
from src.models.config import Config
from src.services.schedule_calculator import ScheduleCalculator
from src.services.template_engine import TemplateEngine
from src.services.outlook_contact_service import OutlookContactService
from src.services.email_service import EmailService
from src.services.scheduler_service import SchedulerService
from src.ui.main_window import MainWindow

CONFIG_FILE = os.path.join(os.path.dirname(__file__), 'config.yaml')
TEMPLATES_FILE = os.path.join(os.path.dirname(__file__), 'templates.yaml')


def load_config() -> Config:
    """加载配置文件"""
    manager = ConfigManager(CONFIG_FILE)
    return manager.load()


def build_scheduler(config: Config) -> SchedulerService:
    """构建调度器及其依赖"""
    excel_reader = ExcelReader(config.excel)
    excel_writer = ExcelWriter(config.excel)
    template_loader = TemplateLoader(TEMPLATES_FILE)
    contact_service = OutlookContactService()
    template_engine = TemplateEngine(template_loader)
    email_service = EmailService(template_engine, contact_service)
    calc = ScheduleCalculator(
        config.workdays.weekdays,
        config.workdays.holidays
    )
    return SchedulerService(
        config=config,
        excel_reader=excel_reader,
        excel_writer=excel_writer,
        email_service=email_service,
        schedule_calculator=calc,
    )


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("Outlook 自动发送助手")

    # 加载配置
    try:
        config = load_config()
    except FileNotFoundError:
        QMessageBox.critical(None, "启动失败", f"配置文件不存在: {CONFIG_FILE}")
        return 1
    except Exception as e:
        QMessageBox.critical(None, "启动失败", f"配置文件加载失败: {e}")
        return 1

    # 初始化日志
    log_manager = LogManager(config.logging)
    logger = log_manager.get_logger()
    logger.info("应用启动")

    # 构建调度器
    scheduler = build_scheduler(config)

    # 创建主窗口
    window = MainWindow()
    window.set_file_path(config.excel.main_file)

    # 连接信号
    def on_start():
        scheduler.start()
        window.set_running_state(True)
        window.append_log("调度器已启动")

    def on_stop():
        scheduler.stop()
        window.set_running_state(False)
        window.append_log("调度器已停止")

    def on_manual():
        scheduler.trigger_manual()
        window.append_log("手动触发发送检查")

    def on_refresh():
        try:
            tasks = ExcelReader(config.excel).read_tasks(config.excel.main_file)
            window.update_tasks([t for t in tasks if t.status == '待发送'])
        except Exception as e:
            window.append_log(f"刷新失败: {e}")

    window.start_requested.connect(on_start)
    window.stop_requested.connect(on_stop)
    window.manual_send_requested.connect(on_manual)
    window.refresh_requested.connect(on_refresh)

    window.show()
    return app.exec_()


if __name__ == '__main__':
    sys.exit(main())
