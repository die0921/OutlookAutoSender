import logging
from datetime import datetime, timedelta
from typing import Callable, Optional
from dataclasses import dataclass

from src.models.config import Config
from src.data.excel_reader import ExcelReader
from src.data.excel_writer import ExcelWriter
from src.services.email_service import EmailService
from src.services.schedule_calculator import ScheduleCalculator


@dataclass(frozen=True)
class SchedulerStatus:
    """调度器状态"""
    running: bool
    last_check: Optional[datetime]
    error_message: Optional[str]


class SchedulerService:
    """定时循环调度服务：读取→重置→计算→筛选→渲染→发送→更新"""

    def __init__(self, config: Config, excel_reader: ExcelReader,
                 excel_writer: ExcelWriter, email_service: EmailService,
                 schedule_calculator: ScheduleCalculator,
                 on_error: Optional[Callable[[str, Exception], None]] = None):
        self.config = config
        self.excel_reader = excel_reader
        self.excel_writer = excel_writer
        self.email_service = email_service
        self.schedule_calculator = schedule_calculator
        self.on_error = on_error

        self._running = False
        self._last_check: Optional[datetime] = None
        self._error_message: Optional[str] = None
        self._logger = logging.getLogger(__name__)

    def start(self) -> None:
        """启动调度器（立即执行一次检查）"""
        self._running = True
        self._error_message = None
        self._logger.info("调度器已启动")
        self._check_and_send()

    def stop(self) -> None:
        """停止调度器"""
        self._running = False
        self._logger.info("调度器已停止")

    def trigger_manual(self) -> None:
        """手动触发一次检查发送"""
        self._logger.info("手动触发发送检查")
        self._check_and_send()

    def get_status(self) -> SchedulerStatus:
        """获取调度器状态"""
        return SchedulerStatus(
            running=self._running,
            last_check=self._last_check,
            error_message=self._error_message,
        )

    def _check_and_send(self) -> None:
        """核心检查逻辑：读取→重置→计算→筛选→发送→更新"""
        now = datetime.now()
        self._last_check = now
        file_path = self.config.excel.main_file

        try:
            # 1. 读取所有任务
            tasks = self.excel_reader.read_tasks(file_path)

            # 2. 状态预处理：重复任务状态重置
            for task in tasks:
                if (task.send_mode == self.config.send_mode_values.get('repeat', '重复') and
                        task.status == self.config.status_values.get('sent', '已发送')):
                    if self.schedule_calculator.should_reset_status(
                            task.send_mode, task.schedule_type, task.schedule_params,
                            task.last_sent_date, now.date()):
                        self.excel_writer.reset_to_pending(file_path, task.row_index)
                        self._logger.info(f"重置任务状态: 行 {task.row_index}")

            # 3. 重新读取（状态可能已更新）
            tasks = self.excel_reader.read_tasks(file_path)

            # 4. 计算计划发送时间（待发送且无计划时间的任务）
            pending_status = self.config.status_values.get('pending', '待发送')
            for task in tasks:
                if task.status == pending_status and task.planned_time is None:
                    next_time = self.schedule_calculator.calculate_next_time(
                        task.schedule_type, task.schedule_params, now)
                    if next_time:
                        self.excel_writer.update_planned_time(
                            file_path, task.row_index, next_time)

            # 5. 再次读取，筛选到期任务
            tasks = self.excel_reader.read_tasks(file_path)
            tolerance = timedelta(seconds=30)
            due_tasks = [
                t for t in tasks
                if t.status == pending_status
                and t.planned_time is not None
                and t.planned_time <= now + tolerance
            ]

            # 6. 发送
            for task in due_tasks:
                if not self._running:
                    break
                success, error_msg = self.email_service.send_email(task)

                if success:
                    self.excel_writer.update_sent(
                        file_path, task.row_index,
                        self.config.status_values.get('sent', '已发送'),
                        now, now.date()
                    )
                    self._logger.info(f"发送成功: 行 {task.row_index} -> {task.recipient}")
                else:
                    failed_status = self.config.status_values.get('failed', '发送失败')
                    self.excel_writer.update_status(file_path, task.row_index, failed_status)
                    self._error_message = error_msg
                    self._running = False
                    self._logger.error(f"发送失败: 行 {task.row_index}, 错误: {error_msg}")

                    if self.on_error:
                        self.on_error(error_msg, Exception(error_msg))
                    return  # 停止处理后续任务

        except Exception as e:
            self._error_message = str(e)
            self._running = False
            self._logger.error(f"调度器异常: {e}")
            if self.on_error:
                self.on_error(str(e), e)
