from dataclasses import dataclass
from datetime import datetime, date
from typing import Optional, Dict, Any


@dataclass(frozen=True)
class EmailTask:
    """邮件任务数据模型"""
    row_index: int
    recipient: str
    cc: str
    subject: str
    attachments: str
    status: str
    send_mode: str
    template_name: str
    schedule_type: str
    schedule_params: str
    planned_time: Optional[datetime]
    last_sent_date: Optional[date]
    sent_time: Optional[datetime]
    variables: Dict[str, Any]

    def is_pending(self) -> bool:
        """是否待发送（发送状态 == '待发送'）"""
        return self.status == "待发送"

    def is_sent(self) -> bool:
        """是否已发送（发送状态 == '已发送'）"""
        return self.status == "已发送"

    def is_failed(self) -> bool:
        """是否发送失败（发送状态 == '发送失败'）"""
        return self.status == "发送失败"

    def is_repeat_mode(self) -> bool:
        """是否重复模式（发送模式 == '重复'）"""
        return self.send_mode == "重复"
