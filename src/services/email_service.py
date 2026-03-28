from dataclasses import dataclass
from typing import List, Optional, Tuple
from src.models.email_task import EmailTask
from src.services.template_engine import TemplateEngine
from src.services.outlook_contact_service import OutlookContactService


@dataclass(frozen=True)
class EmailPreview:
    """邮件预览数据"""
    recipients: List[str]
    cc_list: List[str]
    subject: str
    body: str
    attachments: List[str]
    planned_time: Optional[str]


class EmailService:
    """通过 Outlook COM 发送邮件"""

    def __init__(self, template_engine: TemplateEngine,
                 contact_service: OutlookContactService,
                 outlook_app=None):
        self.template_engine = template_engine
        self.contact_service = contact_service
        self._outlook_app = outlook_app

    def send_email(self, task: EmailTask) -> Tuple[bool, str]:
        """发送邮件

        Returns:
            (success, error_message)
        """
        try:
            # 1. 渲染模板或使用任务主题/正文
            subject, body = self._prepare_content(task)

            # 2. 解析收件人
            recipients = self.contact_service.resolve_recipients(task.recipient)
            cc_list = []
            if task.cc and task.cc.strip():
                cc_list = self.contact_service.resolve_recipients(task.cc)

            # 3. 解析附件
            attachments = self._parse_attachments(task.attachments)

            # 4. 调用 Outlook 发送
            self._send_via_outlook(recipients, cc_list, subject, body, attachments)
            return True, ""

        except Exception as e:
            return False, str(e)

    def preview_email(self, task: EmailTask) -> EmailPreview:
        """生成邮件预览（不发送）"""
        subject, body = self._prepare_content(task)

        try:
            recipients = self.contact_service.resolve_recipients(task.recipient)
        except Exception:
            recipients = [task.recipient]

        cc_list = []
        if task.cc and task.cc.strip():
            try:
                cc_list = self.contact_service.resolve_recipients(task.cc)
            except Exception:
                cc_list = [task.cc]

        attachments = self._parse_attachments(task.attachments)
        planned_str = (task.planned_time.strftime('%Y-%m-%d %H:%M')
                       if task.planned_time else None)

        return EmailPreview(
            recipients=recipients,
            cc_list=cc_list,
            subject=subject,
            body=body,
            attachments=attachments,
            planned_time=planned_str,
        )

    def _prepare_content(self, task: EmailTask) -> Tuple[str, str]:
        """准备邮件主题和正文"""
        if task.template_name and task.template_name.strip():
            return self.template_engine.render(task.template_name, task.variables)
        return task.subject, ''

    def _parse_attachments(self, attachments_str: str) -> List[str]:
        """解析附件路径字符串（分号分隔）"""
        if not attachments_str or not attachments_str.strip():
            return []
        return [p.strip() for p in attachments_str.split(';') if p.strip()]

    def _get_outlook(self):
        """延迟初始化 Outlook COM"""
        if self._outlook_app is None:
            import win32com.client
            self._outlook_app = win32com.client.Dispatch("Outlook.Application")
        return self._outlook_app

    def _send_via_outlook(self, recipients: List[str], cc_list: List[str],
                          subject: str, body: str, attachments: List[str]) -> None:
        """通过 Outlook COM API 发送邮件"""
        outlook = self._get_outlook()
        mail = outlook.CreateItem(0)  # olMailItem = 0

        mail.Subject = subject
        mail.HTMLBody = body if '<html>' in body.lower() else f'<html><body>{body}</body></html>'

        for recipient in recipients:
            mail.Recipients.Add(recipient)
        for cc in cc_list:
            recipient = mail.Recipients.Add(cc)
            recipient.Type = 2  # olCC = 2

        for attachment in attachments:
            mail.Attachments.Add(attachment)

        mail.Recipients.ResolveAll()
        mail.Send()
