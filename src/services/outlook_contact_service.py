import re
from typing import List, Optional


class OutlookContactService:
    """Outlook 联系人组服务：展开联系人组为邮箱列表"""

    EMAIL_PATTERN = re.compile(r'^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$')

    def __init__(self, outlook_app=None):
        """
        Args:
            outlook_app: win32com Outlook.Application 对象，None 时延迟初始化
        """
        self._outlook_app = outlook_app
        self._address_book = None

    def resolve_recipients(self, recipient_string: str) -> List[str]:
        """将收件人字符串展开为邮箱列表

        支持分号分隔的邮箱地址和联系人组名称混合：
        "销售组;user@test.com;VIP客户组"

        Returns:
            去重后的邮箱列表

        Raises:
            ValueError: 联系人组不存在时
        """
        if not recipient_string or not recipient_string.strip():
            raise ValueError("收件人不能为空")

        items = [item.strip() for item in recipient_string.split(';') if item.strip()]
        emails = []

        for item in items:
            if self._is_email(item):
                emails.append(item.lower())
            else:
                # 当作联系人组处理
                group_emails = self.get_distribution_list(item)
                emails.extend(group_emails)

        # 去重保序
        seen = set()
        result = []
        for email in emails:
            if email not in seen:
                seen.add(email)
                result.append(email)

        return result

    def get_distribution_list(self, group_name: str) -> List[str]:
        """获取联系人组的成员邮箱列表

        Raises:
            ValueError: 联系人组不存在
        """
        outlook = self._get_outlook()
        address_lists = outlook.Session.AddressLists

        for i in range(1, address_lists.Count + 1):
            addr_list = address_lists.Item(i)
            entries = addr_list.AddressEntries
            for j in range(1, entries.Count + 1):
                entry = entries.Item(j)
                if entry.Name == group_name:
                    return self._expand_entry(entry)

        raise ValueError(f"联系人组不存在: {group_name}")

    def validate_group_exists(self, group_name: str) -> bool:
        """检查联系人组是否存在"""
        try:
            self.get_distribution_list(group_name)
            return True
        except ValueError:
            return False

    def _get_outlook(self):
        """延迟初始化 Outlook COM 对象"""
        if self._outlook_app is None:
            import win32com.client
            self._outlook_app = win32com.client.Dispatch("Outlook.Application")
        return self._outlook_app

    def _expand_entry(self, entry) -> List[str]:
        """递归展开通讯组列表成员"""
        emails = []
        try:
            if entry.AddressEntryUserType == 1:  # Distribution list
                members = entry.Members
                for i in range(1, members.Count + 1):
                    emails.extend(self._expand_entry(members.Item(i)))
            else:
                email = getattr(entry, 'Address', None) or \
                        getattr(entry, 'GetExchangeUser', lambda: None)()
                if email and isinstance(email, str) and self._is_email(email):
                    emails.append(email.lower())
        except Exception:
            pass
        return emails

    def _is_email(self, text: str) -> bool:
        """判断字符串是否为邮箱地址"""
        return bool(self.EMAIL_PATTERN.match(text.strip()))
