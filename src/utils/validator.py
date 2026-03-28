import os
import re
from typing import List, Tuple


class Validator:
    """输入验证工具类"""

    # 简单邮箱正则（支持常见格式）
    EMAIL_PATTERN = re.compile(r'^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$')

    def validate_email(self, email: str) -> Tuple[bool, str]:
        """验证单个邮箱地址格式

        Returns:
            (is_valid, error_message)
        """
        email = email.strip()
        if not email:
            return False, "邮箱地址不能为空"
        if not self.EMAIL_PATTERN.match(email):
            return False, f"邮箱格式不正确: {email}"
        return True, ""

    def validate_emails(self, emails_str: str) -> Tuple[bool, str]:
        """验证分号分隔的多个邮箱地址

        Returns:
            (is_valid, error_message)
        """
        if not emails_str or not emails_str.strip():
            return False, "收件人不能为空"

        emails = [e.strip() for e in emails_str.split(';') if e.strip()]
        if not emails:
            return False, "收件人不能为空"

        for email in emails:
            # 跳过可能是联系人组名称的（不含 @）
            if '@' not in email:
                continue
            valid, msg = self.validate_email(email)
            if not valid:
                return False, msg

        return True, ""

    def validate_file_exists(self, file_path: str) -> Tuple[bool, str]:
        """验证文件是否存在

        Returns:
            (is_valid, error_message)
        """
        if not file_path or not file_path.strip():
            return False, "文件路径不能为空"
        if not os.path.isfile(file_path.strip()):
            return False, f"文件不存在: {file_path}"
        return True, ""

    def validate_required_fields(self, fields: dict) -> Tuple[bool, List[str]]:
        """验证必填字段不为空

        Args:
            fields: {field_name: value} 字典

        Returns:
            (is_valid, list_of_missing_field_names)
        """
        missing = [name for name, val in fields.items()
                   if val is None or str(val).strip() == '']
        return len(missing) == 0, missing
