from dataclasses import dataclass
from typing import List, Dict, Optional


@dataclass(frozen=True)
class DataSource:
    """数据源配置"""
    name: str
    file: str


@dataclass(frozen=True)
class BodyPart:
    """正文部分"""
    type: str  # "text" or "excel_range"
    content: Optional[str] = None
    source: Optional[str] = None
    sheet: Optional[str] = None
    range: Optional[str] = None
    format: Optional[str] = None
    include_header: Optional[bool] = None


@dataclass(frozen=True)
class Template:
    """邮件模板"""
    name: str
    subject: str
    body_type: str
    body_parts: List[BodyPart]  # Note: List in frozen dataclass is allowed
    data_sources: Dict[str, DataSource]
