from datetime import date, datetime, timedelta
from typing import List, Optional


class DateUtils:
    """日期时间工具类"""

    # 星期名称映射（中文 -> isoweekday: 1=周一, 7=周日）
    WEEKDAY_MAP = {
        '周一': 1, '星期一': 1, '一': 1,
        '周二': 2, '星期二': 2, '二': 2,
        '周三': 3, '星期三': 3, '三': 3,
        '周四': 4, '星期四': 4, '四': 4,
        '周五': 5, '星期五': 5, '五': 5,
        '周六': 6, '星期六': 6, '六': 6,
        '周日': 7, '星期日': 7, '日': 7, '周天': 7,
        'Monday': 1, 'Tuesday': 2, 'Wednesday': 3,
        'Thursday': 4, 'Friday': 5, 'Saturday': 6, 'Sunday': 7,
    }

    def parse_time(self, time_str: str) -> Optional[tuple]:
        """解析时间字符串，返回 (hour, minute) 元组

        支持格式: "9:00", "09:00", "9:30"
        """
        time_str = time_str.strip()
        try:
            parts = time_str.split(':')
            if len(parts) != 2:
                return None
            hour = int(parts[0])
            minute = int(parts[1])
            if 0 <= hour <= 23 and 0 <= minute <= 59:
                return (hour, minute)
            return None
        except (ValueError, AttributeError):
            return None

    def parse_weekday(self, weekday_str: str) -> Optional[int]:
        """解析星期字符串，返回 isoweekday (1=周一, 7=周日)

        支持: "周一", "星期一", "Monday" 等
        """
        return self.WEEKDAY_MAP.get(weekday_str.strip())

    def is_workday(self, d: date, workday_numbers: List[int],
                   holidays: List[str]) -> bool:
        """判断指定日期是否为工作日

        Args:
            d: 要判断的日期
            workday_numbers: 工作日列表 (isoweekday: 1=周一)
            holidays: 节假日列表 (YYYY-MM-DD 格式字符串)
        """
        date_str = d.strftime('%Y-%m-%d')
        if date_str in holidays:
            return False
        return d.isoweekday() in workday_numbers

    def next_workday(self, from_date: date, workday_numbers: List[int],
                     holidays: List[str]) -> date:
        """获取从指定日期起（含当天）的下一个工作日"""
        d = from_date
        for _ in range(365):  # 最多查找365天
            if self.is_workday(d, workday_numbers, holidays):
                return d
            d += timedelta(days=1)
        raise ValueError("找不到工作日（可能节假日配置有误）")

    def get_nth_workday_of_month(self, year: int, month: int, nth: int,
                                  workday_numbers: List[int],
                                  holidays: List[str]) -> Optional[date]:
        """获取某月第 N 个工作日

        Args:
            nth: 第几个工作日（从 1 开始）
        """
        d = date(year, month, 1)
        count = 0
        while d.month == month:
            if self.is_workday(d, workday_numbers, holidays):
                count += 1
                if count == nth:
                    return d
            d += timedelta(days=1)
        return None  # 该月没有第 nth 个工作日

    def get_last_workday_of_month(self, year: int, month: int,
                                   workday_numbers: List[int],
                                   holidays: List[str]) -> Optional[date]:
        """获取某月最后一个工作日"""
        # 找到下月第一天，往前推
        if month == 12:
            next_month = date(year + 1, 1, 1)
        else:
            next_month = date(year, month + 1, 1)

        d = next_month - timedelta(days=1)
        while d.month == month:
            if self.is_workday(d, workday_numbers, holidays):
                return d
            d -= timedelta(days=1)
        return None
