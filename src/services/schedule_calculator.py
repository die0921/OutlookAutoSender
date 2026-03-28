from datetime import date, datetime, timedelta
from typing import List, Optional
from src.utils.date_utils import DateUtils


class ScheduleCalculator:
    """5种定时方式的发送时间计算器"""

    def __init__(self, workday_numbers: List[int], holidays: List[str]):
        self.workday_numbers = workday_numbers
        self.holidays = holidays
        self._date_utils = DateUtils()

    def calculate_next_time(self, schedule_type: str, schedule_params: str,
                             reference_dt: datetime) -> Optional[datetime]:
        """计算下次发送时间

        Args:
            schedule_type: 发送方式（按天/按周/按月-固定/按月-工作日/按日期列表）
            schedule_params: 发送参数
            reference_dt: 参考时间（通常为当前时间）

        Returns:
            下次发送的 datetime，或 None（参数无效/无合适时间）
        """
        if schedule_type == '按天':
            return self._calc_daily(schedule_params, reference_dt)
        elif schedule_type == '按周':
            return self._calc_weekly(schedule_params, reference_dt)
        elif schedule_type == '按月-固定':
            return self._calc_monthly_fixed(schedule_params, reference_dt)
        elif schedule_type == '按月-工作日':
            return self._calc_monthly_workday(schedule_params, reference_dt)
        elif schedule_type == '按日期列表':
            return self._calc_date_list(schedule_params, reference_dt)
        return None

    def should_reset_status(self, send_mode: str, schedule_type: str,
                             schedule_params: str, last_sent_date: Optional[date],
                             current_date: date) -> bool:
        """判断重复任务是否应重置为待发送

        Args:
            send_mode: 发送模式（重复/一次性）
            schedule_type: 发送方式
            schedule_params: 发送参数
            last_sent_date: 最后发送日期（None 表示从未发送）
            current_date: 当前日期

        Returns:
            True 表示应重置为待发送
        """
        if send_mode != '重复':
            return False
        if last_sent_date is None:
            return False  # 从未发送，保持待发送状态（不需要重置）

        if schedule_type == '按天':
            return last_sent_date < current_date
        elif schedule_type == '按周':
            return self._should_reset_weekly(schedule_params, last_sent_date, current_date)
        elif schedule_type in ('按月-固定', '按月-工作日'):
            # 不同月份则重置
            return (last_sent_date.year != current_date.year or
                    last_sent_date.month != current_date.month)
        elif schedule_type == '按日期列表':
            return self._should_reset_date_list(schedule_params, last_sent_date, current_date)
        return False

    # --- Private calculation methods ---

    def _calc_daily(self, params: str, ref: datetime) -> Optional[datetime]:
        """按天：params = "9:00" """
        t = self._date_utils.parse_time(params)
        if t is None:
            return None
        hour, minute = t
        candidate = ref.replace(hour=hour, minute=minute, second=0, microsecond=0)
        if candidate <= ref:
            candidate += timedelta(days=1)
        return candidate

    def _calc_weekly(self, params: str, ref: datetime) -> Optional[datetime]:
        """按周：params = "周一,9:00" """
        parts = params.split(',')
        if len(parts) != 2:
            return None
        weekday = self._date_utils.parse_weekday(parts[0].strip())
        t = self._date_utils.parse_time(parts[1].strip())
        if weekday is None or t is None:
            return None
        hour, minute = t

        # 找下一个目标星期几
        d = ref.date()
        for i in range(8):
            candidate_date = d + timedelta(days=i)
            candidate = datetime(candidate_date.year, candidate_date.month,
                                  candidate_date.day, hour, minute)
            if candidate_date.isoweekday() == weekday and candidate > ref:
                return candidate
        return None

    def _calc_monthly_fixed(self, params: str, ref: datetime) -> Optional[datetime]:
        """按月-固定：params = "1日,9:00" """
        parts = params.split(',')
        if len(parts) != 2:
            return None
        day_str = parts[0].strip().replace('日', '').replace('号', '')
        t = self._date_utils.parse_time(parts[1].strip())
        try:
            day = int(day_str)
        except ValueError:
            return None
        if t is None:
            return None
        hour, minute = t

        # 当月
        try:
            candidate = ref.replace(day=day, hour=hour, minute=minute,
                                     second=0, microsecond=0)
            if candidate > ref:
                return candidate
        except ValueError:
            pass  # 当月没有该日

        # 下月
        year, month = ref.year, ref.month
        month += 1
        if month > 12:
            month = 1
            year += 1
        try:
            return datetime(year, month, day, hour, minute)
        except ValueError:
            return None

    def _calc_monthly_workday(self, params: str, ref: datetime) -> Optional[datetime]:
        """按月-工作日：params = "第3个工作日,9:00" 或 "最后工作日,16:00" """
        parts = params.split(',')
        if len(parts) != 2:
            return None
        desc = parts[0].strip()
        t = self._date_utils.parse_time(parts[1].strip())
        if t is None:
            return None
        hour, minute = t

        def get_target_date(year: int, month: int) -> Optional[date]:
            if desc == '最后工作日':
                return self._date_utils.get_last_workday_of_month(
                    year, month, self.workday_numbers, self.holidays)
            elif desc.startswith('第') and '个工作日' in desc:
                nth_str = desc.replace('第', '').replace('个工作日', '').strip()
                try:
                    nth = int(nth_str)
                except ValueError:
                    return None
                return self._date_utils.get_nth_workday_of_month(
                    year, month, nth, self.workday_numbers, self.holidays)
            return None

        # 当月
        target = get_target_date(ref.year, ref.month)
        if target:
            candidate = datetime(target.year, target.month, target.day, hour, minute)
            if candidate > ref:
                return candidate

        # 下月
        year, month = ref.year, ref.month + 1
        if month > 12:
            month = 1
            year += 1
        target = get_target_date(year, month)
        if target:
            return datetime(target.year, target.month, target.day, hour, minute)
        return None

    def _calc_date_list(self, params: str, ref: datetime) -> Optional[datetime]:
        """按日期列表：params = "2026-03-28,2026-04-10;9:00" """
        if ';' not in params:
            return None
        dates_str, time_str = params.rsplit(';', 1)
        t = self._date_utils.parse_time(time_str.strip())
        if t is None:
            return None
        hour, minute = t

        date_strings = [d.strip() for d in dates_str.split(',') if d.strip()]
        candidates = []
        for ds in date_strings:
            try:
                d = date.fromisoformat(ds)
                candidate = datetime(d.year, d.month, d.day, hour, minute)
                if candidate > ref:
                    candidates.append(candidate)
            except ValueError:
                continue

        return min(candidates) if candidates else None

    def _should_reset_weekly(self, params: str, last_sent: date, current: date) -> bool:
        """判断按周任务是否需要重置"""
        parts = params.split(',')
        if len(parts) != 2:
            return False
        weekday = self._date_utils.parse_weekday(parts[0].strip())
        if weekday is None:
            return False
        # 找上次发送后的下一个目标星期
        d = last_sent + timedelta(days=1)
        for _ in range(8):
            if d.isoweekday() == weekday:
                return current >= d
            d += timedelta(days=1)
        return False

    def _should_reset_date_list(self, params: str, last_sent: date, current: date) -> bool:
        """判断按日期列表任务是否需要重置"""
        if ';' not in params:
            return False
        dates_str, _ = params.rsplit(';', 1)
        for ds in dates_str.split(','):
            ds = ds.strip()
            try:
                d = date.fromisoformat(ds)
                if last_sent < d <= current:
                    return True
            except ValueError:
                continue
        return False
