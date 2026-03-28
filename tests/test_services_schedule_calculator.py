import pytest
from datetime import date, datetime
from src.services.schedule_calculator import ScheduleCalculator

WORKDAYS = [1, 2, 3, 4, 5]
HOLIDAYS = ['2026-01-01']


@pytest.fixture
def calc():
    return ScheduleCalculator(WORKDAYS, HOLIDAYS)


# --- calculate_next_time: 按天 ---

def test_daily_future_today(calc):
    """今天8点查询，下次发送时间应为今天9点"""
    ref = datetime(2026, 3, 28, 8, 0)
    result = calc.calculate_next_time('按天', '9:00', ref)
    assert result == datetime(2026, 3, 28, 9, 0)


def test_daily_past_today(calc):
    """今天10点查询，下次发送时间应为明天9点"""
    ref = datetime(2026, 3, 28, 10, 0)
    result = calc.calculate_next_time('按天', '9:00', ref)
    assert result == datetime(2026, 3, 29, 9, 0)


# --- calculate_next_time: 按周 ---

def test_weekly_next_monday(calc):
    """周六查询，下次周一9点"""
    ref = datetime(2026, 3, 28, 10, 0)  # Saturday
    result = calc.calculate_next_time('按周', '周一,9:00', ref)
    assert result == datetime(2026, 3, 30, 9, 0)  # Next Monday


def test_weekly_same_day_past(calc):
    """周一10点查询下次周一，应返回下周一"""
    ref = datetime(2026, 3, 23, 10, 0)  # Monday
    result = calc.calculate_next_time('按周', '周一,9:00', ref)
    assert result == datetime(2026, 3, 30, 9, 0)  # Next Monday


# --- calculate_next_time: 按月-固定 ---

def test_monthly_fixed_future_this_month(calc):
    """3月5日查询，下次1日为4月1日"""
    ref = datetime(2026, 3, 5, 10, 0)
    result = calc.calculate_next_time('按月-固定', '1日,9:00', ref)
    assert result == datetime(2026, 4, 1, 9, 0)


def test_monthly_fixed_before_day(calc):
    """3月1日8点查询，当月1日9点仍未到"""
    ref = datetime(2026, 3, 1, 8, 0)
    result = calc.calculate_next_time('按月-固定', '1日,9:00', ref)
    assert result == datetime(2026, 3, 1, 9, 0)


# --- calculate_next_time: 按月-工作日 ---

def test_monthly_workday_third(calc):
    """查询本月第3个工作日"""
    ref = datetime(2026, 3, 1, 8, 0)
    result = calc.calculate_next_time('按月-工作日', '第3个工作日,9:00', ref)
    assert result is not None
    assert result.day == 4  # 2026-03: 1=Sun, 2=Mon(1st), 3=Tue(2nd), 4=Wed(3rd)


def test_monthly_workday_last(calc):
    """查询本月最后工作日"""
    ref = datetime(2026, 3, 1, 8, 0)
    result = calc.calculate_next_time('按月-工作日', '最后工作日,16:00', ref)
    assert result is not None
    assert result.month == 3
    assert result.day == 31  # 2026-03-31 is Tuesday


# --- calculate_next_time: 按日期列表 ---

def test_date_list_next_upcoming(calc):
    ref = datetime(2026, 3, 27, 10, 0)
    result = calc.calculate_next_time('按日期列表', '2026-03-28,2026-04-10;9:00', ref)
    assert result == datetime(2026, 3, 28, 9, 0)


def test_date_list_all_past(calc):
    ref = datetime(2026, 5, 1, 10, 0)
    result = calc.calculate_next_time('按日期列表', '2026-03-28,2026-04-10;9:00', ref)
    assert result is None


# --- should_reset_status ---

def test_no_reset_once_mode(calc):
    """一次性模式永不重置"""
    assert calc.should_reset_status('一次性', '按天', '9:00',
                                     date(2026, 3, 27), date(2026, 3, 28)) is False


def test_no_reset_never_sent(calc):
    """从未发送不触发重置（last_sent_date=None）"""
    assert calc.should_reset_status('重复', '按天', '9:00',
                                     None, date(2026, 3, 28)) is False


def test_reset_daily_next_day(calc):
    """按天：昨天发送了，今天应重置"""
    assert calc.should_reset_status('重复', '按天', '9:00',
                                     date(2026, 3, 27), date(2026, 3, 28)) is True


def test_no_reset_daily_same_day(calc):
    """按天：今天已发送，不重置"""
    assert calc.should_reset_status('重复', '按天', '9:00',
                                     date(2026, 3, 28), date(2026, 3, 28)) is False


def test_reset_monthly_new_month(calc):
    """按月：上月发送了，本月应重置"""
    assert calc.should_reset_status('重复', '按月-固定', '1日,9:00',
                                     date(2026, 2, 1), date(2026, 3, 1)) is True


def test_no_reset_monthly_same_month(calc):
    """按月：本月已发送，不重置"""
    assert calc.should_reset_status('重复', '按月-固定', '1日,9:00',
                                     date(2026, 3, 1), date(2026, 3, 15)) is False
