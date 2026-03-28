import pytest
from datetime import date
from src.utils.date_utils import DateUtils


WORKDAYS = [1, 2, 3, 4, 5]  # 周一到周五
HOLIDAYS = ['2026-01-01', '2026-10-01']


@pytest.fixture
def date_utils():
    return DateUtils()


# --- parse_time ---

def test_parse_time_valid(date_utils):
    assert date_utils.parse_time("9:00") == (9, 0)
    assert date_utils.parse_time("09:30") == (9, 30)
    assert date_utils.parse_time("18:00") == (18, 0)


def test_parse_time_invalid(date_utils):
    assert date_utils.parse_time("25:00") is None
    assert date_utils.parse_time("abc") is None
    assert date_utils.parse_time("9") is None


# --- parse_weekday ---

def test_parse_weekday_chinese(date_utils):
    assert date_utils.parse_weekday("周一") == 1
    assert date_utils.parse_weekday("周五") == 5
    assert date_utils.parse_weekday("周日") == 7


def test_parse_weekday_english(date_utils):
    assert date_utils.parse_weekday("Monday") == 1
    assert date_utils.parse_weekday("Friday") == 5


def test_parse_weekday_invalid(date_utils):
    assert date_utils.parse_weekday("周八") is None


# --- is_workday ---

def test_is_workday_monday(date_utils):
    monday = date(2026, 3, 23)  # 周一
    assert date_utils.is_workday(monday, WORKDAYS, []) is True


def test_is_workday_saturday(date_utils):
    saturday = date(2026, 3, 28)  # 周六
    assert date_utils.is_workday(saturday, WORKDAYS, []) is False


def test_is_workday_holiday(date_utils):
    new_year = date(2026, 1, 1)  # 节假日
    assert date_utils.is_workday(new_year, WORKDAYS, HOLIDAYS) is False


# --- next_workday ---

def test_next_workday_from_monday(date_utils):
    monday = date(2026, 3, 23)
    assert date_utils.next_workday(monday, WORKDAYS, []) == monday


def test_next_workday_from_saturday(date_utils):
    saturday = date(2026, 3, 28)  # 周六 -> 下周一
    result = date_utils.next_workday(saturday, WORKDAYS, [])
    assert result == date(2026, 3, 30)


# --- get_nth_workday_of_month ---

def test_get_first_workday_of_march_2026(date_utils):
    # 2026-03-01 is Sunday, first workday is Monday 2026-03-02
    result = date_utils.get_nth_workday_of_month(2026, 3, 1, WORKDAYS, [])
    assert result == date(2026, 3, 2)


def test_get_third_workday_of_march_2026(date_utils):
    result = date_utils.get_nth_workday_of_month(2026, 3, 3, WORKDAYS, [])
    assert result == date(2026, 3, 4)


# --- get_last_workday_of_month ---

def test_get_last_workday_of_march_2026(date_utils):
    # 2026-03-31 is Tuesday, so last workday = 2026-03-31
    result = date_utils.get_last_workday_of_month(2026, 3, WORKDAYS, [])
    assert result == date(2026, 3, 31)


def test_get_last_workday_of_april_2026(date_utils):
    # 2026-04-30 is Thursday, so last workday = 2026-04-30
    result = date_utils.get_last_workday_of_month(2026, 4, WORKDAYS, [])
    assert result == date(2026, 4, 30)
