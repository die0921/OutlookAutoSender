import pytest
import os
from src.utils.validator import Validator


@pytest.fixture
def validator():
    return Validator()


# --- validate_email ---

def test_valid_email(validator):
    ok, msg = validator.validate_email("user@example.com")
    assert ok is True
    assert msg == ""


def test_invalid_email_no_at(validator):
    ok, msg = validator.validate_email("userexample.com")
    assert ok is False
    assert "格式不正确" in msg


def test_invalid_email_empty(validator):
    ok, msg = validator.validate_email("")
    assert ok is False
    assert "不能为空" in msg


def test_valid_email_with_dots(validator):
    ok, msg = validator.validate_email("user.name+tag@sub.example.co.uk")
    assert ok is True


# --- validate_emails ---

def test_validate_emails_single(validator):
    ok, msg = validator.validate_emails("user@example.com")
    assert ok is True


def test_validate_emails_multiple(validator):
    ok, msg = validator.validate_emails("user1@a.com;user2@b.com")
    assert ok is True


def test_validate_emails_with_group(validator):
    # 联系人组名称（不含@）应被跳过不验证
    ok, msg = validator.validate_emails("销售组;user@example.com")
    assert ok is True


def test_validate_emails_empty(validator):
    ok, msg = validator.validate_emails("")
    assert ok is False


def test_validate_emails_invalid(validator):
    ok, msg = validator.validate_emails("valid@a.com;invalid-email")
    # invalid-email has no @, treated as group name, should pass
    ok2, _ = validator.validate_emails("valid@a.com;bad@")
    assert ok2 is False


# --- validate_file_exists ---

def test_validate_file_exists_valid(validator, tmp_path):
    f = tmp_path / "test.txt"
    f.write_text("hello")
    ok, msg = validator.validate_file_exists(str(f))
    assert ok is True


def test_validate_file_not_found(validator):
    ok, msg = validator.validate_file_exists("/nonexistent/path/file.txt")
    assert ok is False
    assert "不存在" in msg


def test_validate_file_empty_path(validator):
    ok, msg = validator.validate_file_exists("")
    assert ok is False


# --- validate_required_fields ---

def test_required_fields_all_present(validator):
    ok, missing = validator.validate_required_fields({
        "recipient": "user@example.com",
        "status": "待发送",
    })
    assert ok is True
    assert missing == []


def test_required_fields_missing(validator):
    ok, missing = validator.validate_required_fields({
        "recipient": "",
        "status": "待发送",
        "subject": None,
    })
    assert ok is False
    assert "recipient" in missing
    assert "subject" in missing
    assert "status" not in missing
