import pytest
import logging
import os
from src.data.log_manager import LogManager
from src.models.config import LoggingConfig


@pytest.fixture
def log_config(tmp_path):
    return LoggingConfig(
        level="DEBUG",
        file=str(tmp_path / "logs" / "test.log"),
        max_size_mb=1,
        backup_count=3,
    )


@pytest.fixture(autouse=True)
def cleanup_loggers():
    """Clean up loggers between tests to avoid handler accumulation"""
    yield
    # Remove all handlers from test loggers
    for name in ['test_logger', 'outlook_sender']:
        logger = logging.getLogger(name)
        for handler in logger.handlers[:]:
            handler.close()
            logger.removeHandler(handler)


def test_get_logger_creates_file(log_config, tmp_path):
    manager = LogManager(log_config)
    logger = manager.get_logger('test_logger')
    logger.info("Test message")
    manager.close()

    log_file = tmp_path / "logs" / "test.log"
    assert log_file.exists()
    content = log_file.read_text(encoding='utf-8')
    assert "Test message" in content


def test_get_logger_level(log_config):
    manager = LogManager(log_config)
    logger = manager.get_logger('test_logger')
    assert logger.level == logging.DEBUG
    manager.close()


def test_get_logger_has_two_handlers(log_config):
    manager = LogManager(log_config)
    logger = manager.get_logger('test_logger')
    assert len(logger.handlers) == 2  # file + console
    manager.close()


def test_close_removes_handlers(log_config):
    manager = LogManager(log_config)
    logger = manager.get_logger('test_logger')
    assert len(logger.handlers) == 2
    manager.close()
    assert len(logger.handlers) == 0


def test_log_levels(log_config, tmp_path):
    manager = LogManager(log_config)
    logger = manager.get_logger('test_logger')

    logger.debug("Debug message")
    logger.info("Info message")
    logger.warning("Warning message")
    logger.error("Error message")
    manager.close()

    content = (tmp_path / "logs" / "test.log").read_text(encoding='utf-8')
    assert "Debug message" in content
    assert "Info message" in content
    assert "Warning message" in content
    assert "Error message" in content
