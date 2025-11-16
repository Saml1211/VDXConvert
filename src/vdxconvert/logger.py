# -*- coding: utf-8 -*-
"""
Logging configuration for VDXConvert.

Provides centralized logging setup with console and file output,
with optional colored output support.
"""

import logging
from typing import ClassVar, Dict, Optional

from .config import LOGS_DIR, LOG_FILE_NAME, LOG_FORMAT_FILE, LOG_FORMAT_CONSOLE

# Try to import optional color support
try:
    import colorama
    from colorama import Fore, Style
    colorama.init(autoreset=True)
    COLOR_SUPPORT = True
except ImportError:
    COLOR_SUPPORT = False


class ColoredFormatter(logging.Formatter):
    """Custom formatter that adds colors to log levels."""

    formats: ClassVar[Dict[int, str]] = {
        logging.DEBUG: Fore.CYAN + '%(message)s' + Style.RESET_ALL if COLOR_SUPPORT else '%(message)s',
        logging.INFO: '%(message)s',
        logging.WARNING: Fore.YELLOW + '%(message)s' + Style.RESET_ALL if COLOR_SUPPORT else '%(message)s',
        logging.ERROR: Fore.RED + '%(message)s' + Style.RESET_ALL if COLOR_SUPPORT else '%(message)s',
        logging.CRITICAL: Fore.RED + Style.BRIGHT + '%(message)s' + Style.RESET_ALL if COLOR_SUPPORT else '%(message)s'
    }

    def format(self, record: logging.LogRecord) -> str:
        """
        Format the log record with appropriate color.

        Args:
            record: The log record to format

        Returns:
            Formatted log message string
        """
        log_fmt = self.formats.get(record.levelno, '%(message)s')
        formatter = logging.Formatter(log_fmt)
        return formatter.format(record)


def setup_logging(verbose: bool = False) -> logging.Logger:
    """
    Configure the logging system.

    Args:
        verbose: If True, sets log level to DEBUG; otherwise INFO

    Returns:
        Configured logger instance

    Example:
        >>> logger = setup_logging(verbose=True)
        >>> logger.info("Application started")
    """
    # Ensure logs directory exists
    LOGS_DIR.mkdir(parents=True, exist_ok=True)

    log_level = logging.DEBUG if verbose else logging.INFO
    log_file = LOGS_DIR / LOG_FILE_NAME

    # Configure root logger
    logger = logging.getLogger()
    logger.setLevel(log_level)

    # Clear existing handlers
    for handler in logger.handlers[:]:
        logger.removeHandler(handler)

    # File handler
    file_handler = logging.FileHandler(log_file, mode='a', encoding='utf-8')
    file_handler.setLevel(log_level)
    file_format = logging.Formatter(LOG_FORMAT_FILE)
    file_handler.setFormatter(file_format)
    logger.addHandler(file_handler)

    # Console handler
    console_handler = logging.StreamHandler()
    console_handler.setLevel(log_level)

    if COLOR_SUPPORT:
        console_handler.setFormatter(ColoredFormatter())
    else:
        console_format = logging.Formatter(LOG_FORMAT_CONSOLE)
        console_handler.setFormatter(console_format)

    logger.addHandler(console_handler)

    return logger


def get_logger(name: Optional[str] = None) -> logging.Logger:
    """
    Get a logger instance.

    Args:
        name: Optional name for the logger

    Returns:
        Logger instance
    """
    return logging.getLogger(name)
