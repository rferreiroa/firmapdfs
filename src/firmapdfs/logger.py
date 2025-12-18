"""
Logging Module
==============

Configures logging for the FirmaPDFs application with support for
console output, file rotation, and colored output.
"""

import logging
import sys
from datetime import datetime
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Optional

# Try to import colorama for colored console output on Windows
try:
    import colorama
    from colorama import Fore, Style
    colorama.init(autoreset=True)
    COLORAMA_AVAILABLE = True
except ImportError:
    COLORAMA_AVAILABLE = False


# =============================================================================
# Custom Formatter with Colors
# =============================================================================

class ColoredFormatter(logging.Formatter):
    """
    Logging formatter that adds colors to log levels.

    Uses colorama for Windows compatibility.
    """

    COLORS = {
        logging.DEBUG: Fore.CYAN if COLORAMA_AVAILABLE else "",
        logging.INFO: Fore.GREEN if COLORAMA_AVAILABLE else "",
        logging.WARNING: Fore.YELLOW if COLORAMA_AVAILABLE else "",
        logging.ERROR: Fore.RED if COLORAMA_AVAILABLE else "",
        logging.CRITICAL: Fore.RED + Style.BRIGHT if COLORAMA_AVAILABLE else "",
    }
    RESET = Style.RESET_ALL if COLORAMA_AVAILABLE else ""

    def format(self, record: logging.LogRecord) -> str:
        # Add color to levelname
        color = self.COLORS.get(record.levelno, "")
        original_levelname = record.levelname
        record.levelname = f"{color}{record.levelname}{self.RESET}"

        # Format the message
        result = super().format(record)

        # Restore original levelname
        record.levelname = original_levelname

        return result


# =============================================================================
# Logging Setup
# =============================================================================

def setup_logging(
    level: str = "INFO",
    console: bool = True,
    file: bool = True,
    log_dir: Optional[Path] = None,
    file_pattern: str = "firmapdfs_{date}.log",
    max_size_mb: int = 10,
    backup_count: int = 5,
    log_format: str = "%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
    date_format: str = "%Y-%m-%d %H:%M:%S",
) -> logging.Logger:
    """
    Configure logging for the application.

    Args:
        level: Log level (DEBUG, INFO, WARNING, ERROR, CRITICAL)
        console: Enable console output
        file: Enable file output
        log_dir: Directory for log files
        file_pattern: Log file name pattern ({date} will be replaced)
        max_size_mb: Maximum log file size in MB
        backup_count: Number of backup files to keep
        log_format: Log message format
        date_format: Date format in log messages

    Returns:
        Root logger instance
    """
    # Get root logger
    root_logger = logging.getLogger()

    # Clear existing handlers
    root_logger.handlers.clear()

    # Set log level
    numeric_level = getattr(logging, level.upper(), logging.INFO)
    root_logger.setLevel(numeric_level)

    # Create formatters
    plain_formatter = logging.Formatter(log_format, datefmt=date_format)
    colored_formatter = ColoredFormatter(log_format, datefmt=date_format)

    # Console handler
    if console:
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(numeric_level)
        # Use colored formatter for console
        console_handler.setFormatter(colored_formatter)
        root_logger.addHandler(console_handler)

    # File handler
    if file and log_dir:
        log_dir = Path(log_dir)
        log_dir.mkdir(parents=True, exist_ok=True)

        # Generate log filename
        today = datetime.now().strftime("%Y-%m-%d")
        log_filename = file_pattern.replace("{date}", today)
        log_path = log_dir / log_filename

        # Rotating file handler
        file_handler = RotatingFileHandler(
            log_path,
            maxBytes=max_size_mb * 1024 * 1024,
            backupCount=backup_count,
            encoding="utf-8",
        )
        file_handler.setLevel(numeric_level)
        file_handler.setFormatter(plain_formatter)
        root_logger.addHandler(file_handler)

    return root_logger


def get_logger(name: str) -> logging.Logger:
    """
    Get a logger instance with the given name.

    Args:
        name: Logger name (usually __name__)

    Returns:
        Logger instance
    """
    return logging.getLogger(name)


# =============================================================================
# Logging Context Manager
# =============================================================================

class LogContext:
    """
    Context manager for logging operations with duration tracking.

    Usage:
        with LogContext(logger, "Processing file", filename=file.name):
            process_file(file)
    """

    def __init__(
        self,
        logger: logging.Logger,
        operation: str,
        level: int = logging.INFO,
        **context
    ):
        self.logger = logger
        self.operation = operation
        self.level = level
        self.context = context
        self.start_time = None

    def __enter__(self):
        self.start_time = datetime.now()
        context_str = " ".join(f"{k}={v}" for k, v in self.context.items())
        self.logger.log(self.level, f"START {self.operation} {context_str}".strip())
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        duration = (datetime.now() - self.start_time).total_seconds() * 1000
        context_str = " ".join(f"{k}={v}" for k, v in self.context.items())

        if exc_type is None:
            self.logger.log(
                self.level,
                f"END {self.operation} {context_str} duration={duration:.0f}ms".strip()
            )
        else:
            self.logger.error(
                f"FAIL {self.operation} {context_str} duration={duration:.0f}ms "
                f"error={exc_type.__name__}: {exc_val}"
            )

        # Don't suppress exceptions
        return False


# =============================================================================
# Sensitive Data Filter
# =============================================================================

class SensitiveDataFilter(logging.Filter):
    """
    Filter that removes or masks sensitive data from log messages.

    Masks:
    - Passwords
    - Certificate paths (optionally)
    - File paths with sensitive directories
    """

    SENSITIVE_PATTERNS = [
        ("password", "********"),
        ("pwd", "********"),
        ("secret", "********"),
        ("token", "********"),
        ("key", "********"),
    ]

    def filter(self, record: logging.LogRecord) -> bool:
        # Mask sensitive data in message
        message = str(record.msg)
        for pattern, replacement in self.SENSITIVE_PATTERNS:
            # Simple pattern matching (case-insensitive)
            lower_msg = message.lower()
            if pattern in lower_msg:
                # Find the pattern and mask the value after it
                # This handles cases like "password=secret123"
                import re
                pattern_regex = re.compile(
                    rf'({pattern}\s*[=:]\s*)([^\s,;]+)',
                    re.IGNORECASE
                )
                message = pattern_regex.sub(rf'\1{replacement}', message)

        record.msg = message
        return True


def add_sensitive_filter(logger: logging.Logger) -> None:
    """Add sensitive data filter to a logger."""
    logger.addFilter(SensitiveDataFilter())
