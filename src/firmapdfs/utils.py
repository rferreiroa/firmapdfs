"""
Utility Functions
=================

Common utility functions used across the FirmaPDFs application.
"""

import hashlib
import os
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Optional, TypeVar
from functools import wraps

from .logger import get_logger

logger = get_logger(__name__)


# =============================================================================
# Type Variables
# =============================================================================

T = TypeVar("T")


# =============================================================================
# File Utilities
# =============================================================================

def calculate_file_hash(filepath: Path, algorithm: str = "sha256") -> str:
    """
    Calculate hash of a file.

    Args:
        filepath: Path to the file
        algorithm: Hash algorithm (sha256, md5, etc.)

    Returns:
        Hex digest of the file hash
    """
    hash_func = hashlib.new(algorithm)
    with open(filepath, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            hash_func.update(chunk)
    return hash_func.hexdigest()


def is_file_locked(filepath: Path) -> bool:
    """
    Check if a file is locked by another process.

    Args:
        filepath: Path to the file

    Returns:
        True if file is locked, False otherwise
    """
    if not filepath.exists():
        return False

    try:
        # Try to open file exclusively
        with open(filepath, "r+b"):
            return False
    except (IOError, OSError):
        return True


def wait_for_file_unlock(
    filepath: Path,
    timeout: int = 30,
    poll_interval: float = 0.5
) -> bool:
    """
    Wait for a file to be unlocked.

    Args:
        filepath: Path to the file
        timeout: Maximum wait time in seconds
        poll_interval: Time between checks in seconds

    Returns:
        True if file is unlocked within timeout, False otherwise
    """
    start_time = time.time()

    while time.time() - start_time < timeout:
        if not is_file_locked(filepath):
            return True
        time.sleep(poll_interval)

    return False


def safe_filename(filename: str, replacement: str = "_") -> str:
    """
    Convert filename to safe version without invalid characters.

    Args:
        filename: Original filename
        replacement: Character to replace invalid chars with

    Returns:
        Safe filename
    """
    # Characters not allowed in Windows filenames
    invalid_chars = '<>:"/\\|?*'

    result = filename
    for char in invalid_chars:
        result = result.replace(char, replacement)

    # Remove leading/trailing spaces and dots
    result = result.strip(". ")

    # Limit length
    if len(result) > 200:
        name, ext = os.path.splitext(result)
        result = name[:200 - len(ext)] + ext

    return result


def ensure_unique_filename(filepath: Path) -> Path:
    """
    Ensure filename is unique by adding suffix if needed.

    Args:
        filepath: Desired file path

    Returns:
        Unique file path (may have _1, _2, etc. suffix)
    """
    if not filepath.exists():
        return filepath

    stem = filepath.stem
    suffix = filepath.suffix
    parent = filepath.parent

    counter = 1
    while True:
        new_path = parent / f"{stem}_{counter}{suffix}"
        if not new_path.exists():
            return new_path
        counter += 1


def get_file_timestamp(filepath: Path) -> datetime:
    """
    Get modification timestamp of a file.

    Args:
        filepath: Path to the file

    Returns:
        Datetime of last modification
    """
    mtime = filepath.stat().st_mtime
    return datetime.fromtimestamp(mtime)


# =============================================================================
# Retry Decorator
# =============================================================================

def retry(
    max_attempts: int = 3,
    delay: float = 1.0,
    backoff: float = 2.0,
    exceptions: tuple = (Exception,),
) -> Callable:
    """
    Decorator for retrying a function on failure.

    Args:
        max_attempts: Maximum number of attempts
        delay: Initial delay between retries (seconds)
        backoff: Multiplier for delay after each attempt
        exceptions: Tuple of exceptions to catch

    Returns:
        Decorated function
    """
    def decorator(func: Callable[..., T]) -> Callable[..., T]:
        @wraps(func)
        def wrapper(*args, **kwargs) -> T:
            current_delay = delay

            for attempt in range(1, max_attempts + 1):
                try:
                    return func(*args, **kwargs)
                except exceptions as e:
                    if attempt == max_attempts:
                        logger.error(
                            f"{func.__name__} failed after {max_attempts} attempts: {e}"
                        )
                        raise

                    logger.warning(
                        f"{func.__name__} attempt {attempt}/{max_attempts} failed: {e}. "
                        f"Retrying in {current_delay:.1f}s..."
                    )
                    time.sleep(current_delay)
                    current_delay *= backoff

            # Should never reach here
            raise RuntimeError("Retry loop exited unexpectedly")

        return wrapper
    return decorator


# =============================================================================
# Timing Context Manager
# =============================================================================

class Timer:
    """
    Context manager for timing operations.

    Usage:
        with Timer() as t:
            do_something()
        print(f"Took {t.elapsed_ms}ms")
    """

    def __init__(self):
        self.start_time: Optional[float] = None
        self.end_time: Optional[float] = None

    def __enter__(self) -> "Timer":
        self.start_time = time.perf_counter()
        return self

    def __exit__(self, *args) -> None:
        self.end_time = time.perf_counter()

    @property
    def elapsed(self) -> float:
        """Elapsed time in seconds."""
        if self.start_time is None:
            return 0.0
        end = self.end_time or time.perf_counter()
        return end - self.start_time

    @property
    def elapsed_ms(self) -> int:
        """Elapsed time in milliseconds."""
        return int(self.elapsed * 1000)


# =============================================================================
# Platform Utilities
# =============================================================================

def is_windows() -> bool:
    """Check if running on Windows."""
    return sys.platform == "win32"


def is_admin() -> bool:
    """Check if running with administrator privileges (Windows)."""
    if not is_windows():
        return os.geteuid() == 0

    try:
        import ctypes
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception:
        return False


def get_windows_version() -> Optional[str]:
    """Get Windows version string."""
    if not is_windows():
        return None

    try:
        import platform
        return platform.version()
    except Exception:
        return None


# =============================================================================
# Date Utilities
# =============================================================================

def format_date(date: Optional[datetime] = None, format_str: str = "%d/%m/%Y") -> str:
    """
    Format date according to specified format.

    Args:
        date: Datetime object (default: now)
        format_str: strftime format string

    Returns:
        Formatted date string
    """
    if date is None:
        date = datetime.now()
    return date.strftime(format_str)


def parse_date(date_str: str, format_str: str = "%d/%m/%Y") -> Optional[datetime]:
    """
    Parse date string.

    Args:
        date_str: Date string to parse
        format_str: Expected format

    Returns:
        Datetime object or None if parsing fails
    """
    try:
        return datetime.strptime(date_str, format_str)
    except ValueError:
        return None


# =============================================================================
# Progress Tracking
# =============================================================================

class ProgressTracker:
    """
    Simple progress tracker for batch operations.

    Usage:
        tracker = ProgressTracker(total=100)
        for item in items:
            process(item)
            tracker.update()
            print(tracker.progress_str())
    """

    def __init__(self, total: int):
        self.total = total
        self.current = 0
        self.errors = 0
        self.skipped = 0
        self.start_time = time.time()

    def update(self, success: bool = True, skipped: bool = False) -> None:
        """Update progress."""
        self.current += 1
        if skipped:
            self.skipped += 1
        elif not success:
            self.errors += 1

    @property
    def percentage(self) -> float:
        """Get completion percentage."""
        if self.total == 0:
            return 100.0
        return (self.current / self.total) * 100

    @property
    def elapsed_seconds(self) -> float:
        """Get elapsed time in seconds."""
        return time.time() - self.start_time

    @property
    def eta_seconds(self) -> float:
        """Estimate remaining time in seconds."""
        if self.current == 0:
            return 0.0
        rate = self.elapsed_seconds / self.current
        remaining = self.total - self.current
        return rate * remaining

    def progress_str(self) -> str:
        """Get progress string for display."""
        return (
            f"[{self.current}/{self.total}] {self.percentage:.1f}% "
            f"(errors: {self.errors}, skipped: {self.skipped})"
        )


# =============================================================================
# String Utilities
# =============================================================================

def truncate(text: str, max_length: int = 50, suffix: str = "...") -> str:
    """
    Truncate text to maximum length.

    Args:
        text: Text to truncate
        max_length: Maximum length
        suffix: Suffix to add if truncated

    Returns:
        Truncated text
    """
    if len(text) <= max_length:
        return text
    return text[: max_length - len(suffix)] + suffix


def normalize_text(text: str) -> str:
    """
    Normalize text for comparison (lowercase, strip, collapse whitespace).

    Args:
        text: Text to normalize

    Returns:
        Normalized text
    """
    import re
    text = text.lower().strip()
    text = re.sub(r"\s+", " ", text)
    return text
