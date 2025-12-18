"""
Tests for Utility Functions
===========================
"""

import pytest
import time
from pathlib import Path
from datetime import datetime

from firmapdfs.utils import (
    calculate_file_hash,
    safe_filename,
    ensure_unique_filename,
    format_date,
    parse_date,
    truncate,
    normalize_text,
    retry,
    Timer,
    ProgressTracker,
)


class TestFileUtils:
    """Tests for file utility functions."""

    def test_calculate_file_hash(self, tmp_path):
        """Test file hash calculation."""
        # Create a test file
        test_file = tmp_path / "test.txt"
        test_file.write_text("Hello, World!")

        hash1 = calculate_file_hash(test_file)
        hash2 = calculate_file_hash(test_file)

        assert hash1 == hash2
        assert len(hash1) == 64  # SHA256 hex length

    def test_calculate_file_hash_different_content(self, tmp_path):
        """Test that different content produces different hash."""
        file1 = tmp_path / "file1.txt"
        file2 = tmp_path / "file2.txt"

        file1.write_text("Content 1")
        file2.write_text("Content 2")

        hash1 = calculate_file_hash(file1)
        hash2 = calculate_file_hash(file2)

        assert hash1 != hash2

    def test_safe_filename(self):
        """Test safe filename conversion."""
        assert safe_filename("normal.txt") == "normal.txt"
        assert safe_filename("file:name.txt") == "file_name.txt"
        assert safe_filename("file<>name.txt") == "file__name.txt"
        assert safe_filename('file"name.txt') == "file_name.txt"
        assert safe_filename("  file.txt  ") == "file.txt"

    def test_safe_filename_long(self):
        """Test safe filename with very long name."""
        long_name = "a" * 300 + ".txt"
        safe = safe_filename(long_name)

        assert len(safe) <= 204  # 200 + .txt

    def test_ensure_unique_filename(self, tmp_path):
        """Test unique filename generation."""
        # Create initial file
        file1 = tmp_path / "test.txt"
        file1.write_text("content")

        # Should return different path
        unique = ensure_unique_filename(file1)

        assert unique != file1
        assert unique.stem == "test_1"
        assert unique.suffix == ".txt"

    def test_ensure_unique_filename_multiple(self, tmp_path):
        """Test unique filename with multiple existing files."""
        # Create multiple files
        (tmp_path / "test.txt").write_text("1")
        (tmp_path / "test_1.txt").write_text("2")
        (tmp_path / "test_2.txt").write_text("3")

        unique = ensure_unique_filename(tmp_path / "test.txt")

        assert unique.name == "test_3.txt"


class TestDateUtils:
    """Tests for date utility functions."""

    def test_format_date_default(self):
        """Test date formatting with default format."""
        date = datetime(2024, 1, 15)
        formatted = format_date(date)

        assert formatted == "15/01/2024"

    def test_format_date_custom(self):
        """Test date formatting with custom format."""
        date = datetime(2024, 1, 15)
        formatted = format_date(date, "%Y-%m-%d")

        assert formatted == "2024-01-15"

    def test_format_date_now(self):
        """Test formatting current date."""
        formatted = format_date()

        # Should be today's date
        today = datetime.now().strftime("%d/%m/%Y")
        assert formatted == today

    def test_parse_date(self):
        """Test date parsing."""
        parsed = parse_date("15/01/2024")

        assert parsed is not None
        assert parsed.day == 15
        assert parsed.month == 1
        assert parsed.year == 2024

    def test_parse_date_invalid(self):
        """Test parsing invalid date."""
        parsed = parse_date("invalid")

        assert parsed is None


class TestStringUtils:
    """Tests for string utility functions."""

    def test_truncate_short(self):
        """Test truncate with short text."""
        text = "Hello"
        result = truncate(text, max_length=50)

        assert result == "Hello"

    def test_truncate_long(self):
        """Test truncate with long text."""
        text = "This is a very long text that should be truncated"
        result = truncate(text, max_length=20)

        assert len(result) == 20
        assert result.endswith("...")

    def test_normalize_text(self):
        """Test text normalization."""
        assert normalize_text("  Hello  World  ") == "hello world"
        assert normalize_text("UPPERCASE") == "uppercase"
        assert normalize_text("multiple   spaces") == "multiple spaces"


class TestRetryDecorator:
    """Tests for retry decorator."""

    def test_retry_success_first_attempt(self):
        """Test function that succeeds on first attempt."""
        call_count = 0

        @retry(max_attempts=3)
        def success_func():
            nonlocal call_count
            call_count += 1
            return "success"

        result = success_func()

        assert result == "success"
        assert call_count == 1

    def test_retry_success_after_failures(self):
        """Test function that succeeds after failures."""
        call_count = 0

        @retry(max_attempts=3, delay=0.01)
        def flaky_func():
            nonlocal call_count
            call_count += 1
            if call_count < 3:
                raise ValueError("Not yet")
            return "success"

        result = flaky_func()

        assert result == "success"
        assert call_count == 3

    def test_retry_all_failures(self):
        """Test function that always fails."""
        call_count = 0

        @retry(max_attempts=3, delay=0.01)
        def fail_func():
            nonlocal call_count
            call_count += 1
            raise ValueError("Always fails")

        with pytest.raises(ValueError):
            fail_func()

        assert call_count == 3


class TestTimer:
    """Tests for Timer context manager."""

    def test_timer_basic(self):
        """Test basic timer functionality."""
        with Timer() as t:
            time.sleep(0.1)

        assert t.elapsed >= 0.1
        assert t.elapsed < 0.2
        assert t.elapsed_ms >= 100

    def test_timer_fast(self):
        """Test timer with very fast operation."""
        with Timer() as t:
            pass

        assert t.elapsed >= 0
        assert t.elapsed < 0.1


class TestProgressTracker:
    """Tests for ProgressTracker."""

    def test_progress_basic(self):
        """Test basic progress tracking."""
        tracker = ProgressTracker(total=10)

        assert tracker.percentage == 0
        assert tracker.current == 0

        tracker.update()
        assert tracker.current == 1
        assert tracker.percentage == 10

    def test_progress_errors(self):
        """Test progress tracking with errors."""
        tracker = ProgressTracker(total=10)

        tracker.update(success=True)
        tracker.update(success=False)
        tracker.update(skipped=True)

        assert tracker.current == 3
        assert tracker.errors == 1
        assert tracker.skipped == 1

    def test_progress_string(self):
        """Test progress string generation."""
        tracker = ProgressTracker(total=10)
        tracker.update()
        tracker.update(success=False)

        progress_str = tracker.progress_str()

        assert "[2/10]" in progress_str
        assert "20.0%" in progress_str
        assert "errors: 1" in progress_str
