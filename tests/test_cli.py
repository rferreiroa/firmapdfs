"""
Tests for CLI Module
====================
"""

import pytest
from click.testing import CliRunner
from pathlib import Path
import yaml

from firmapdfs.cli import cli
from firmapdfs import __version__


@pytest.fixture
def runner():
    """Create CLI test runner."""
    return CliRunner()


@pytest.fixture
def setup_project(tmp_path):
    """Set up a minimal project structure for testing."""
    # Create config directory and files
    config_dir = tmp_path / "config"
    config_dir.mkdir()

    config_content = {
        "paths": {
            "inbox": "samples/inbox",
            "output_docx": "processed/docx",
            "output_pdf": "processed/pdf",
            "output_signed": "processed/signed",
            "logs": "logs",
        },
        "word": {
            "engine": "docx",
        },
        "logging": {
            "level": "WARNING",
            "console": False,
            "file": False,
        },
    }
    (config_dir / "config.yaml").write_text(yaml.dump(config_content))

    # Create directories
    (tmp_path / "samples/inbox").mkdir(parents=True)
    (tmp_path / "processed/docx").mkdir(parents=True)
    (tmp_path / "processed/pdf").mkdir(parents=True)
    (tmp_path / "processed/signed").mkdir(parents=True)
    (tmp_path / "logs").mkdir(parents=True)

    return tmp_path


class TestCLIBasic:
    """Basic CLI tests."""

    def test_version(self, runner):
        """Test --version flag."""
        result = runner.invoke(cli, ["--version"])

        assert result.exit_code == 0
        assert __version__ in result.output

    def test_help(self, runner):
        """Test --help flag."""
        result = runner.invoke(cli, ["--help"])

        assert result.exit_code == 0
        assert "FirmaPDFs" in result.output
        assert "process" in result.output
        assert "watch" in result.output

    def test_no_command_shows_help(self, runner):
        """Test that running without command shows help."""
        result = runner.invoke(cli)

        assert result.exit_code == 0
        assert "FirmaPDFs" in result.output


class TestInitCommand:
    """Tests for init command."""

    def test_init_creates_directories(self, runner, tmp_path):
        """Test that init creates directory structure."""
        with runner.isolated_filesystem(temp_dir=tmp_path) as td:
            result = runner.invoke(cli, ["init"])

            assert result.exit_code == 0
            assert "Initialization complete" in result.output

            # Check directories were created
            assert Path(td, "config").exists()
            assert Path(td, "samples/inbox").exists()
            assert Path(td, "processed/docx").exists()

    def test_init_with_existing_config(self, runner, tmp_path):
        """Test init when config already exists."""
        with runner.isolated_filesystem(temp_dir=tmp_path) as td:
            # Create existing config
            config_dir = Path(td, "config")
            config_dir.mkdir()
            (config_dir / "config.yaml").write_text("existing: true")

            result = runner.invoke(cli, ["init"])

            assert result.exit_code == 0
            assert "already exists" in result.output


class TestProcessCommand:
    """Tests for process command."""

    def test_process_empty_inbox(self, runner, setup_project):
        """Test processing with empty inbox."""
        with runner.isolated_filesystem(temp_dir=setup_project):
            result = runner.invoke(cli, ["process"])

            assert result.exit_code == 0
            assert "No .docx files found" in result.output

    def test_process_dry_run(self, runner, setup_project):
        """Test process with dry-run flag."""
        with runner.isolated_filesystem(temp_dir=setup_project):
            # Create a test file
            inbox = setup_project / "samples/inbox"
            (inbox / "test.docx").write_bytes(b"dummy")

            result = runner.invoke(cli, ["--dry-run", "process"])

            assert result.exit_code == 0
            assert "DRY RUN" in result.output

    def test_process_with_verbose(self, runner, setup_project):
        """Test process with verbose flag."""
        with runner.isolated_filesystem(temp_dir=setup_project):
            result = runner.invoke(cli, ["-v", "process"])

            assert result.exit_code == 0


class TestStatusCommand:
    """Tests for status command."""

    def test_status_with_config(self, runner, setup_project):
        """Test status command with valid config."""
        with runner.isolated_filesystem(temp_dir=setup_project):
            result = runner.invoke(cli, ["status"])

            assert result.exit_code == 0
            assert "FirmaPDFs - System Status" in result.output
            assert "Version:" in result.output

    def test_status_shows_directories(self, runner, setup_project):
        """Test that status shows directory info."""
        with runner.isolated_filesystem(temp_dir=setup_project):
            result = runner.invoke(cli, ["status"])

            assert "Directories:" in result.output
            assert "Inbox" in result.output


class TestCheckCommand:
    """Tests for check command."""

    def test_check_basic(self, runner, setup_project):
        """Test check command."""
        with runner.isolated_filesystem(temp_dir=setup_project):
            result = runner.invoke(cli, ["check"])

            # Should show checking output
            assert "Checking" in result.output
            assert "Python version" in result.output

    def test_check_with_fix(self, runner, setup_project):
        """Test check command with --fix flag."""
        with runner.isolated_filesystem(temp_dir=setup_project):
            result = runner.invoke(cli, ["check", "--fix"])

            # Should either create dirs or show already created
            assert result.exit_code in (0, 1)  # May have warnings


class TestReportCommand:
    """Tests for report command."""

    def test_report_basic(self, runner, setup_project):
        """Test report command."""
        with runner.isolated_filesystem(temp_dir=setup_project):
            result = runner.invoke(cli, ["report"])

            assert result.exit_code == 0
            assert "Report" in result.output


class TestCLIOptions:
    """Tests for CLI global options."""

    def test_config_option(self, runner, setup_project):
        """Test custom config path option."""
        with runner.isolated_filesystem(temp_dir=setup_project):
            result = runner.invoke(cli, ["-c", "config/config.yaml", "status"])

            assert result.exit_code == 0

    def test_invalid_config_for_process(self, runner, tmp_path):
        """Test with invalid config path for process command."""
        with runner.isolated_filesystem(temp_dir=tmp_path):
            result = runner.invoke(cli, ["-c", "nonexistent.yaml", "process"])

            assert result.exit_code == 1
            assert "not found" in result.output.lower() or "error" in result.output.lower()
