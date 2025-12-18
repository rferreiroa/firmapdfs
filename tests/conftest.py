"""
Pytest Configuration and Fixtures
=================================

Shared fixtures for all tests.
"""

import pytest
from pathlib import Path
import tempfile
import os


@pytest.fixture
def tmp_path():
    """Create a temporary directory for test files."""
    with tempfile.TemporaryDirectory() as tmpdir:
        yield Path(tmpdir)


@pytest.fixture
def sample_docx_content():
    """
    Return bytes for a minimal valid DOCX file.

    Note: This is a simplified placeholder. For actual Word document testing,
    you would need to create real DOCX files using python-docx.
    """
    # Minimal ZIP structure for a DOCX (not fully valid but useful for testing)
    return b"PK\x03\x04dummy_docx_content"


@pytest.fixture
def project_root():
    """Get the project root directory."""
    # Find project root by looking for pyproject.toml
    current = Path(__file__).parent
    while current != current.parent:
        if (current / "pyproject.toml").exists():
            return current
        current = current.parent
    return Path(__file__).parent.parent


@pytest.fixture
def sample_config():
    """Return sample configuration dictionary."""
    return {
        "paths": {
            "inbox": "samples/inbox",
            "output_docx": "processed/docx",
            "output_pdf": "processed/pdf",
            "output_signed": "processed/signed",
            "logs": "logs",
        },
        "word": {
            "engine": "docx",
            "date_format": "%d/%m/%Y",
        },
        "processing": {
            "dry_run": False,
            "skip_processed": True,
        },
        "logging": {
            "level": "INFO",
            "console": False,
            "file": False,
        },
    }


@pytest.fixture
def sample_rules():
    """Return sample rules dictionary."""
    return {
        "version": "1.0",
        "checkbox_mappings": {
            "aprobado": ["chkAprobado", "Aprobado"],
            "rechazado": ["chkRechazado", "Rechazado"],
        },
        "field_mappings": {
            "fecha": ["Fecha", "DATE"],
        },
        "rules": [
            {
                "name": "test_rule",
                "description": "Test rule",
                "enabled": True,
                "priority": 100,
                "conditions": [],
                "actions": [
                    {"type": "set_date", "field": "fecha"}
                ],
            }
        ],
    }


@pytest.fixture
def setup_minimal_project(tmp_path, sample_config, sample_rules):
    """
    Set up a minimal project structure for testing.

    Creates config files and directory structure.
    """
    import yaml

    # Create config directory
    config_dir = tmp_path / "config"
    config_dir.mkdir()

    # Write config.yaml
    config_file = config_dir / "config.yaml"
    config_file.write_text(yaml.dump(sample_config))

    # Write rules.yaml
    rules_file = config_dir / "rules.yaml"
    rules_file.write_text(yaml.dump(sample_rules))

    # Create directories
    dirs = [
        "samples/inbox",
        "samples/templates",
        "processed/docx",
        "processed/pdf",
        "processed/signed",
        "logs",
    ]
    for d in dirs:
        (tmp_path / d).mkdir(parents=True, exist_ok=True)

    return tmp_path


# Skip markers for Windows-specific tests
skip_on_non_windows = pytest.mark.skipif(
    os.name != 'nt',
    reason="Windows-specific test"
)

skip_without_word = pytest.mark.skipif(
    os.name != 'nt',  # Simplified check - real check would try to import win32com
    reason="Requires Microsoft Word"
)
