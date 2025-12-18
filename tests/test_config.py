"""
Tests for Configuration Module
==============================
"""

import pytest
from pathlib import Path
import tempfile
import os

from firmapdfs.config import Config, RulesConfig, PathsConfig, WordConfig


class TestConfig:
    """Tests for Config class."""

    def test_default_config(self):
        """Test creating config with default values."""
        config = Config.get_default()

        assert config is not None
        assert config.word.engine in ("com", "docx")
        assert config.processing.dry_run is False
        assert config.logging.level == "INFO"

    def test_paths_config(self, tmp_path):
        """Test PathsConfig creation."""
        data = {
            "inbox": "input",
            "output_docx": "output/docx",
        }

        paths = PathsConfig.from_dict(data, tmp_path)

        assert paths.inbox == tmp_path / "input"
        assert paths.output_docx == tmp_path / "output/docx"

    def test_word_config_defaults(self):
        """Test WordConfig default values."""
        config = WordConfig()

        assert config.engine == "com"
        assert config.date_format == "%d/%m/%Y"
        assert config.locale == "es_ES"

    def test_config_load_from_file(self, tmp_path):
        """Test loading config from YAML file."""
        # Create minimal config file
        config_dir = tmp_path / "config"
        config_dir.mkdir()

        config_content = """
paths:
  inbox: "test_inbox"
word:
  engine: "docx"
processing:
  dry_run: true
"""
        config_file = config_dir / "config.yaml"
        config_file.write_text(config_content)

        config = Config.load(str(config_file), base_path=tmp_path)

        assert config.paths.inbox == tmp_path / "test_inbox"
        assert config.word.engine == "docx"
        assert config.processing.dry_run is True

    def test_config_file_not_found(self, tmp_path):
        """Test error when config file doesn't exist."""
        with pytest.raises(FileNotFoundError):
            Config.load(str(tmp_path / "nonexistent.yaml"))

    def test_ensure_directories(self, tmp_path):
        """Test directory creation."""
        data = {
            "inbox": "inbox",
            "output_docx": "processed/docx",
            "output_pdf": "processed/pdf",
            "output_signed": "processed/signed",
        }

        config_data = {"paths": data}

        config_dir = tmp_path / "config"
        config_dir.mkdir()
        config_file = config_dir / "config.yaml"

        import yaml
        config_file.write_text(yaml.dump(config_data))

        config = Config.load(str(config_file), base_path=tmp_path)
        config.ensure_directories()

        assert (tmp_path / "inbox").exists()
        assert (tmp_path / "processed/docx").exists()
        assert (tmp_path / "processed/pdf").exists()

    def test_config_validate(self, tmp_path):
        """Test configuration validation."""
        config = Config.get_default()
        issues = config.validate()

        # Should have some issues since cert doesn't exist
        assert isinstance(issues, list)


class TestRulesConfig:
    """Tests for RulesConfig class."""

    def test_empty_rules(self, tmp_path):
        """Test loading when rules file doesn't exist."""
        rules = RulesConfig.load(base_path=tmp_path)

        assert rules.version == "1.0"
        assert len(rules.rules) == 0
        assert len(rules.checkbox_mappings) == 0

    def test_load_rules_from_file(self, tmp_path):
        """Test loading rules from YAML file."""
        config_dir = tmp_path / "config"
        config_dir.mkdir()

        rules_content = """
version: "1.0"
checkbox_mappings:
  aprobado:
    - "chkAprobado"
    - "Aprobado"
rules:
  - name: "test_rule"
    description: "Test rule"
    enabled: true
    priority: 100
    conditions: []
    actions:
      - type: "set_checkbox"
        checkbox: "aprobado"
        value: true
"""
        rules_file = config_dir / "rules.yaml"
        rules_file.write_text(rules_content)

        rules = RulesConfig.load(base_path=tmp_path)

        assert rules.version == "1.0"
        assert len(rules.rules) == 1
        assert rules.rules[0].name == "test_rule"
        assert "aprobado" in rules.checkbox_mappings

    def test_get_enabled_rules(self, tmp_path):
        """Test filtering enabled rules."""
        config_dir = tmp_path / "config"
        config_dir.mkdir()

        rules_content = """
rules:
  - name: "enabled_rule"
    enabled: true
    priority: 50
    conditions: []
    actions: []
  - name: "disabled_rule"
    enabled: false
    priority: 50
    conditions: []
    actions: []
"""
        rules_file = config_dir / "rules.yaml"
        rules_file.write_text(rules_content)

        rules = RulesConfig.load(base_path=tmp_path)
        enabled = rules.get_enabled_rules()

        assert len(enabled) == 1
        assert enabled[0].name == "enabled_rule"

    def test_find_checkbox_tags(self, tmp_path):
        """Test checkbox tag lookup."""
        config_dir = tmp_path / "config"
        config_dir.mkdir()

        rules_content = """
checkbox_mappings:
  aprobado:
    - "chkAprobado"
    - "CheckBoxAprobado"
"""
        rules_file = config_dir / "rules.yaml"
        rules_file.write_text(rules_content)

        rules = RulesConfig.load(base_path=tmp_path)
        tags = rules.find_checkbox_tags("aprobado")

        assert "chkAprobado" in tags
        assert "CheckBoxAprobado" in tags

    def test_find_checkbox_tags_unknown(self, tmp_path):
        """Test checkbox tag lookup for unknown name."""
        rules = RulesConfig.load(base_path=tmp_path)
        tags = rules.find_checkbox_tags("unknown_checkbox")

        # Should return the name itself as fallback
        assert tags == ["unknown_checkbox"]


class TestConfigIntegration:
    """Integration tests for config and rules."""

    def test_full_config_loading(self, tmp_path):
        """Test loading both config and rules together."""
        config_dir = tmp_path / "config"
        config_dir.mkdir()

        # Create config file
        config_content = """
paths:
  inbox: "inbox"
word:
  engine: "docx"
signature:
  certificate_path: "config/cert.pfx"
"""
        (config_dir / "config.yaml").write_text(config_content)

        # Create rules file
        rules_content = """
rules:
  - name: "date_rule"
    enabled: true
    priority: 100
    conditions: []
    actions:
      - type: "set_date"
        field: "fecha"
"""
        (config_dir / "rules.yaml").write_text(rules_content)

        # Load both
        config = Config.load(base_path=tmp_path)
        rules = RulesConfig.load(base_path=tmp_path)

        assert config.word.engine == "docx"
        assert len(rules.get_enabled_rules()) == 1
