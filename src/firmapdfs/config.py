"""
Configuration Management Module
===============================

Handles loading, validation, and access to configuration settings
from YAML files and environment variables.
"""

import os
import sys
from pathlib import Path
from typing import Any, Dict, Optional, List
from dataclasses import dataclass, field

import yaml


# =============================================================================
# Configuration Data Classes
# =============================================================================

@dataclass
class PathsConfig:
    """Directory paths configuration."""
    inbox: Path
    templates: Path
    output_docx: Path
    output_pdf: Path
    output_signed: Path
    logs: Path
    archive: Path

    @classmethod
    def from_dict(cls, data: Dict[str, str], base_path: Path) -> "PathsConfig":
        """Create PathsConfig from dictionary, resolving relative paths."""
        def resolve_path(p: str) -> Path:
            path = Path(p)
            if not path.is_absolute():
                path = base_path / path
            return path.resolve()

        return cls(
            inbox=resolve_path(data.get("inbox", "samples/inbox")),
            templates=resolve_path(data.get("templates", "samples/templates")),
            output_docx=resolve_path(data.get("output_docx", "processed/docx")),
            output_pdf=resolve_path(data.get("output_pdf", "processed/pdf")),
            output_signed=resolve_path(data.get("output_signed", "processed/signed")),
            logs=resolve_path(data.get("logs", "logs")),
            archive=resolve_path(data.get("archive", "processed/archive")),
        )


@dataclass
class WordConfig:
    """Word processing configuration."""
    engine: str = "com"  # "com" or "docx"
    date_format: str = "%d/%m/%Y"
    locale: str = "es_ES"
    date_field_names: List[str] = field(default_factory=lambda: ["Fecha", "DATE"])
    keep_original: bool = True
    com_timeout: int = 60

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "WordConfig":
        return cls(
            engine=data.get("engine", "com"),
            date_format=data.get("date_format", "%d/%m/%Y"),
            locale=data.get("locale", "es_ES"),
            date_field_names=data.get("date_field_names", ["Fecha", "DATE"]),
            keep_original=data.get("keep_original", True),
            com_timeout=data.get("com_timeout", 60),
        )


@dataclass
class PDFConfig:
    """PDF export configuration."""
    quality: str = "high"
    pdfa_compliance: bool = False
    optimize_for_web: bool = False
    include_properties: bool = True

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "PDFConfig":
        return cls(
            quality=data.get("quality", "high"),
            pdfa_compliance=data.get("pdfa_compliance", False),
            optimize_for_web=data.get("optimize_for_web", False),
            include_properties=data.get("include_properties", True),
        )


@dataclass
class SignatureAppearanceConfig:
    """Signature visual appearance configuration."""
    visible: bool = True
    position: str = "bottom_right"
    custom_x: int = 400
    custom_y: int = 50
    width: int = 200
    height: int = 50
    page: str = "last"
    text_template: str = "Firmado digitalmente por:\n{signer_name}\nFecha: {date}"
    font_name: str = "Helvetica"
    font_size: int = 8
    font_color: str = "#000000"
    background_color: str = ""
    border: bool = True
    border_color: str = "#000000"
    border_width: int = 1

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "SignatureAppearanceConfig":
        return cls(
            visible=data.get("visible", True),
            position=data.get("position", "bottom_right"),
            custom_x=data.get("custom_x", 400),
            custom_y=data.get("custom_y", 50),
            width=data.get("width", 200),
            height=data.get("height", 50),
            page=data.get("page", "last"),
            text_template=data.get("text_template",
                "Firmado digitalmente por:\n{signer_name}\nFecha: {date}"),
            font_name=data.get("font_name", "Helvetica"),
            font_size=data.get("font_size", 8),
            font_color=data.get("font_color", "#000000"),
            background_color=data.get("background_color", ""),
            border=data.get("border", True),
            border_color=data.get("border_color", "#000000"),
            border_width=data.get("border_width", 1),
        )


@dataclass
class SignatureConfig:
    """Digital signature configuration."""
    certificate_path: Path
    password_storage: str = "credential_manager"
    credential_name: str = "FirmaPDFs_Certificate"
    appearance: SignatureAppearanceConfig = field(default_factory=SignatureAppearanceConfig)
    reason: str = "Documento revisado y aprobado"
    location: str = "España"
    contact_info: str = ""
    timestamp_url: str = ""
    algorithm: str = "sha256"

    @classmethod
    def from_dict(cls, data: Dict[str, Any], base_path: Path) -> "SignatureConfig":
        cert_path = Path(data.get("certificate_path", "config/certificate.pfx"))
        if not cert_path.is_absolute():
            cert_path = base_path / cert_path

        appearance_data = data.get("appearance", {})

        return cls(
            certificate_path=cert_path,
            password_storage=data.get("password_storage", "credential_manager"),
            credential_name=data.get("credential_name", "FirmaPDFs_Certificate"),
            appearance=SignatureAppearanceConfig.from_dict(appearance_data),
            reason=data.get("reason", "Documento revisado y aprobado"),
            location=data.get("location", "España"),
            contact_info=data.get("contact_info", ""),
            timestamp_url=data.get("timestamp_url", ""),
            algorithm=data.get("algorithm", "sha256"),
        )


@dataclass
class ProcessingConfig:
    """Processing behavior configuration."""
    dry_run: bool = False
    skip_processed: bool = True
    processed_detection: str = "hash"
    hash_algorithm: str = "sha256"
    max_retries: int = 3
    retry_delay: int = 5
    continue_on_error: bool = True
    lock_timeout: int = 30
    delete_source: bool = False
    archive_source: bool = False

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "ProcessingConfig":
        return cls(
            dry_run=data.get("dry_run", False),
            skip_processed=data.get("skip_processed", True),
            processed_detection=data.get("processed_detection", "hash"),
            hash_algorithm=data.get("hash_algorithm", "sha256"),
            max_retries=data.get("max_retries", 3),
            retry_delay=data.get("retry_delay", 5),
            continue_on_error=data.get("continue_on_error", True),
            lock_timeout=data.get("lock_timeout", 30),
            delete_source=data.get("delete_source", False),
            archive_source=data.get("archive_source", False),
        )


@dataclass
class WatcherConfig:
    """File watcher configuration."""
    enabled: bool = False
    poll_interval: int = 5
    debounce_time: int = 2
    recursive: bool = False
    patterns: List[str] = field(default_factory=lambda: ["*.docx", "*.DOCX"])
    ignore_patterns: List[str] = field(default_factory=lambda: ["~$*", "*.tmp"])

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "WatcherConfig":
        return cls(
            enabled=data.get("enabled", False),
            poll_interval=data.get("poll_interval", 5),
            debounce_time=data.get("debounce_time", 2),
            recursive=data.get("recursive", False),
            patterns=data.get("patterns", ["*.docx", "*.DOCX"]),
            ignore_patterns=data.get("ignore_patterns", ["~$*", "*.tmp"]),
        )


@dataclass
class LoggingConfig:
    """Logging configuration."""
    level: str = "INFO"
    console: bool = True
    file: bool = True
    file_pattern: str = "firmapdfs_{date}.log"
    max_size_mb: int = 10
    backup_count: int = 5
    format: str = "%(asctime)s | %(levelname)-8s | %(name)s | %(message)s"
    date_format: str = "%Y-%m-%d %H:%M:%S"

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "LoggingConfig":
        return cls(
            level=data.get("level", "INFO"),
            console=data.get("console", True),
            file=data.get("file", True),
            file_pattern=data.get("file_pattern", "firmapdfs_{date}.log"),
            max_size_mb=data.get("max_size_mb", 10),
            backup_count=data.get("backup_count", 5),
            format=data.get("format",
                "%(asctime)s | %(levelname)-8s | %(name)s | %(message)s"),
            date_format=data.get("date_format", "%Y-%m-%d %H:%M:%S"),
        )


@dataclass
class ReportingConfig:
    """Reporting configuration."""
    csv_enabled: bool = True
    csv_path: Path = None
    csv_columns: List[str] = field(default_factory=list)
    html_enabled: bool = False
    html_path: Path = None

    @classmethod
    def from_dict(cls, data: Dict[str, Any], base_path: Path) -> "ReportingConfig":
        csv_path = Path(data.get("csv_path", "logs/report.csv"))
        if not csv_path.is_absolute():
            csv_path = base_path / csv_path

        html_path = Path(data.get("html_path", "logs/report.html"))
        if not html_path.is_absolute():
            html_path = base_path / html_path

        return cls(
            csv_enabled=data.get("csv_enabled", True),
            csv_path=csv_path,
            csv_columns=data.get("csv_columns", [
                "timestamp", "filename", "status", "message",
                "docx_output", "pdf_output", "signed_output", "duration_ms"
            ]),
            html_enabled=data.get("html_enabled", False),
            html_path=html_path,
        )


# =============================================================================
# Main Configuration Class
# =============================================================================

class Config:
    """
    Main configuration manager.

    Loads and validates configuration from YAML files and environment variables.

    Usage:
        config = Config.load("config/config.yaml")
        print(config.paths.inbox)
    """

    def __init__(
        self,
        paths: PathsConfig,
        word: WordConfig,
        pdf: PDFConfig,
        signature: SignatureConfig,
        processing: ProcessingConfig,
        watcher: WatcherConfig,
        logging: LoggingConfig,
        reporting: ReportingConfig,
        base_path: Path,
        raw_config: Dict[str, Any],
    ):
        self.paths = paths
        self.word = word
        self.pdf = pdf
        self.signature = signature
        self.processing = processing
        self.watcher = watcher
        self.logging = logging
        self.reporting = reporting
        self.base_path = base_path
        self._raw = raw_config

    @classmethod
    def load(
        cls,
        config_path: Optional[str] = None,
        base_path: Optional[Path] = None,
    ) -> "Config":
        """
        Load configuration from YAML file.

        Args:
            config_path: Path to config.yaml (default: config/config.yaml)
            base_path: Base path for resolving relative paths (default: project root)

        Returns:
            Config instance

        Raises:
            FileNotFoundError: If config file doesn't exist
            yaml.YAMLError: If config file is invalid YAML
        """
        # Determine base path (project root)
        if base_path is None:
            # Try to find project root by looking for config directory
            current = Path.cwd()
            while current != current.parent:
                if (current / "config").is_dir():
                    base_path = current
                    break
                current = current.parent
            else:
                base_path = Path.cwd()

        base_path = Path(base_path).resolve()

        # Determine config file path
        if config_path is None:
            config_path = base_path / "config" / "config.yaml"
        else:
            config_path = Path(config_path)
            if not config_path.is_absolute():
                config_path = base_path / config_path

        # Load YAML file
        if not config_path.exists():
            raise FileNotFoundError(f"Configuration file not found: {config_path}")

        with open(config_path, "r", encoding="utf-8") as f:
            raw_config = yaml.safe_load(f) or {}

        # Parse configuration sections
        return cls(
            paths=PathsConfig.from_dict(raw_config.get("paths", {}), base_path),
            word=WordConfig.from_dict(raw_config.get("word", {})),
            pdf=PDFConfig.from_dict(raw_config.get("pdf", {})),
            signature=SignatureConfig.from_dict(raw_config.get("signature", {}), base_path),
            processing=ProcessingConfig.from_dict(raw_config.get("processing", {})),
            watcher=WatcherConfig.from_dict(raw_config.get("watcher", {})),
            logging=LoggingConfig.from_dict(raw_config.get("logging", {})),
            reporting=ReportingConfig.from_dict(raw_config.get("reporting", {}), base_path),
            base_path=base_path,
            raw_config=raw_config,
        )

    @classmethod
    def get_default(cls) -> "Config":
        """Get configuration with all default values."""
        base_path = Path.cwd()
        return cls(
            paths=PathsConfig.from_dict({}, base_path),
            word=WordConfig(),
            pdf=PDFConfig(),
            signature=SignatureConfig.from_dict({}, base_path),
            processing=ProcessingConfig(),
            watcher=WatcherConfig(),
            logging=LoggingConfig(),
            reporting=ReportingConfig.from_dict({}, base_path),
            base_path=base_path,
            raw_config={},
        )

    def ensure_directories(self) -> None:
        """Create all configured directories if they don't exist."""
        directories = [
            self.paths.inbox,
            self.paths.templates,
            self.paths.output_docx,
            self.paths.output_pdf,
            self.paths.output_signed,
            self.paths.logs,
            self.paths.archive,
        ]

        for directory in directories:
            directory.mkdir(parents=True, exist_ok=True)

    def validate(self) -> List[str]:
        """
        Validate configuration and return list of warnings/errors.

        Returns:
            List of warning/error messages (empty if valid)
        """
        issues = []

        # Check inbox directory
        if not self.paths.inbox.exists():
            issues.append(f"Inbox directory does not exist: {self.paths.inbox}")

        # Check certificate file (only warn, might be configured later)
        if not self.signature.certificate_path.exists():
            issues.append(
                f"Certificate file not found: {self.signature.certificate_path} "
                "(required for signing)"
            )

        # Validate Word engine
        if self.word.engine not in ("com", "docx"):
            issues.append(
                f"Invalid Word engine: {self.word.engine}. Must be 'com' or 'docx'"
            )

        # Validate password storage method
        valid_storage = ("config", "env", "credential_manager", "dpapi")
        if self.signature.password_storage not in valid_storage:
            issues.append(
                f"Invalid password storage: {self.signature.password_storage}. "
                f"Must be one of: {valid_storage}"
            )

        # Check if running on Windows for COM
        if self.word.engine == "com" and sys.platform != "win32":
            issues.append(
                "Word COM engine requires Windows. "
                "Consider using 'docx' engine on non-Windows systems."
            )

        return issues

    def to_dict(self) -> Dict[str, Any]:
        """Convert configuration to dictionary (for serialization)."""
        return self._raw

    def __repr__(self) -> str:
        return f"Config(base_path={self.base_path})"


# =============================================================================
# Rules Configuration
# =============================================================================

@dataclass
class Rule:
    """A single processing rule."""
    name: str
    description: str
    enabled: bool
    priority: int
    conditions: List[Dict[str, Any]]
    actions: List[Dict[str, Any]]

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "Rule":
        return cls(
            name=data.get("name", "unnamed"),
            description=data.get("description", ""),
            enabled=data.get("enabled", True),
            priority=data.get("priority", 50),
            conditions=data.get("conditions", []),
            actions=data.get("actions", []),
        )


class RulesConfig:
    """
    Rules configuration for document processing.

    Loads and manages rules from rules.yaml file.
    """

    def __init__(
        self,
        version: str,
        settings: Dict[str, Any],
        checkbox_mappings: Dict[str, List[str]],
        field_mappings: Dict[str, List[str]],
        rules: List[Rule],
        default_field_values: Dict[str, str],
        post_processing: Dict[str, Any],
    ):
        self.version = version
        self.settings = settings
        self.checkbox_mappings = checkbox_mappings
        self.field_mappings = field_mappings
        self.rules = rules
        self.default_field_values = default_field_values
        self.post_processing = post_processing

    @classmethod
    def load(cls, rules_path: Optional[str] = None, base_path: Optional[Path] = None) -> "RulesConfig":
        """Load rules from YAML file."""
        if base_path is None:
            base_path = Path.cwd()

        if rules_path is None:
            rules_path = base_path / "config" / "rules.yaml"
        else:
            rules_path = Path(rules_path)
            if not rules_path.is_absolute():
                rules_path = base_path / rules_path

        if not rules_path.exists():
            # Return empty rules config if file doesn't exist
            return cls(
                version="1.0",
                settings={},
                checkbox_mappings={},
                field_mappings={},
                rules=[],
                default_field_values={},
                post_processing={},
            )

        with open(rules_path, "r", encoding="utf-8") as f:
            raw = yaml.safe_load(f) or {}

        rules = [Rule.from_dict(r) for r in raw.get("rules", [])]
        # Sort by priority (higher first)
        rules.sort(key=lambda r: r.priority, reverse=True)

        return cls(
            version=raw.get("version", "1.0"),
            settings=raw.get("settings", {}),
            checkbox_mappings=raw.get("checkbox_mappings", {}),
            field_mappings=raw.get("field_mappings", {}),
            rules=rules,
            default_field_values=raw.get("default_field_values", {}),
            post_processing=raw.get("post_processing", {}),
        )

    def get_enabled_rules(self) -> List[Rule]:
        """Get list of enabled rules, sorted by priority."""
        return [r for r in self.rules if r.enabled]

    def find_checkbox_tags(self, logical_name: str) -> List[str]:
        """Get all possible tags for a checkbox logical name."""
        return self.checkbox_mappings.get(logical_name, [logical_name])

    def find_field_tags(self, logical_name: str) -> List[str]:
        """Get all possible tags for a field logical name."""
        return self.field_mappings.get(logical_name, [logical_name])
