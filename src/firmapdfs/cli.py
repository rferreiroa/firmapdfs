"""
Command Line Interface
======================

Main CLI entry point for FirmaPDFs using Click.
Provides commands for processing documents, watching folders,
and managing configuration.
"""

import sys
from datetime import datetime
from pathlib import Path
from typing import Optional

import click

from . import __version__
from .config import Config, RulesConfig
from .logger import setup_logging, get_logger, add_sensitive_filter


# =============================================================================
# CLI Context
# =============================================================================

class CLIContext:
    """Shared context for CLI commands."""

    def __init__(self):
        self.config: Optional[Config] = None
        self.rules: Optional[RulesConfig] = None
        self.logger = None
        self.verbose: bool = False
        self.dry_run: bool = False
        self._config_path: Optional[str] = None

    def setup(
        self,
        config_path: Optional[str] = None,
        verbose: bool = False,
        dry_run: bool = False,
    ) -> bool:
        """
        Initialize context with configuration.

        Returns:
            True if setup successful, False otherwise
        """
        self.verbose = verbose
        self.dry_run = dry_run
        self._config_path = config_path

        # Load configuration
        try:
            self.config = Config.load(config_path)
        except FileNotFoundError as e:
            click.echo(f"Error: {e}", err=True)
            click.echo("Run 'firmapdfs init' to create default configuration.", err=True)
            return False

        # Override dry_run from CLI
        if dry_run:
            self.config.processing.dry_run = True

        # Load rules
        self.rules = RulesConfig.load(base_path=self.config.base_path)

        # Setup logging
        log_level = "DEBUG" if verbose else self.config.logging.level
        self.logger = setup_logging(
            level=log_level,
            console=self.config.logging.console,
            file=self.config.logging.file,
            log_dir=self.config.paths.logs,
            file_pattern=self.config.logging.file_pattern,
            max_size_mb=self.config.logging.max_size_mb,
            backup_count=self.config.logging.backup_count,
            log_format=self.config.logging.format,
            date_format=self.config.logging.date_format,
        )

        # Add sensitive data filter
        add_sensitive_filter(self.logger)

        # Ensure directories exist
        self.config.ensure_directories()

        return True


pass_context = click.make_pass_decorator(CLIContext, ensure=True)


# =============================================================================
# Main CLI Group
# =============================================================================

@click.group(invoke_without_command=True)
@click.version_option(version=__version__, prog_name="FirmaPDFs")
@click.option(
    "-c", "--config",
    type=click.Path(exists=False),
    default=None,
    help="Path to configuration file (default: config/config.yaml)"
)
@click.option(
    "-v", "--verbose",
    is_flag=True,
    default=False,
    help="Enable verbose output (DEBUG level)"
)
@click.option(
    "--dry-run",
    is_flag=True,
    default=False,
    help="Simulate processing without making changes"
)
@click.pass_context
def cli(click_ctx, config: Optional[str], verbose: bool, dry_run: bool):
    """
    FirmaPDFs - Document Automation System

    Automates Word document processing, PDF export, and digital signing.

    \b
    Quick Start:
      1. Run 'firmapdfs init' to create default configuration
      2. Place .docx files in samples/inbox/
      3. Run 'firmapdfs process' to process all documents

    \b
    Examples:
      firmapdfs process                    Process all documents in inbox
      firmapdfs process --dry-run          Simulate processing
      firmapdfs process -f document.docx   Process a specific file
      firmapdfs watch                      Watch inbox for new files
      firmapdfs status                     Show system status
    """
    # Ensure CLIContext exists in Click context
    click_ctx.ensure_object(CLIContext)
    ctx = click_ctx.obj

    # Store options for later use
    ctx.verbose = verbose
    ctx.dry_run = dry_run
    ctx._config_path = config

    # Commands that don't need config
    no_config_commands = {"init"}

    # If no subcommand, show help
    if click_ctx.invoked_subcommand is None:
        click.echo(click_ctx.get_help())
        return

    # Load config for commands that need it
    if click_ctx.invoked_subcommand not in no_config_commands:
        if not ctx.setup(config, verbose, dry_run):
            # Setup failed - some commands can continue without config
            if click_ctx.invoked_subcommand not in {"status", "check"}:
                sys.exit(1)


# =============================================================================
# Init Command
# =============================================================================

@cli.command()
@click.option(
    "--force",
    is_flag=True,
    default=False,
    help="Overwrite existing configuration"
)
def init(force: bool):
    """
    Initialize project structure and configuration.

    Creates default configuration files and directory structure.
    """
    base_path = Path.cwd()

    # Define files to create
    config_dir = base_path / "config"
    config_file = config_dir / "config.yaml"
    rules_file = config_dir / "rules.yaml"

    # Check if files exist
    if config_file.exists() and not force:
        click.echo(f"Configuration already exists: {config_file}")
        click.echo("Use --force to overwrite.")
        return

    click.echo("Initializing FirmaPDFs project structure...")

    # Create directories
    directories = [
        "config",
        "samples/inbox",
        "samples/templates",
        "processed/docx",
        "processed/pdf",
        "processed/signed",
        "processed/archive",
        "logs",
        "tests",
    ]

    for dir_name in directories:
        dir_path = base_path / dir_name
        dir_path.mkdir(parents=True, exist_ok=True)
        click.echo(f"  Created: {dir_name}/")

    # Check if config files exist in the project
    if config_file.exists():
        click.echo(f"  Config exists: {config_file}")
    else:
        click.echo(f"  TODO: Copy default config to {config_file}")

    if rules_file.exists():
        click.echo(f"  Rules exists: {rules_file}")
    else:
        click.echo(f"  TODO: Copy default rules to {rules_file}")

    click.echo("")
    click.echo("Initialization complete!")
    click.echo("")
    click.echo("Next steps:")
    click.echo("  1. Edit config/config.yaml with your settings")
    click.echo("  2. Place your .pfx certificate in config/")
    click.echo("  3. Set certificate password (see README)")
    click.echo("  4. Place .docx files in samples/inbox/")
    click.echo("  5. Run: firmapdfs process")


# =============================================================================
# Process Command
# =============================================================================

@cli.command()
@click.option(
    "-f", "--file",
    type=click.Path(exists=True),
    default=None,
    help="Process a specific file instead of inbox"
)
@click.option(
    "--no-sign",
    is_flag=True,
    default=False,
    help="Skip PDF signing step"
)
@click.option(
    "--no-pdf",
    is_flag=True,
    default=False,
    help="Skip PDF export (only process Word)"
)
@pass_context
def process(ctx: CLIContext, file: Optional[str], no_sign: bool, no_pdf: bool):
    """
    Process documents from inbox folder.

    Processes .docx files: fills fields, exports to PDF, and signs.
    """
    logger = get_logger("cli.process")

    if ctx.config is None:
        click.echo("Error: Configuration not loaded", err=True)
        sys.exit(1)

    if ctx.dry_run:
        click.echo("[DRY RUN MODE] No changes will be made")
        click.echo("")

    # Determine files to process
    if file:
        files = [Path(file)]
        click.echo(f"Processing single file: {file}")
    else:
        inbox = ctx.config.paths.inbox
        files = list(inbox.glob("*.docx")) + list(inbox.glob("*.DOCX"))
        # Filter out temp files
        files = [f for f in files if not f.name.startswith("~$")]

        if not files:
            click.echo(f"No .docx files found in {inbox}")
            return

        click.echo(f"Found {len(files)} file(s) in {inbox}")

    click.echo("")

    # Process each file
    # TODO: Implement actual processing in Sprint 1+
    for filepath in files:
        click.echo(f"Processing: {filepath.name}")

        if ctx.dry_run:
            click.echo(f"  [DRY RUN] Would process: {filepath.name}")
            click.echo(f"  [DRY RUN] Would save docx to: {ctx.config.paths.output_docx}")
            if not no_pdf:
                click.echo(f"  [DRY RUN] Would export PDF to: {ctx.config.paths.output_pdf}")
            if not no_sign and not no_pdf:
                click.echo(f"  [DRY RUN] Would sign PDF to: {ctx.config.paths.output_signed}")
        else:
            # TODO: Implement processing
            click.echo(f"  [TODO] Processing not yet implemented")

        logger.info(f"Processed: {filepath.name}")

    click.echo("")
    click.echo(f"Processing complete. Processed {len(files)} file(s).")


# =============================================================================
# Watch Command
# =============================================================================

@cli.command()
@click.option(
    "--interval",
    type=int,
    default=None,
    help="Polling interval in seconds (overrides config)"
)
@pass_context
def watch(ctx: CLIContext, interval: Optional[int]):
    """
    Watch inbox folder for new documents.

    Monitors the inbox folder and automatically processes new .docx files.
    Press Ctrl+C to stop.
    """
    logger = get_logger("cli.watch")

    if ctx.config is None:
        click.echo("Error: Configuration not loaded", err=True)
        sys.exit(1)

    inbox = ctx.config.paths.inbox
    poll_interval = interval or ctx.config.watcher.poll_interval

    click.echo(f"Watching {inbox} for new .docx files...")
    click.echo(f"Poll interval: {poll_interval} seconds")
    click.echo("Press Ctrl+C to stop.")
    click.echo("")

    # TODO: Implement file watcher in Sprint 5
    click.echo("[TODO] File watcher not yet implemented")
    logger.info("Watch mode started")


# =============================================================================
# Status Command
# =============================================================================

@cli.command()
@pass_context
def status(ctx: CLIContext):
    """
    Show system status and configuration.

    Displays current configuration, directory status, and system info.
    """
    click.echo("=" * 60)
    click.echo("FirmaPDFs - System Status")
    click.echo("=" * 60)
    click.echo("")

    # Version info
    click.echo(f"Version: {__version__}")
    click.echo(f"Python: {sys.version.split()[0]}")
    click.echo(f"Platform: {sys.platform}")
    click.echo("")

    if ctx.config is None:
        click.echo("Configuration: NOT LOADED")
        click.echo("Run 'firmapdfs init' to create configuration.")
        return

    # Configuration
    click.echo("Configuration:")
    click.echo(f"  Base path: {ctx.config.base_path}")
    click.echo(f"  Word engine: {ctx.config.word.engine}")
    click.echo(f"  Dry run: {ctx.config.processing.dry_run}")
    click.echo("")

    # Directories
    click.echo("Directories:")
    directories = {
        "Inbox": ctx.config.paths.inbox,
        "Output (docx)": ctx.config.paths.output_docx,
        "Output (pdf)": ctx.config.paths.output_pdf,
        "Output (signed)": ctx.config.paths.output_signed,
        "Logs": ctx.config.paths.logs,
    }

    for name, path in directories.items():
        exists = "OK" if path.exists() else "MISSING"
        file_count = len(list(path.glob("*"))) if path.exists() else 0
        click.echo(f"  {name}: {path} [{exists}] ({file_count} files)")

    click.echo("")

    # Certificate
    click.echo("Certificate:")
    cert_path = ctx.config.signature.certificate_path
    cert_exists = cert_path.exists()
    click.echo(f"  Path: {cert_path}")
    click.echo(f"  Status: {'OK' if cert_exists else 'NOT FOUND'}")
    click.echo(f"  Password storage: {ctx.config.signature.password_storage}")
    click.echo("")

    # Rules
    click.echo("Rules:")
    if ctx.rules:
        enabled_rules = len(ctx.rules.get_enabled_rules())
        total_rules = len(ctx.rules.rules)
        click.echo(f"  Loaded: {enabled_rules}/{total_rules} rules enabled")
        click.echo(f"  Checkbox mappings: {len(ctx.rules.checkbox_mappings)}")
        click.echo(f"  Field mappings: {len(ctx.rules.field_mappings)}")
    else:
        click.echo("  Rules: NOT LOADED")

    click.echo("")

    # Validation
    click.echo("Validation:")
    issues = ctx.config.validate()
    if issues:
        for issue in issues:
            click.echo(f"  WARNING: {issue}")
    else:
        click.echo("  All checks passed")

    click.echo("")
    click.echo("=" * 60)


# =============================================================================
# Check Command
# =============================================================================

@cli.command()
@click.option(
    "--fix",
    is_flag=True,
    default=False,
    help="Attempt to fix issues (create directories)"
)
@pass_context
def check(ctx: CLIContext, fix: bool):
    """
    Validate configuration and dependencies.

    Checks that all required components are properly configured.
    """
    click.echo("Checking FirmaPDFs configuration...")
    click.echo("")

    issues = []
    warnings = []

    # Check Python version
    py_version = sys.version_info
    if py_version < (3, 8):
        issues.append(f"Python 3.8+ required, found {py_version.major}.{py_version.minor}")
    else:
        click.echo(f"[OK] Python version: {py_version.major}.{py_version.minor}.{py_version.micro}")

    # Check required packages
    packages = [
        ("click", "CLI framework"),
        ("yaml", "YAML configuration"),
        ("docx", "Word document processing"),
    ]

    for package, description in packages:
        try:
            __import__(package)
            click.echo(f"[OK] {package}: {description}")
        except ImportError:
            issues.append(f"Missing package: {package} ({description})")

    # Check optional packages
    optional_packages = [
        ("win32com.client", "Word COM automation (Windows only)"),
        ("colorama", "Colored console output"),
        ("watchdog", "File system watcher"),
        ("endesive", "PDF signing"),
    ]

    for package, description in optional_packages:
        try:
            __import__(package)
            click.echo(f"[OK] {package}: {description}")
        except ImportError:
            warnings.append(f"Optional package not installed: {package} ({description})")

    click.echo("")

    # Check configuration
    if ctx.config:
        click.echo("Checking configuration...")
        config_issues = ctx.config.validate()
        issues.extend(config_issues)

        # Check/create directories
        if fix:
            click.echo("Creating missing directories...")
            ctx.config.ensure_directories()
            click.echo("[OK] Directories created")

    click.echo("")

    # Summary
    if issues:
        click.echo("ISSUES FOUND:")
        for issue in issues:
            click.echo(f"  [ERROR] {issue}")
    else:
        click.echo("[OK] No critical issues found")

    if warnings:
        click.echo("")
        click.echo("WARNINGS:")
        for warning in warnings:
            click.echo(f"  [WARN] {warning}")

    click.echo("")

    if issues:
        sys.exit(1)


# =============================================================================
# Set Password Command
# =============================================================================

@cli.command("set-password")
@click.option(
    "--method",
    type=click.Choice(["credential_manager", "dpapi", "env"]),
    default="credential_manager",
    help="Password storage method"
)
@pass_context
def set_password(ctx: CLIContext, method: str):
    """
    Set certificate password securely.

    Stores the certificate password using the specified method.
    """
    click.echo("Setting certificate password...")
    click.echo(f"Storage method: {method}")
    click.echo("")

    # Prompt for password (hidden input)
    password = click.prompt("Certificate password", hide_input=True, confirmation_prompt=True)

    if method == "credential_manager":
        # TODO: Implement Windows Credential Manager storage
        click.echo("[TODO] Credential Manager storage not yet implemented")
        click.echo("For now, set environment variable FIRMAPDFS_CERT_PASSWORD")
    elif method == "dpapi":
        # TODO: Implement DPAPI storage
        click.echo("[TODO] DPAPI storage not yet implemented")
    elif method == "env":
        click.echo("To use environment variable, set:")
        click.echo("  set FIRMAPDFS_CERT_PASSWORD=<your_password>")
        click.echo("")
        click.echo("Or add to your shell profile for persistence.")

    click.echo("")
    click.echo("Password configuration complete.")


# =============================================================================
# Report Command
# =============================================================================

@cli.command()
@click.option(
    "--format",
    "output_format",  # Renamed to avoid shadowing builtin
    type=click.Choice(["table", "csv", "json"]),
    default="table",
    help="Output format"
)
@click.option(
    "--last",
    type=int,
    default=10,
    help="Show last N entries"
)
@pass_context
def report(ctx: CLIContext, output_format: str, last: int):
    """
    Show processing report.

    Displays summary of processed documents.
    """
    click.echo("Processing Report")
    click.echo("=" * 60)

    if ctx.config is None:
        click.echo("Error: Configuration not loaded", err=True)
        return

    # TODO: Implement report reading
    click.echo("[TODO] Report functionality not yet implemented")
    click.echo(f"Report file: {ctx.config.reporting.csv_path}")


# =============================================================================
# Entry Point
# =============================================================================

def main():
    """Main entry point for the CLI."""
    try:
        cli(obj=CLIContext())
    except Exception as e:
        click.echo(f"Error: {e}", err=True)
        sys.exit(1)


if __name__ == "__main__":
    main()
