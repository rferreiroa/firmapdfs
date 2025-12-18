"""
Reporting Module
================

Handles generation of CSV and HTML reports for document processing.
"""

import csv
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional
from dataclasses import dataclass, field, asdict

from .logger import get_logger

logger = get_logger(__name__)


# =============================================================================
# Report Entry Data Class
# =============================================================================

@dataclass
class ReportEntry:
    """Single entry in the processing report."""

    timestamp: str = ""
    filename: str = ""
    status: str = ""  # "OK", "ERROR", "SKIPPED"
    message: str = ""
    docx_output: str = ""
    pdf_output: str = ""
    signed_output: str = ""
    duration_ms: int = 0

    # Additional metadata
    rules_applied: List[str] = field(default_factory=list)
    checkboxes_modified: List[str] = field(default_factory=list)
    fields_modified: List[str] = field(default_factory=list)

    @classmethod
    def create(
        cls,
        filename: str,
        status: str,
        message: str = "",
        **kwargs
    ) -> "ReportEntry":
        """Create a new report entry with current timestamp."""
        return cls(
            timestamp=datetime.now().isoformat(),
            filename=filename,
            status=status,
            message=message,
            **kwargs
        )

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for CSV/JSON export."""
        data = asdict(self)
        # Convert lists to comma-separated strings for CSV
        data["rules_applied"] = ",".join(self.rules_applied)
        data["checkboxes_modified"] = ",".join(self.checkboxes_modified)
        data["fields_modified"] = ",".join(self.fields_modified)
        return data


# =============================================================================
# Report Manager
# =============================================================================

class ReportManager:
    """
    Manages processing reports.

    Handles reading and writing CSV reports, with support for
    appending entries and generating summaries.
    """

    def __init__(
        self,
        csv_path: Path,
        columns: Optional[List[str]] = None,
        html_path: Optional[Path] = None,
    ):
        """
        Initialize report manager.

        Args:
            csv_path: Path to CSV report file
            columns: Column names for CSV (default: all ReportEntry fields)
            html_path: Optional path for HTML report
        """
        self.csv_path = Path(csv_path)
        self.html_path = Path(html_path) if html_path else None

        # Default columns (matching ReportEntry fields)
        self.columns = columns or [
            "timestamp",
            "filename",
            "status",
            "message",
            "docx_output",
            "pdf_output",
            "signed_output",
            "duration_ms",
        ]

        # Ensure parent directory exists
        self.csv_path.parent.mkdir(parents=True, exist_ok=True)

        # Create CSV file with headers if it doesn't exist
        if not self.csv_path.exists():
            self._write_headers()

    def _write_headers(self) -> None:
        """Write CSV headers to file."""
        with open(self.csv_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=self.columns)
            writer.writeheader()

    def add_entry(self, entry: ReportEntry) -> None:
        """
        Add a new entry to the report.

        Args:
            entry: ReportEntry to add
        """
        try:
            with open(self.csv_path, "a", newline="", encoding="utf-8") as f:
                writer = csv.DictWriter(f, fieldnames=self.columns)
                # Only write columns we care about
                row = {k: v for k, v in entry.to_dict().items() if k in self.columns}
                writer.writerow(row)

            logger.debug(f"Added report entry: {entry.filename} ({entry.status})")

        except Exception as e:
            logger.error(f"Failed to write report entry: {e}")

    def add_success(
        self,
        filename: str,
        docx_output: str = "",
        pdf_output: str = "",
        signed_output: str = "",
        duration_ms: int = 0,
        **kwargs
    ) -> None:
        """Add a successful processing entry."""
        entry = ReportEntry.create(
            filename=filename,
            status="OK",
            message="Processed successfully",
            docx_output=docx_output,
            pdf_output=pdf_output,
            signed_output=signed_output,
            duration_ms=duration_ms,
            **kwargs
        )
        self.add_entry(entry)

    def add_error(
        self,
        filename: str,
        error_message: str,
        duration_ms: int = 0,
        **kwargs
    ) -> None:
        """Add an error entry."""
        entry = ReportEntry.create(
            filename=filename,
            status="ERROR",
            message=error_message,
            duration_ms=duration_ms,
            **kwargs
        )
        self.add_entry(entry)

    def add_skipped(
        self,
        filename: str,
        reason: str = "Already processed",
        **kwargs
    ) -> None:
        """Add a skipped entry."""
        entry = ReportEntry.create(
            filename=filename,
            status="SKIPPED",
            message=reason,
            **kwargs
        )
        self.add_entry(entry)

    def get_entries(self, limit: Optional[int] = None) -> List[Dict[str, Any]]:
        """
        Read entries from the report.

        Args:
            limit: Maximum number of entries to return (from end)

        Returns:
            List of entry dictionaries
        """
        entries = []

        try:
            with open(self.csv_path, "r", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                entries = list(reader)

            if limit:
                entries = entries[-limit:]

        except FileNotFoundError:
            logger.warning(f"Report file not found: {self.csv_path}")
        except Exception as e:
            logger.error(f"Failed to read report: {e}")

        return entries

    def get_summary(self) -> Dict[str, Any]:
        """
        Get summary statistics from the report.

        Returns:
            Dictionary with summary statistics
        """
        entries = self.get_entries()

        summary = {
            "total": len(entries),
            "ok": 0,
            "error": 0,
            "skipped": 0,
            "avg_duration_ms": 0,
        }

        if not entries:
            return summary

        total_duration = 0
        ok_count = 0

        for entry in entries:
            status = entry.get("status", "").upper()
            if status == "OK":
                summary["ok"] += 1
                ok_count += 1
                try:
                    total_duration += int(entry.get("duration_ms", 0))
                except ValueError:
                    pass
            elif status == "ERROR":
                summary["error"] += 1
            elif status == "SKIPPED":
                summary["skipped"] += 1

        if ok_count > 0:
            summary["avg_duration_ms"] = total_duration // ok_count

        return summary

    def clear(self) -> None:
        """Clear all entries and reset the report."""
        self._write_headers()
        logger.info(f"Report cleared: {self.csv_path}")

    def generate_html(self) -> Optional[str]:
        """
        Generate HTML report.

        Returns:
            HTML content as string, or None if disabled
        """
        if not self.html_path:
            return None

        entries = self.get_entries()
        summary = self.get_summary()

        # Simple HTML template
        html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>FirmaPDFs Report</title>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 20px; }}
        h1 {{ color: #333; }}
        .summary {{ background: #f5f5f5; padding: 15px; margin-bottom: 20px; }}
        .summary span {{ margin-right: 20px; }}
        .ok {{ color: green; }}
        .error {{ color: red; }}
        .skipped {{ color: orange; }}
        table {{ border-collapse: collapse; width: 100%; }}
        th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
        th {{ background: #4CAF50; color: white; }}
        tr:nth-child(even) {{ background: #f2f2f2; }}
    </style>
</head>
<body>
    <h1>FirmaPDFs Processing Report</h1>
    <p>Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>

    <div class="summary">
        <strong>Summary:</strong>
        <span>Total: {summary['total']}</span>
        <span class="ok">OK: {summary['ok']}</span>
        <span class="error">Errors: {summary['error']}</span>
        <span class="skipped">Skipped: {summary['skipped']}</span>
        <span>Avg Duration: {summary['avg_duration_ms']}ms</span>
    </div>

    <table>
        <tr>
            <th>Timestamp</th>
            <th>Filename</th>
            <th>Status</th>
            <th>Message</th>
            <th>Duration (ms)</th>
        </tr>
"""

        for entry in entries:
            status_class = entry.get("status", "").lower()
            html += f"""        <tr>
            <td>{entry.get('timestamp', '')}</td>
            <td>{entry.get('filename', '')}</td>
            <td class="{status_class}">{entry.get('status', '')}</td>
            <td>{entry.get('message', '')}</td>
            <td>{entry.get('duration_ms', '')}</td>
        </tr>
"""

        html += """    </table>
</body>
</html>"""

        # Write HTML file
        try:
            with open(self.html_path, "w", encoding="utf-8") as f:
                f.write(html)
            logger.info(f"HTML report generated: {self.html_path}")
        except Exception as e:
            logger.error(f"Failed to generate HTML report: {e}")

        return html
