"""
FirmaPDFs - Document Automation System
======================================

Automated system for processing Word documents, exporting to PDF,
and digitally signing with a certificate.

Author: FirmaPDFs Team
Version: 1.0.0
License: MIT
"""

__version__ = "1.0.0"
__author__ = "FirmaPDFs Team"

from .config import Config
from .logger import setup_logging, get_logger

__all__ = [
    "__version__",
    "Config",
    "setup_logging",
    "get_logger",
]
