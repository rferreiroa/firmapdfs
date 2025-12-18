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

from .config import Config, RulesConfig
from .logger import setup_logging, get_logger
from .word_processor import (
    DocumentProcessor,
    DocxEngine,
    COMEngine,
    ProcessingResult,
    get_engine,
)

__all__ = [
    "__version__",
    "Config",
    "RulesConfig",
    "setup_logging",
    "get_logger",
    "DocumentProcessor",
    "DocxEngine",
    "COMEngine",
    "ProcessingResult",
    "get_engine",
]
