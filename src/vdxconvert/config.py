#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Configuration module for VDXConvert.

Contains all application constants, defaults, and configuration values.
"""

import os
from pathlib import Path
from typing import Final, List

# Application metadata
APP_NAME: Final[str] = "VDXConvert"
VERSION: Final[str] = "1.0.0"
AUTHOR: Final[str] = "Sam Lyndon"
LICENSE: Final[str] = "MIT"

# Supported file extensions
SUPPORTED_EXTENSIONS: Final[List[str]] = ['.vsd', '.vsdx', '.vsdm', '.vdw']
VSDX_EXTENSIONS: Final[List[str]] = ['.vsdx', '.vsdm']
VSD_EXTENSIONS: Final[List[str]] = ['.vsd', '.vdw']
OUTPUT_EXTENSION: Final[str] = '.vdx'

# Directory paths
SCRIPT_DIR: Final[Path] = Path(__file__).parent.parent.parent.resolve()
INPUT_DIR: Final[Path] = SCRIPT_DIR / "input"
OUTPUT_DIR: Final[Path] = SCRIPT_DIR / "output"
ARCHIVE_DIR: Final[Path] = SCRIPT_DIR / "archive"
LOGS_DIR: Final[Path] = SCRIPT_DIR / "logs"
TEMP_DIR: Final[Path] = LOGS_DIR / "temp"

# File limits (security)
MAX_FILE_SIZE_MB: Final[int] = 500  # Maximum file size in MB
MAX_FILENAME_LENGTH: Final[int] = 255  # Maximum filename length

# Logging configuration
LOG_FILE_NAME: Final[str] = f"{APP_NAME}.log"
LOG_FORMAT_FILE: Final[str] = '%(asctime)s - %(levelname)s - %(message)s'
LOG_FORMAT_CONSOLE: Final[str] = '%(message)s'

# VDX XML namespace
VDX_NAMESPACE: Final[str] = "http://schemas.microsoft.com/visio/2003/core"

# LibreOffice/soffice configuration
SOFFICE_CMD: Final[str] = "soffice"
SOFFICE_WINDOWS_DEFAULT_PATH: Final[str] = r"C:\Program Files\LibreOffice\program\soffice.exe"
SOFFICE_TIMEOUT_SECONDS: Final[int] = 300  # 5 minutes

# CSV report configuration
CSV_FIELDNAMES: Final[List[str]] = ["filename", "output", "archive", "success", "time", "error"]

# User prompts
PROMPT_SAVE_CSV: Final[str] = "Save detailed CSV report? [Y/n]: "
