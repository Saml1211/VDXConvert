#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Utility functions for VDXConvert.

Provides file operation helpers, unique filename generation,
and other common utilities.
"""

import csv
import shutil
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Any

from .config import SUPPORTED_EXTENSIONS, CSV_FIELDNAMES
from .exceptions import FileOperationError


def get_visio_files(directory: Path) -> List[Path]:
    """
    Get a list of Visio files in the given directory.

    Args:
        directory: Directory to search for Visio files

    Returns:
        List of Path objects for Visio files

    Example:
        >>> files = get_visio_files(Path("input"))
        >>> print([f.name for f in files])
        ['diagram1.vsdx', 'diagram2.vsd']
    """
    files = []
    for filepath in directory.iterdir():
        if filepath.is_file():
            ext = filepath.suffix.lower()
            if ext in SUPPORTED_EXTENSIONS:
                files.append(filepath)
    return files


def get_unique_filename(filepath: Path) -> Path:
    """
    Generate a unique filename if file already exists.

    Appends _1, _2, etc. to the filename until a unique name is found.

    Args:
        filepath: Desired file path

    Returns:
        Unique file path

    Example:
        >>> path = get_unique_filename(Path("output/file.vdx"))
        >>> print(path)
        output/file.vdx  # or output/file_1.vdx if it exists
    """
    if not filepath.exists():
        return filepath

    directory = filepath.parent
    stem = filepath.stem
    suffix = filepath.suffix

    counter = 1
    while True:
        new_filepath = directory / f"{stem}_{counter}{suffix}"
        if not new_filepath.exists():
            return new_filepath
        counter += 1


def safe_move_file(src: Path, dst: Path) -> Path:
    """
    Safely move a file to a destination.

    Args:
        src: Source file path
        dst: Destination file path

    Returns:
        Final destination path (may be modified for uniqueness)

    Raises:
        FileOperationError: If move operation fails
    """
    try:
        # Ensure destination is unique
        unique_dst = get_unique_filename(dst)

        # Move the file
        shutil.move(str(src), str(unique_dst))

        return unique_dst
    except (OSError, shutil.Error) as e:
        raise FileOperationError(f"Failed to move file: {src} -> {dst}", str(e))


def safe_copy_file(src: Path, dst: Path) -> Path:
    """
    Safely copy a file to a destination.

    Args:
        src: Source file path
        dst: Destination file path

    Returns:
        Final destination path (may be modified for uniqueness)

    Raises:
        FileOperationError: If copy operation fails
    """
    try:
        # Ensure destination is unique
        unique_dst = get_unique_filename(dst)

        # Copy the file with metadata
        shutil.copy2(str(src), str(unique_dst))

        return unique_dst
    except (OSError, shutil.Error) as e:
        raise FileOperationError(f"Failed to copy file: {src} -> {dst}", str(e))


def save_csv_report(results: List[Dict[str, Any]], save_path: Path) -> None:
    """
    Save analysis report to CSV file.

    Args:
        results: List of result dictionaries
        save_path: Path where CSV file should be saved

    Raises:
        FileOperationError: If CSV write fails

    Example:
        >>> results = [{"filename": "test.vsdx", "success": True, ...}]
        >>> save_csv_report(results, Path("logs/report.csv"))
    """
    try:
        with open(save_path, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=CSV_FIELDNAMES)
            writer.writeheader()
            for result in results:
                writer.writerow(result)
    except (OSError, csv.Error) as e:
        raise FileOperationError(f"Failed to write CSV report: {save_path}", str(e))


def generate_report_filename() -> str:
    """
    Generate a timestamped filename for conversion reports.

    Returns:
        Filename string with timestamp

    Example:
        >>> filename = generate_report_filename()
        >>> print(filename)
        conversion_report_20250116_143022.csv
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"conversion_report_{timestamp}.csv"


def ensure_directories_exist(directories: List[Path]) -> None:
    """
    Ensure all required directories exist.

    Args:
        directories: List of directory paths to create

    Raises:
        FileOperationError: If directory creation fails
    """
    for directory in directories:
        try:
            directory.mkdir(parents=True, exist_ok=True)
        except OSError as e:
            raise FileOperationError(f"Failed to create directory: {directory}", str(e))
