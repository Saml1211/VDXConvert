#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Input validation and security checks for VDXConvert.

Provides validation for file paths, extensions, and security constraints
to prevent path traversal, oversized files, and other security issues.
"""

import os
from pathlib import Path
from typing import List

from .config import (
    SUPPORTED_EXTENSIONS,
    MAX_FILE_SIZE_MB,
    MAX_FILENAME_LENGTH,
    INPUT_DIR,
    OUTPUT_DIR,
    ARCHIVE_DIR
)
from .exceptions import ValidationError, SecurityError


def validate_file_extension(filepath: Path, allowed_extensions: List[str] = None) -> bool:
    """
    Validate that file has an allowed extension.

    Args:
        filepath: Path to the file
        allowed_extensions: List of allowed extensions (default: SUPPORTED_EXTENSIONS)

    Returns:
        True if extension is valid

    Raises:
        ValidationError: If extension is not allowed

    Example:
        >>> validate_file_extension(Path("test.vsdx"))
        True
    """
    if allowed_extensions is None:
        allowed_extensions = SUPPORTED_EXTENSIONS

    ext = filepath.suffix.lower()
    if ext not in allowed_extensions:
        raise ValidationError(
            f"Invalid file extension: {ext}",
            f"Allowed extensions: {', '.join(allowed_extensions)}"
        )
    return True


def validate_file_exists(filepath: Path) -> bool:
    """
    Validate that file exists and is a regular file.

    Args:
        filepath: Path to the file

    Returns:
        True if file exists

    Raises:
        ValidationError: If file doesn't exist or is not a regular file
    """
    if not filepath.exists():
        raise ValidationError(f"File not found: {filepath}")

    if not filepath.is_file():
        raise ValidationError(f"Path is not a file: {filepath}")

    return True


def validate_file_size(filepath: Path, max_size_mb: int = MAX_FILE_SIZE_MB) -> bool:
    """
    Validate that file size is within limits.

    Args:
        filepath: Path to the file
        max_size_mb: Maximum allowed file size in megabytes

    Returns:
        True if file size is valid

    Raises:
        SecurityError: If file size exceeds limit
    """
    file_size_mb = filepath.stat().st_size / (1024 * 1024)

    if file_size_mb > max_size_mb:
        raise SecurityError(
            f"File size exceeds limit: {file_size_mb:.2f}MB",
            f"Maximum allowed: {max_size_mb}MB"
        )

    return True


def validate_filename(filename: str, max_length: int = MAX_FILENAME_LENGTH) -> bool:
    """
    Validate filename for security issues.

    Args:
        filename: The filename to validate
        max_length: Maximum allowed filename length

    Returns:
        True if filename is valid

    Raises:
        ValidationError: If filename is invalid
        SecurityError: If filename contains suspicious patterns
    """
    # Check length
    if len(filename) > max_length:
        raise ValidationError(
            f"Filename too long: {len(filename)} characters",
            f"Maximum: {max_length} characters"
        )

    # Check for null bytes
    if '\0' in filename:
        raise SecurityError("Filename contains null bytes")

    # Check for path traversal attempts
    if '..' in filename or filename.startswith('/') or filename.startswith('\\'):
        raise SecurityError(f"Potential path traversal in filename: {filename}")

    # Check for suspicious characters (OS-specific)
    suspicious_chars = ['<', '>', ':', '"', '|', '?', '*']
    for char in suspicious_chars:
        if char in filename:
            raise SecurityError(f"Suspicious character in filename: {char}")

    return True


def validate_path_within_directory(filepath: Path, allowed_dir: Path) -> bool:
    """
    Validate that a path is within an allowed directory (prevents path traversal).

    Args:
        filepath: Path to validate
        allowed_dir: Directory that filepath must be within

    Returns:
        True if path is safe

    Raises:
        SecurityError: If path is outside allowed directory
    """
    try:
        # Resolve both paths to absolute paths
        resolved_file = filepath.resolve()
        resolved_dir = allowed_dir.resolve()

        # Check if file is within the allowed directory
        if not str(resolved_file).startswith(str(resolved_dir)):
            raise SecurityError(
                f"Path traversal detected: {filepath}",
                f"Must be within: {allowed_dir}"
            )

        return True
    except (OSError, RuntimeError) as e:
        raise SecurityError(f"Invalid path: {filepath}", str(e))


def validate_input_file(filepath: Path) -> bool:
    """
    Comprehensive validation for input files.

    Args:
        filepath: Path to the input file

    Returns:
        True if all validations pass

    Raises:
        ValidationError: If validation fails
        SecurityError: If security check fails
    """
    # Validate filename
    validate_filename(filepath.name)

    # Validate file exists
    validate_file_exists(filepath)

    # Validate extension
    validate_file_extension(filepath)

    # Validate file size
    validate_file_size(filepath)

    # Validate path is within input directory
    validate_path_within_directory(filepath, INPUT_DIR)

    return True


def sanitize_path(path_str: str, base_dir: Path) -> Path:
    """
    Sanitize a path string and return a safe Path object.

    Args:
        path_str: The path string to sanitize
        base_dir: Base directory for the path

    Returns:
        Sanitized Path object

    Raises:
        SecurityError: If path contains null bytes or is otherwise unsafe
    """
    # Check for null bytes and reject immediately
    if '\0' in path_str:
        raise SecurityError("Path contains null bytes")

    # Create Path object
    path = Path(path_str)

    # Get just the filename (removes any directory components)
    safe_name = path.name

    # Validate the filename
    validate_filename(safe_name)

    # Return path within base directory
    return base_dir / safe_name
