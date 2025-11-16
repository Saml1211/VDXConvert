#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Custom exception hierarchy for VDXConvert.

Provides specific exception types for better error handling and debugging.
"""


class VDXConvertError(Exception):
    """Base exception for all VDXConvert errors."""

    def __init__(self, message: str, details: str = None) -> None:
        """
        Initialize the exception.

        Args:
            message: Human-readable error message
            details: Optional additional details about the error
        """
        self.message = message
        self.details = details
        super().__init__(self.message)

    def __str__(self) -> str:
        if self.details:
            return f"{self.message}: {self.details}"
        return self.message


class ValidationError(VDXConvertError):
    """Raised when input validation fails."""
    pass


class ConversionError(VDXConvertError):
    """Raised when file conversion fails."""
    pass


class DependencyError(VDXConvertError):
    """Raised when required dependencies are missing or unavailable."""
    pass


class FileOperationError(VDXConvertError):
    """Raised when file operations (read, write, move) fail."""
    pass


class SecurityError(ValidationError):
    """Raised when security checks fail (e.g., path traversal, file size)."""
    pass
