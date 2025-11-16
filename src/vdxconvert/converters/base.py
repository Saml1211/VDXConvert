#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Base converter class for VDXConvert.

Defines the interface that all file converters must implement.
"""

from abc import ABC, abstractmethod
from pathlib import Path
from typing import Optional


class BaseConverter(ABC):
    """Abstract base class for file converters."""

    @abstractmethod
    def can_convert(self, filepath: Path) -> bool:
        """
        Check if this converter can handle the given file.

        Args:
            filepath: Path to the file to check

        Returns:
            True if this converter can handle the file
        """
        pass

    @abstractmethod
    def convert(self, input_file: Path, output_file: Path) -> bool:
        """
        Convert a file to VDX format.

        Args:
            input_file: Path to input file
            output_file: Path where VDX output should be saved

        Returns:
            True if conversion was successful

        Raises:
            ConversionError: If conversion fails
            DependencyError: If required dependencies are missing
        """
        pass

    @abstractmethod
    def check_dependencies(self) -> bool:
        """
        Check if all dependencies required by this converter are available.

        Returns:
            True if all dependencies are available
        """
        pass

    @abstractmethod
    def get_missing_dependencies(self) -> list[str]:
        """
        Get a list of missing dependencies.

        Returns:
            List of missing dependency names
        """
        pass
