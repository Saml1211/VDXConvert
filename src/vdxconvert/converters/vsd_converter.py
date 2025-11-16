#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
VSD/VDW file converter for VDXConvert.

Converts .vsd and .vdw files to VDX format using LibreOffice/unoconv.
"""

import logging
import os
import platform
import shutil
import subprocess
import traceback
from pathlib import Path
from typing import List, Optional

from ..config import (
    VSD_EXTENSIONS,
    TEMP_DIR,
    SOFFICE_CMD,
    SOFFICE_WINDOWS_DEFAULT_PATH,
    SOFFICE_TIMEOUT_SECONDS
)
from ..exceptions import ConversionError, DependencyError
from .base import BaseConverter

logger = logging.getLogger(__name__)


class VSDConverter(BaseConverter):
    """Converter for .vsd and .vdw files using LibreOffice/unoconv."""

    def __init__(self) -> None:
        """Initialize the VSD converter."""
        self._has_unoconv: Optional[bool] = None
        self._has_libreoffice: Optional[bool] = None
        self._soffice_path: Optional[str] = None

    def can_convert(self, filepath: Path) -> bool:
        """
        Check if this converter can handle the given file.

        Args:
            filepath: Path to the file to check

        Returns:
            True if file extension is .vsd or .vdw
        """
        return filepath.suffix.lower() in VSD_EXTENSIONS

    def check_dependencies(self) -> bool:
        """
        Check if LibreOffice or unoconv is available.

        Returns:
            True if either LibreOffice or unoconv is available
        """
        return self._check_libreoffice() or self._check_unoconv()

    def get_missing_dependencies(self) -> List[str]:
        """
        Get list of missing dependencies.

        Returns:
            List of missing dependencies
        """
        missing = []
        if not self._check_libreoffice():
            missing.append("LibreOffice")
        if not self._check_unoconv():
            missing.append("unoconv")
        return missing

    def _check_unoconv(self) -> bool:
        """
        Check if unoconv is available.

        Returns:
            True if unoconv is found
        """
        if self._has_unoconv is not None:
            return self._has_unoconv

        try:
            cmd = "where" if platform.system() == "Windows" else "which"
            result = subprocess.run(
                [cmd, "unoconv"],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                timeout=5
            )
            self._has_unoconv = result.returncode == 0
            if self._has_unoconv:
                logger.debug(f"unoconv found at: {result.stdout.strip()}")
        except (subprocess.SubprocessError, FileNotFoundError, subprocess.TimeoutExpired):
            self._has_unoconv = False

        return self._has_unoconv

    def _check_libreoffice(self) -> bool:
        """
        Check if LibreOffice is available.

        Returns:
            True if LibreOffice is found
        """
        if self._has_libreoffice is not None:
            return self._has_libreoffice

        try:
            if platform.system() == 'Windows':
                # Check Windows registry
                try:
                    import winreg
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\LibreOffice") as key:
                        logger.debug("LibreOffice found in Windows registry")
                        self._has_libreoffice = True
                        return True
                except (ImportError, WindowsError):
                    pass

                # Check default installation path
                if Path(SOFFICE_WINDOWS_DEFAULT_PATH).exists():
                    self._soffice_path = SOFFICE_WINDOWS_DEFAULT_PATH
                    self._has_libreoffice = True
                    logger.debug(f"LibreOffice found at: {self._soffice_path}")
                    return True

                self._has_libreoffice = False
            else:
                # Check for soffice on macOS/Linux
                result = subprocess.run(
                    ["which", "soffice"],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                    timeout=5
                )
                if result.returncode == 0 and result.stdout:
                    self._soffice_path = result.stdout.strip()
                    self._has_libreoffice = True
                    logger.debug(f"LibreOffice found at: {self._soffice_path}")
                else:
                    self._has_libreoffice = False

        except (subprocess.SubprocessError, FileNotFoundError, subprocess.TimeoutExpired) as e:
            logger.debug(f"Error checking for LibreOffice: {str(e)}")
            self._has_libreoffice = False

        return self._has_libreoffice

    def convert(self, input_file: Path, output_file: Path) -> bool:
        """
        Convert .vsd or .vdw file to .vdx format.

        Args:
            input_file: Path to input VSD/VDW file
            output_file: Path where VDX output should be saved

        Returns:
            True if conversion was successful

        Raises:
            DependencyError: If neither LibreOffice nor unoconv is available
            ConversionError: If conversion fails
        """
        if not self.check_dependencies():
            raise DependencyError(
                "Neither LibreOffice nor unoconv is available",
                "Install LibreOffice from libreoffice.org"
            )

        # Create temporary directory
        TEMP_DIR.mkdir(parents=True, exist_ok=True)

        try:
            # Try unoconv first (better quality)
            if self._check_unoconv():
                if self._convert_with_unoconv(input_file, output_file):
                    return True

            # Fall back to LibreOffice
            if self._check_libreoffice():
                if self._convert_with_libreoffice(input_file, output_file):
                    return True

            raise ConversionError(
                f"All conversion methods failed for {input_file.name}",
                "Both unoconv and LibreOffice conversion attempts failed"
            )

        finally:
            # Clean up temp directory
            if TEMP_DIR.exists():
                try:
                    shutil.rmtree(TEMP_DIR)
                except OSError:
                    pass  # Best effort cleanup

    def _convert_with_unoconv(self, input_file: Path, output_file: Path) -> bool:
        """
        Convert file using unoconv.

        Args:
            input_file: Input file path
            output_file: Output file path

        Returns:
            True if successful, False otherwise
        """
        try:
            logger.debug(f"Attempting conversion with unoconv: {input_file}")

            result = subprocess.run(
                ["unoconv", "-f", "vdx", "-o", str(TEMP_DIR), str(input_file)],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                timeout=SOFFICE_TIMEOUT_SECONDS
            )

            # Check if VDX file was created
            temp_vdx = TEMP_DIR / f"{input_file.stem}.vdx"
            if temp_vdx.exists():
                shutil.copy2(temp_vdx, output_file)
                temp_vdx.unlink()
                logger.debug(f"Successfully converted with unoconv: {input_file.name}")
                return True

            logger.debug(f"unoconv conversion failed: {result.stderr}")
            return False

        except (subprocess.SubprocessError, FileNotFoundError, subprocess.TimeoutExpired) as e:
            logger.debug(f"unoconv conversion error: {str(e)}")
            return False

    def _convert_with_libreoffice(self, input_file: Path, output_file: Path) -> bool:
        """
        Convert file using LibreOffice.

        Args:
            input_file: Input file path
            output_file: Output file path

        Returns:
            True if successful, False otherwise
        """
        try:
            logger.debug(f"Attempting conversion with LibreOffice: {input_file}")

            # Determine soffice command
            if platform.system() == "Windows" and self._soffice_path:
                soffice_cmd = f'"{self._soffice_path}"'
            else:
                soffice_cmd = SOFFICE_CMD

            convert_cmd = [
                soffice_cmd,
                "--headless",
                "--convert-to", "vdx",
                "--outdir", str(TEMP_DIR),
                str(input_file)
            ]

            result = subprocess.run(
                convert_cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                timeout=SOFFICE_TIMEOUT_SECONDS,
                shell=platform.system() == "Windows"  # Shell needed for quoted paths on Windows
            )

            # Check if VDX file was created
            temp_vdx = TEMP_DIR / f"{input_file.stem}.vdx"
            if temp_vdx.exists():
                shutil.copy2(temp_vdx, output_file)
                temp_vdx.unlink()
                logger.debug(f"Successfully converted with LibreOffice: {input_file.name}")
                return True

            logger.debug(f"LibreOffice conversion failed: {result.stderr}")
            return False

        except (subprocess.SubprocessError, FileNotFoundError, subprocess.TimeoutExpired) as e:
            logger.debug(f"LibreOffice conversion error: {str(e)}")
            return False
