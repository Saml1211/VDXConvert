#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Tests for converter modules.
"""

import pytest
from pathlib import Path
from unittest.mock import Mock, patch, MagicMock
from src.vdxconvert.converters import VSDXConverter, VSDConverter
from src.vdxconvert.exceptions import ConversionError, DependencyError


class TestVSDXConverter:
    """Tests for VSDXConverter class."""

    def test_can_convert_vsdx(self):
        """Test that converter recognizes .vsdx files."""
        converter = VSDXConverter()
        assert converter.can_convert(Path("test.vsdx")) is True

    def test_can_convert_vsdm(self):
        """Test that converter recognizes .vsdm files."""
        converter = VSDXConverter()
        assert converter.can_convert(Path("test.vsdm")) is True

    def test_cannot_convert_vsd(self):
        """Test that converter rejects .vsd files."""
        converter = VSDXConverter()
        assert converter.can_convert(Path("test.vsd")) is False

    def test_cannot_convert_other_formats(self):
        """Test that converter rejects other file types."""
        converter = VSDXConverter()
        assert converter.can_convert(Path("test.txt")) is False
        assert converter.can_convert(Path("test.pdf")) is False

    def test_case_insensitive_extension(self):
        """Test that extension check is case-insensitive."""
        converter = VSDXConverter()
        assert converter.can_convert(Path("test.VSDX")) is True
        assert converter.can_convert(Path("test.VsDm")) is True

    @patch('src.vdxconvert.converters.vsdx_converter.VSDX_AVAILABLE', True)
    def test_check_dependencies_available(self):
        """Test dependency check when vsdx is available."""
        converter = VSDXConverter()
        assert converter.check_dependencies() is True
        assert converter.get_missing_dependencies() == []

    @patch('src.vdxconvert.converters.vsdx_converter.VSDX_AVAILABLE', False)
    def test_check_dependencies_missing(self):
        """Test dependency check when vsdx is missing."""
        converter = VSDXConverter()
        assert converter.check_dependencies() is False
        assert 'vsdx' in converter.get_missing_dependencies()

    @patch('src.vdxconvert.converters.vsdx_converter.VSDX_AVAILABLE', False)
    def test_convert_without_dependency(self, sample_vsdx_file, mock_output_dir):
        """Test that conversion fails without vsdx library."""
        converter = VSDXConverter()
        output = mock_output_dir / "test.vdx"

        with pytest.raises(DependencyError) as exc_info:
            converter.convert(sample_vsdx_file, output)
        assert "vsdx library not available" in str(exc_info.value)

    @patch('src.vdxconvert.converters.vsdx_converter.VSDX_AVAILABLE', True)
    @patch('src.vdxconvert.converters.vsdx_converter.vsdx')
    def test_convert_success(self, mock_vsdx, sample_vsdx_file, mock_output_dir):
        """Test successful conversion."""
        # Mock the vsdx.VisioFile
        mock_drawing = MagicMock()
        mock_page = MagicMock()
        mock_page.name = "Page1"
        mock_page.width = 8.5
        mock_page.height = 11
        mock_page.shapes = []
        mock_drawing.pages = [mock_page]
        mock_vsdx.VisioFile.return_value = mock_drawing

        converter = VSDXConverter()
        output = mock_output_dir / "test.vdx"

        result = converter.convert(sample_vsdx_file, output)

        assert result is True
        assert output.exists()

    @patch('src.vdxconvert.converters.vsdx_converter.VSDX_AVAILABLE', True)
    @patch('src.vdxconvert.converters.vsdx_converter.vsdx')
    def test_convert_failure(self, mock_vsdx, sample_vsdx_file, mock_output_dir):
        """Test conversion failure handling."""
        mock_vsdx.VisioFile.side_effect = Exception("Test error")

        converter = VSDXConverter()
        output = mock_output_dir / "test.vdx"

        with pytest.raises(ConversionError) as exc_info:
            converter.convert(sample_vsdx_file, output)
        assert "Failed to convert" in str(exc_info.value)


class TestVSDConverter:
    """Tests for VSDConverter class."""

    def test_can_convert_vsd(self):
        """Test that converter recognizes .vsd files."""
        converter = VSDConverter()
        assert converter.can_convert(Path("test.vsd")) is True

    def test_can_convert_vdw(self):
        """Test that converter recognizes .vdw files."""
        converter = VSDConverter()
        assert converter.can_convert(Path("test.vdw")) is True

    def test_cannot_convert_vsdx(self):
        """Test that converter rejects .vsdx files."""
        converter = VSDConverter()
        assert converter.can_convert(Path("test.vsdx")) is False

    def test_cannot_convert_other_formats(self):
        """Test that converter rejects other file types."""
        converter = VSDConverter()
        assert converter.can_convert(Path("test.txt")) is False
        assert converter.can_convert(Path("test.pdf")) is False

    def test_case_insensitive_extension(self):
        """Test that extension check is case-insensitive."""
        converter = VSDConverter()
        assert converter.can_convert(Path("test.VSD")) is True
        assert converter.can_convert(Path("test.VdW")) is True

    @patch.object(VSDConverter, '_check_unoconv', return_value=True)
    @patch.object(VSDConverter, '_check_libreoffice', return_value=False)
    def test_check_dependencies_unoconv_only(self, mock_lo, mock_uno):
        """Test dependency check with only unoconv available."""
        converter = VSDConverter()
        assert converter.check_dependencies() is True

    @patch.object(VSDConverter, '_check_unoconv', return_value=False)
    @patch.object(VSDConverter, '_check_libreoffice', return_value=True)
    def test_check_dependencies_libreoffice_only(self, mock_lo, mock_uno):
        """Test dependency check with only LibreOffice available."""
        converter = VSDConverter()
        assert converter.check_dependencies() is True

    @patch.object(VSDConverter, '_check_unoconv', return_value=False)
    @patch.object(VSDConverter, '_check_libreoffice', return_value=False)
    def test_check_dependencies_none(self, mock_lo, mock_uno):
        """Test dependency check with no dependencies."""
        converter = VSDConverter()
        assert converter.check_dependencies() is False
        deps = converter.get_missing_dependencies()
        assert 'LibreOffice' in deps
        assert 'unoconv' in deps

    @patch.object(VSDConverter, '_check_unoconv', return_value=False)
    @patch.object(VSDConverter, '_check_libreoffice', return_value=False)
    def test_convert_without_dependencies(self, mock_lo, mock_uno, sample_vsd_file, mock_output_dir):
        """Test that conversion fails without dependencies."""
        converter = VSDConverter()
        output = mock_output_dir / "test.vdx"

        with pytest.raises(DependencyError) as exc_info:
            converter.convert(sample_vsd_file, output)
        assert "Neither LibreOffice nor unoconv" in str(exc_info.value)

    @patch.object(VSDConverter, 'check_dependencies', return_value=True)
    @patch.object(VSDConverter, '_check_unoconv', return_value=True)
    @patch.object(VSDConverter, '_convert_with_unoconv', return_value=True)
    def test_convert_with_unoconv_success(self, mock_convert, mock_check, mock_deps,
                                         sample_vsd_file, mock_output_dir):
        """Test successful conversion with unoconv."""
        converter = VSDConverter()
        output = mock_output_dir / "test.vdx"

        result = converter.convert(sample_vsd_file, output)

        assert result is True
        mock_convert.assert_called_once()

    @patch.object(VSDConverter, 'check_dependencies', return_value=True)
    @patch.object(VSDConverter, '_check_unoconv', return_value=False)
    @patch.object(VSDConverter, '_check_libreoffice', return_value=True)
    @patch.object(VSDConverter, '_convert_with_libreoffice', return_value=True)
    def test_convert_with_libreoffice_success(self, mock_convert, mock_check_lo,
                                              mock_check_uno, mock_deps,
                                              sample_vsd_file, mock_output_dir):
        """Test successful conversion with LibreOffice."""
        converter = VSDConverter()
        output = mock_output_dir / "test.vdx"

        result = converter.convert(sample_vsd_file, output)

        assert result is True
        mock_convert.assert_called_once()

    @patch.object(VSDConverter, 'check_dependencies', return_value=True)
    @patch.object(VSDConverter, '_check_unoconv', return_value=True)
    @patch.object(VSDConverter, '_check_libreoffice', return_value=True)
    @patch.object(VSDConverter, '_convert_with_unoconv', return_value=False)
    @patch.object(VSDConverter, '_convert_with_libreoffice', return_value=True)
    def test_convert_fallback_to_libreoffice(self, mock_lo, mock_uno,
                                             mock_check_lo, mock_check_uno, mock_deps,
                                             sample_vsd_file, mock_output_dir):
        """Test fallback to LibreOffice when unoconv fails."""
        converter = VSDConverter()
        output = mock_output_dir / "test.vdx"

        result = converter.convert(sample_vsd_file, output)

        assert result is True
        mock_uno.assert_called_once()
        mock_lo.assert_called_once()

    @patch.object(VSDConverter, 'check_dependencies', return_value=True)
    @patch.object(VSDConverter, '_check_unoconv', return_value=True)
    @patch.object(VSDConverter, '_check_libreoffice', return_value=True)
    @patch.object(VSDConverter, '_convert_with_unoconv', return_value=False)
    @patch.object(VSDConverter, '_convert_with_libreoffice', return_value=False)
    def test_convert_all_methods_fail(self, mock_lo, mock_uno,
                                      mock_check_lo, mock_check_uno, mock_deps,
                                      sample_vsd_file, mock_output_dir):
        """Test that exception is raised when all methods fail."""
        converter = VSDConverter()
        output = mock_output_dir / "test.vdx"

        with pytest.raises(ConversionError) as exc_info:
            converter.convert(sample_vsd_file, output)
        assert "All conversion methods failed" in str(exc_info.value)
