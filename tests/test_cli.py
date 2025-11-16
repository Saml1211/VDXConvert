#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Tests for CLI module.
"""

import pytest
from unittest.mock import Mock, patch, MagicMock
from pathlib import Path
from src.vdxconvert.cli import VDXConvertCLI, parse_arguments
from src.vdxconvert.exceptions import ConversionError, ValidationError


class TestParseArguments:
    """Tests for argument parsing."""

    def test_default_arguments(self):
        """Test parsing with no arguments."""
        with patch('sys.argv', ['vdxconvert']):
            args = parse_arguments()
            assert args.verbose is False
            assert args.no_report is False

    def test_verbose_flag(self):
        """Test parsing with verbose flag."""
        with patch('sys.argv', ['vdxconvert', '-v']):
            args = parse_arguments()
            assert args.verbose is True

        with patch('sys.argv', ['vdxconvert', '--verbose']):
            args = parse_arguments()
            assert args.verbose is True

    def test_no_report_flag(self):
        """Test parsing with no-report flag."""
        with patch('sys.argv', ['vdxconvert', '--no-report']):
            args = parse_arguments()
            assert args.no_report is True

    def test_combined_flags(self):
        """Test parsing with multiple flags."""
        with patch('sys.argv', ['vdxconvert', '-v', '--no-report']):
            args = parse_arguments()
            assert args.verbose is True
            assert args.no_report is True


class TestVDXConvertCLI:
    """Tests for VDXConvertCLI class."""

    def test_initialization(self):
        """Test CLI initialization."""
        cli = VDXConvertCLI()
        assert len(cli.converters) == 2
        assert cli.converters is not None

    def test_get_converter_for_vsdx(self):
        """Test getting converter for .vsdx file."""
        cli = VDXConvertCLI()
        filepath = Path("test.vsdx")

        with patch.object(cli.converters[0], 'can_convert', return_value=True):
            with patch.object(cli.converters[0], 'check_dependencies', return_value=True):
                converter = cli.get_converter_for_file(filepath)
                assert converter is not None

    def test_get_converter_for_vsd(self):
        """Test getting converter for .vsd file."""
        cli = VDXConvertCLI()
        filepath = Path("test.vsd")

        with patch.object(cli.converters[1], 'can_convert', return_value=True):
            with patch.object(cli.converters[1], 'check_dependencies', return_value=True):
                converter = cli.get_converter_for_file(filepath)
                assert converter is not None

    def test_get_converter_no_dependencies(self):
        """Test that None is returned when dependencies missing."""
        cli = VDXConvertCLI()
        filepath = Path("test.vsdx")

        with patch.object(cli.converters[0], 'can_convert', return_value=True):
            with patch.object(cli.converters[0], 'check_dependencies', return_value=False):
                with patch.object(cli.converters[1], 'can_convert', return_value=False):
                    converter = cli.get_converter_for_file(filepath)
                    assert converter is None

    def test_get_converter_unsupported_format(self):
        """Test that None is returned for unsupported formats."""
        cli = VDXConvertCLI()
        filepath = Path("test.txt")

        converter = cli.get_converter_for_file(filepath)
        assert converter is None

    @patch('src.vdxconvert.cli.validate_input_file')
    @patch('src.vdxconvert.cli.safe_move_file')
    def test_process_file_success(self, mock_move, mock_validate,
                                  sample_vsdx_file, mock_output_dir, mock_archive_dir):
        """Test successful file processing."""
        cli = VDXConvertCLI()
        mock_converter = Mock()
        mock_converter.convert.return_value = True

        output_file = mock_output_dir / "test.vdx"
        archive_file = mock_archive_dir / "test.vsdx"
        mock_move.return_value = archive_file

        with patch.object(cli, 'get_converter_for_file', return_value=mock_converter):
            result = cli.process_file(sample_vsdx_file)

        assert result['success'] is True
        assert result['filename'] == sample_vsdx_file.name
        assert result['error'] is None
        assert 'time' in result

    @patch('src.vdxconvert.cli.validate_input_file')
    def test_process_file_validation_error(self, mock_validate, sample_vsdx_file):
        """Test processing with validation error."""
        cli = VDXConvertCLI()
        mock_validate.side_effect = ValidationError("Invalid file")

        result = cli.process_file(sample_vsdx_file)

        assert result['success'] is False
        assert "Invalid file" in result['error']

    @patch('src.vdxconvert.cli.validate_input_file')
    def test_process_file_no_converter(self, mock_validate, sample_vsdx_file):
        """Test processing when no converter available."""
        cli = VDXConvertCLI()

        with patch.object(cli, 'get_converter_for_file', return_value=None):
            result = cli.process_file(sample_vsdx_file)

        assert result['success'] is False
        assert "No converter available" in result['error']

    @patch('src.vdxconvert.cli.validate_input_file')
    def test_process_file_conversion_error(self, mock_validate, sample_vsdx_file):
        """Test processing with conversion error."""
        cli = VDXConvertCLI()
        mock_converter = Mock()
        mock_converter.convert.side_effect = ConversionError("Conversion failed")

        with patch.object(cli, 'get_converter_for_file', return_value=mock_converter):
            result = cli.process_file(sample_vsdx_file)

        assert result['success'] is False
        assert "Conversion failed" in result['error']

    def test_print_summary_all_success(self, capsys):
        """Test summary printing with all successes."""
        cli = VDXConvertCLI()
        results = [
            {'filename': 'test1.vsdx', 'success': True, 'time': 1.0, 'error': None},
            {'filename': 'test2.vsd', 'success': True, 'time': 2.0, 'error': None}
        ]

        cli.print_summary(results)

        captured = capsys.readouterr()
        assert "Total files processed: 2" in captured.out
        assert "Successful conversions: 2" in captured.out
        assert "Failed conversions: 0" in captured.out

    def test_print_summary_with_failures(self, capsys):
        """Test summary printing with failures."""
        cli = VDXConvertCLI()
        results = [
            {'filename': 'test1.vsdx', 'success': True, 'time': 1.0, 'error': None},
            {'filename': 'test2.vsd', 'success': False, 'time': 0.5, 'error': 'Failed'}
        ]

        cli.print_summary(results)

        captured = capsys.readouterr()
        assert "Total files processed: 2" in captured.out
        assert "Successful conversions: 1" in captured.out
        assert "Failed conversions: 1" in captured.out
        assert "test2.vsd" in captured.out
        assert "Failed" in captured.out

    @patch('src.vdxconvert.cli.setup_logging')
    @patch('src.vdxconvert.cli.ensure_directories_exist')
    @patch('src.vdxconvert.cli.get_visio_files')
    def test_run_no_files(self, mock_get_files, mock_ensure_dirs, mock_logging, capsys):
        """Test run with no input files."""
        cli = VDXConvertCLI()
        mock_get_files.return_value = []
        args = Mock(verbose=False, no_report=False)

        result = cli.run(args)

        assert result == 0
        captured = capsys.readouterr()
        assert "No Visio files found" in captured.out

    @patch('src.vdxconvert.cli.setup_logging')
    @patch('src.vdxconvert.cli.ensure_directories_exist')
    @patch('src.vdxconvert.cli.get_visio_files')
    @patch('builtins.input', return_value='n')
    def test_run_with_files(self, mock_input, mock_get_files, mock_ensure_dirs,
                           mock_logging, sample_vsdx_file):
        """Test run with input files."""
        cli = VDXConvertCLI()
        mock_get_files.return_value = [sample_vsdx_file]
        args = Mock(verbose=False, no_report=False)

        # Mock process_file to return success
        with patch.object(cli, 'process_file') as mock_process:
            mock_process.return_value = {
                'filename': 'test.vsdx',
                'success': True,
                'time': 1.0,
                'error': None,
                'output': 'test.vdx',
                'archive': 'test.vsdx'
            }

            result = cli.run(args)

        assert result == 0
        mock_process.assert_called_once_with(sample_vsdx_file)
