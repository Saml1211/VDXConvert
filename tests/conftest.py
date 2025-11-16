#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Pytest configuration and shared fixtures for VDXConvert tests.
"""

import pytest
import tempfile
import shutil
from pathlib import Path
from unittest.mock import Mock


@pytest.fixture
def temp_dir():
    """Create a temporary directory for testing."""
    tmp = tempfile.mkdtemp()
    yield Path(tmp)
    shutil.rmtree(tmp, ignore_errors=True)


@pytest.fixture
def mock_input_dir(temp_dir):
    """Create mock input directory."""
    input_dir = temp_dir / "input"
    input_dir.mkdir()
    return input_dir


@pytest.fixture
def mock_output_dir(temp_dir):
    """Create mock output directory."""
    output_dir = temp_dir / "output"
    output_dir.mkdir()
    return output_dir


@pytest.fixture
def mock_archive_dir(temp_dir):
    """Create mock archive directory."""
    archive_dir = temp_dir / "archive"
    archive_dir.mkdir()
    return archive_dir


@pytest.fixture
def mock_logs_dir(temp_dir):
    """Create mock logs directory."""
    logs_dir = temp_dir / "logs"
    logs_dir.mkdir()
    return logs_dir


@pytest.fixture
def sample_vsdx_file(mock_input_dir):
    """Create a sample .vsdx file for testing."""
    file_path = mock_input_dir / "test.vsdx"
    file_path.write_text("Sample VSDX content")
    return file_path


@pytest.fixture
def sample_vsd_file(mock_input_dir):
    """Create a sample .vsd file for testing."""
    file_path = mock_input_dir / "test.vsd"
    file_path.write_text("Sample VSD content")
    return file_path


@pytest.fixture
def sample_vsdm_file(mock_input_dir):
    """Create a sample .vsdm file for testing."""
    file_path = mock_input_dir / "test.vsdm"
    file_path.write_text("Sample VSDM content")
    return file_path


@pytest.fixture
def sample_vdw_file(mock_input_dir):
    """Create a sample .vdw file for testing."""
    file_path = mock_input_dir / "test.vdw"
    file_path.write_text("Sample VDW content")
    return file_path


@pytest.fixture
def large_file(mock_input_dir):
    """Create a large file exceeding size limits."""
    file_path = mock_input_dir / "large.vsdx"
    # Create a 600MB file (exceeds 500MB limit)
    with open(file_path, 'wb') as f:
        f.seek(600 * 1024 * 1024 - 1)
        f.write(b'\0')
    return file_path


@pytest.fixture
def mock_converter():
    """Create a mock converter for testing."""
    converter = Mock()
    converter.can_convert.return_value = True
    converter.check_dependencies.return_value = True
    converter.get_missing_dependencies.return_value = []
    converter.convert.return_value = True
    return converter
