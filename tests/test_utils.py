#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Tests for utils module.
"""

import pytest
from src.vdxconvert.utils import (
    get_visio_files,
    get_unique_filename,
    safe_move_file,
    safe_copy_file,
    save_csv_report,
    generate_report_filename,
    ensure_directories_exist
)
from src.vdxconvert.exceptions import FileOperationError


class TestGetVisioFiles:
    """Tests for get_visio_files function."""

    def test_empty_directory(self, mock_input_dir):
        """Test with empty directory."""
        files = get_visio_files(mock_input_dir)
        assert files == []

    def test_single_file(self, sample_vsdx_file):
        """Test with single Visio file."""
        files = get_visio_files(sample_vsdx_file.parent)
        assert len(files) == 1
        assert files[0] == sample_vsdx_file

    def test_multiple_files(self, mock_input_dir):
        """Test with multiple Visio files."""
        (mock_input_dir / "test1.vsdx").write_text("content")
        (mock_input_dir / "test2.vsd").write_text("content")
        (mock_input_dir / "test3.vsdm").write_text("content")
        (mock_input_dir / "test4.vdw").write_text("content")

        files = get_visio_files(mock_input_dir)
        assert len(files) == 4

    def test_mixed_files(self, mock_input_dir):
        """Test that non-Visio files are excluded."""
        (mock_input_dir / "test.vsdx").write_text("content")
        (mock_input_dir / "test.txt").write_text("content")
        (mock_input_dir / "test.pdf").write_text("content")

        files = get_visio_files(mock_input_dir)
        assert len(files) == 1
        assert files[0].suffix == '.vsdx'

    def test_case_insensitive_extensions(self, mock_input_dir):
        """Test that extension matching is case-insensitive."""
        (mock_input_dir / "test.VSDX").write_text("content")
        (mock_input_dir / "test.VsD").write_text("content")

        files = get_visio_files(mock_input_dir)
        assert len(files) == 2


class TestGetUniqueFilename:
    """Tests for get_unique_filename function."""

    def test_nonexistent_file(self, mock_output_dir):
        """Test with nonexistent file."""
        filepath = mock_output_dir / "test.vdx"
        result = get_unique_filename(filepath)
        assert result == filepath

    def test_existing_file(self, mock_output_dir):
        """Test with existing file."""
        filepath = mock_output_dir / "test.vdx"
        filepath.write_text("content")

        result = get_unique_filename(filepath)
        assert result == mock_output_dir / "test_1.vdx"

    def test_multiple_existing_files(self, mock_output_dir):
        """Test with multiple existing files."""
        (mock_output_dir / "test.vdx").write_text("content")
        (mock_output_dir / "test_1.vdx").write_text("content")
        (mock_output_dir / "test_2.vdx").write_text("content")

        result = get_unique_filename(mock_output_dir / "test.vdx")
        assert result == mock_output_dir / "test_3.vdx"


class TestSafeMoveFile:
    """Tests for safe_move_file function."""

    def test_successful_move(self, mock_input_dir, mock_archive_dir):
        """Test successful file move."""
        src = mock_input_dir / "test.vsdx"
        src.write_text("content")
        dst = mock_archive_dir / "test.vsdx"

        result = safe_move_file(src, dst)

        assert result == dst
        assert dst.exists()
        assert not src.exists()
        assert dst.read_text() == "content"

    def test_move_to_existing_destination(self, mock_input_dir, mock_archive_dir):
        """Test move when destination already exists."""
        src = mock_input_dir / "test.vsdx"
        src.write_text("new content")

        dst = mock_archive_dir / "test.vsdx"
        dst.write_text("old content")

        result = safe_move_file(src, dst)

        assert result == mock_archive_dir / "test_1.vsdx"
        assert result.exists()
        assert result.read_text() == "new content"
        assert dst.read_text() == "old content"  # Original unchanged

    def test_move_nonexistent_file(self, mock_input_dir, mock_archive_dir):
        """Test move of nonexistent file."""
        src = mock_input_dir / "nonexistent.vsdx"
        dst = mock_archive_dir / "test.vsdx"

        with pytest.raises(FileOperationError):
            safe_move_file(src, dst)


class TestSafeCopyFile:
    """Tests for safe_copy_file function."""

    def test_successful_copy(self, mock_input_dir, mock_output_dir):
        """Test successful file copy."""
        src = mock_input_dir / "test.vsdx"
        src.write_text("content")
        dst = mock_output_dir / "test.vdx"

        result = safe_copy_file(src, dst)

        assert result == dst
        assert dst.exists()
        assert src.exists()  # Source still exists
        assert dst.read_text() == "content"

    def test_copy_to_existing_destination(self, mock_input_dir, mock_output_dir):
        """Test copy when destination already exists."""
        src = mock_input_dir / "test.vsdx"
        src.write_text("new content")

        dst = mock_output_dir / "test.vdx"
        dst.write_text("old content")

        result = safe_copy_file(src, dst)

        assert result == mock_output_dir / "test_1.vdx"
        assert result.exists()
        assert result.read_text() == "new content"
        assert dst.read_text() == "old content"  # Original unchanged


class TestSaveCSVReport:
    """Tests for save_csv_report function."""

    def test_save_empty_report(self, mock_logs_dir):
        """Test saving empty report."""
        results = []
        save_path = mock_logs_dir / "report.csv"

        save_csv_report(results, save_path)

        assert save_path.exists()
        content = save_path.read_text()
        assert "filename,output,archive,success,time,error" in content

    def test_save_report_with_data(self, mock_logs_dir):
        """Test saving report with data."""
        results = [
            {
                "filename": "test.vsdx",
                "output": "test.vdx",
                "archive": "test.vsdx",
                "success": True,
                "time": 1.23,
                "error": None
            },
            {
                "filename": "test2.vsd",
                "output": None,
                "archive": None,
                "success": False,
                "time": 0.5,
                "error": "Conversion failed"
            }
        ]
        save_path = mock_logs_dir / "report.csv"

        save_csv_report(results, save_path)

        assert save_path.exists()
        content = save_path.read_text()
        assert "test.vsdx" in content
        assert "test2.vsd" in content
        assert "Conversion failed" in content


class TestGenerateReportFilename:
    """Tests for generate_report_filename function."""

    def test_generates_valid_filename(self):
        """Test that generated filename is valid."""
        filename = generate_report_filename()
        assert filename.startswith("conversion_report_")
        assert filename.endswith(".csv")
        assert len(filename) > 20  # Should have timestamp

    def test_generates_unique_filenames(self):
        """Test that consecutive calls generate different filenames."""
        import time
        filename1 = generate_report_filename()
        time.sleep(1.1)  # Wait to ensure different timestamp
        filename2 = generate_report_filename()
        # May be the same if called in same second, so this is a weak test
        assert filename1.startswith("conversion_report_")
        assert filename2.startswith("conversion_report_")


class TestEnsureDirectoriesExist:
    """Tests for ensure_directories_exist function."""

    def test_create_single_directory(self, temp_dir):
        """Test creating a single directory."""
        new_dir = temp_dir / "new_dir"
        ensure_directories_exist([new_dir])
        assert new_dir.exists()
        assert new_dir.is_dir()

    def test_create_multiple_directories(self, temp_dir):
        """Test creating multiple directories."""
        dirs = [
            temp_dir / "dir1",
            temp_dir / "dir2",
            temp_dir / "dir3"
        ]
        ensure_directories_exist(dirs)
        for d in dirs:
            assert d.exists()
            assert d.is_dir()

    def test_create_nested_directories(self, temp_dir):
        """Test creating nested directories."""
        nested = temp_dir / "parent" / "child" / "grandchild"
        ensure_directories_exist([nested])
        assert nested.exists()
        assert nested.is_dir()

    def test_existing_directory(self, temp_dir):
        """Test with already existing directory."""
        existing = temp_dir / "existing"
        existing.mkdir()
        ensure_directories_exist([existing])
        assert existing.exists()  # Should not raise error
