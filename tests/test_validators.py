#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Tests for validators module.
"""

import pytest
from pathlib import Path
from src.vdxconvert.validators import (
    validate_file_extension,
    validate_file_exists,
    validate_file_size,
    validate_filename,
    validate_path_within_directory,
    validate_input_file,
    sanitize_path
)
from src.vdxconvert.exceptions import ValidationError, SecurityError


class TestValidateFileExtension:
    """Tests for validate_file_extension function."""

    def test_valid_extension(self):
        """Test validation with valid extension."""
        result = validate_file_extension(Path("test.vsdx"))
        assert result is True

    def test_valid_extensions(self):
        """Test all supported extensions."""
        for ext in ['.vsd', '.vsdx', '.vsdm', '.vdw']:
            assert validate_file_extension(Path(f"test{ext}")) is True

    def test_case_insensitive(self):
        """Test that extension check is case-insensitive."""
        assert validate_file_extension(Path("test.VSDX")) is True
        assert validate_file_extension(Path("test.VsDx")) is True

    def test_invalid_extension(self):
        """Test validation with invalid extension."""
        with pytest.raises(ValidationError) as exc_info:
            validate_file_extension(Path("test.txt"))
        assert "Invalid file extension" in str(exc_info.value)

    def test_custom_allowed_extensions(self):
        """Test with custom allowed extensions."""
        assert validate_file_extension(Path("test.pdf"), ['.pdf']) is True
        with pytest.raises(ValidationError):
            validate_file_extension(Path("test.vsdx"), ['.pdf'])


class TestValidateFileExists:
    """Tests for validate_file_exists function."""

    def test_existing_file(self, sample_vsdx_file):
        """Test validation with existing file."""
        assert validate_file_exists(sample_vsdx_file) is True

    def test_nonexistent_file(self, temp_dir):
        """Test validation with nonexistent file."""
        nonexistent = temp_dir / "nonexistent.vsdx"
        with pytest.raises(ValidationError) as exc_info:
            validate_file_exists(nonexistent)
        assert "File not found" in str(exc_info.value)

    def test_directory_not_file(self, temp_dir):
        """Test validation when path is a directory."""
        with pytest.raises(ValidationError) as exc_info:
            validate_file_exists(temp_dir)
        assert "not a file" in str(exc_info.value)


class TestValidateFileSize:
    """Tests for validate_file_size function."""

    def test_normal_file_size(self, sample_vsdx_file):
        """Test validation with normal file size."""
        assert validate_file_size(sample_vsdx_file) is True

    def test_large_file(self, large_file):
        """Test validation with oversized file."""
        with pytest.raises(SecurityError) as exc_info:
            validate_file_size(large_file)
        assert "exceeds limit" in str(exc_info.value)

    def test_custom_size_limit(self, sample_vsdx_file):
        """Test with custom size limit."""
        # Should fail with very small limit
        with pytest.raises(SecurityError):
            validate_file_size(sample_vsdx_file, max_size_mb=0.000001)


class TestValidateFilename:
    """Tests for validate_filename function."""

    def test_valid_filename(self):
        """Test with valid filename."""
        assert validate_filename("test.vsdx") is True
        assert validate_filename("my-diagram_v1.vsdx") is True

    def test_filename_too_long(self):
        """Test with excessively long filename."""
        long_name = "a" * 300 + ".vsdx"
        with pytest.raises(ValidationError) as exc_info:
            validate_filename(long_name)
        assert "too long" in str(exc_info.value)

    def test_null_byte_in_filename(self):
        """Test filename with null byte."""
        with pytest.raises(SecurityError) as exc_info:
            validate_filename("test\0.vsdx")
        assert "null byte" in str(exc_info.value)

    def test_path_traversal_attempt(self):
        """Test detection of path traversal attempts."""
        malicious_names = [
            "../../../etc/passwd",
            "..\\..\\windows\\system32",
            "/etc/passwd",
            "\\windows\\system32"
        ]
        for name in malicious_names:
            with pytest.raises(SecurityError) as exc_info:
                validate_filename(name)
            assert "traversal" in str(exc_info.value).lower() or "Suspicious" in str(exc_info.value)

    def test_suspicious_characters(self):
        """Test detection of suspicious characters."""
        suspicious = ['<', '>', ':', '"', '|', '?', '*']
        for char in suspicious:
            with pytest.raises(SecurityError) as exc_info:
                validate_filename(f"test{char}.vsdx")
            assert "Suspicious character" in str(exc_info.value)


class TestValidatePathWithinDirectory:
    """Tests for validate_path_within_directory function."""

    def test_path_within_directory(self, mock_input_dir, sample_vsdx_file):
        """Test validation of path within allowed directory."""
        assert validate_path_within_directory(sample_vsdx_file, mock_input_dir) is True

    def test_path_outside_directory(self, mock_input_dir, temp_dir):
        """Test detection of path outside allowed directory."""
        outside_file = temp_dir / "outside.vsdx"
        outside_file.write_text("content")
        with pytest.raises(SecurityError) as exc_info:
            validate_path_within_directory(outside_file, mock_input_dir)
        assert "traversal" in str(exc_info.value).lower()

    def test_path_traversal_symlink(self, mock_input_dir, temp_dir):
        """Test that symlinks are resolved correctly."""
        # This is a basic test - real symlink tests would need actual symlinks
        safe_file = mock_input_dir / "safe.vsdx"
        safe_file.write_text("content")
        assert validate_path_within_directory(safe_file, mock_input_dir) is True


class TestValidateInputFile:
    """Tests for validate_input_file function."""

    def test_valid_input_file(self, sample_vsdx_file, monkeypatch):
        """Test comprehensive validation of valid file."""
        from src.vdxconvert import config
        monkeypatch.setattr(config, 'INPUT_DIR', sample_vsdx_file.parent)
        assert validate_input_file(sample_vsdx_file) is True

    def test_invalid_extension(self, mock_input_dir, monkeypatch):
        """Test validation fails for invalid extension."""
        from src.vdxconvert import config
        monkeypatch.setattr(config, 'INPUT_DIR', mock_input_dir)

        invalid_file = mock_input_dir / "test.txt"
        invalid_file.write_text("content")

        with pytest.raises(ValidationError):
            validate_input_file(invalid_file)


class TestSanitizePath:
    """Tests for sanitize_path function."""

    def test_sanitize_simple_filename(self, mock_input_dir):
        """Test sanitization of simple filename."""
        result = sanitize_path("test.vsdx", mock_input_dir)
        assert result == mock_input_dir / "test.vsdx"

    def test_sanitize_path_traversal(self, mock_input_dir):
        """Test that path traversal is neutralized."""
        result = sanitize_path("../../etc/passwd", mock_input_dir)
        # Should only keep the filename part
        assert result == mock_input_dir / "passwd"

    def test_sanitize_removes_null_bytes(self, mock_input_dir):
        """Test that null bytes are removed."""
        with pytest.raises(SecurityError):
            sanitize_path("test\0.vsdx", mock_input_dir)

    def test_sanitize_directory_components(self, mock_input_dir):
        """Test that directory components are removed."""
        result = sanitize_path("subdir/test.vsdx", mock_input_dir)
        assert result == mock_input_dir / "test.vsdx"
