"""
Unit tests for Helper Functions module.
Tests utility functions used across the application.
"""

import os
import sys
import tempfile
import unittest
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent.parent / 'src'))

from utils.helpers import (
    safe_path, detect_encoding, format_size, sanitize_filename,
    generate_timestamp, calculate_file_hash, read_file_safe,
    find_files, create_directory_tree, merge_dicts, truncate_string
)


class TestSafePath(unittest.TestCase):
    """Tests for safe_path function."""

    def test_convert_string_to_path(self):
        """Test converting string to Path."""
        result = safe_path("/some/path")
        self.assertIsInstance(result, Path)

    def test_resolve_relative_path(self):
        """Test that relative paths are resolved."""
        result = safe_path("./relative/path")
        self.assertTrue(result.is_absolute())


class TestDetectEncoding(unittest.TestCase):
    """Tests for detect_encoding function."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_detect_utf8(self):
        """Test detecting UTF-8 encoding."""
        file_path = os.path.join(self.temp_dir, "utf8.txt")
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write("Hello World")

        encoding = detect_encoding(file_path)
        self.assertIn(encoding.lower(), ['utf-8', 'ascii'])

    def test_detect_utf8_bom(self):
        """Test detecting UTF-8 with BOM."""
        file_path = os.path.join(self.temp_dir, "utf8bom.txt")
        with open(file_path, 'wb') as f:
            f.write(b'\xef\xbb\xbfHello World')

        encoding = detect_encoding(file_path)
        self.assertEqual(encoding, 'utf-8-sig')

    def test_detect_utf16_le(self):
        """Test detecting UTF-16 LE."""
        file_path = os.path.join(self.temp_dir, "utf16le.txt")
        with open(file_path, 'wb') as f:
            f.write(b'\xff\xfeH\x00e\x00l\x00l\x00o\x00')

        encoding = detect_encoding(file_path)
        self.assertEqual(encoding, 'utf-16-le')

    def test_detect_utf16_be(self):
        """Test detecting UTF-16 BE."""
        file_path = os.path.join(self.temp_dir, "utf16be.txt")
        with open(file_path, 'wb') as f:
            f.write(b'\xfe\xff\x00H\x00e\x00l\x00l\x00o')

        encoding = detect_encoding(file_path)
        self.assertEqual(encoding, 'utf-16-be')

    def test_fallback_to_latin1(self):
        """Test fallback to latin-1 for binary-like content."""
        file_path = os.path.join(self.temp_dir, "binary.txt")
        with open(file_path, 'wb') as f:
            f.write(bytes(range(256)))

        encoding = detect_encoding(file_path)
        self.assertEqual(encoding, 'latin-1')


class TestFormatSize(unittest.TestCase):
    """Tests for format_size function."""

    def test_format_bytes(self):
        """Test formatting bytes."""
        self.assertEqual(format_size(500), "500 B")

    def test_format_kilobytes(self):
        """Test formatting kilobytes."""
        result = format_size(1024)
        self.assertIn("KB", result)

    def test_format_megabytes(self):
        """Test formatting megabytes."""
        result = format_size(1024 * 1024)
        self.assertIn("MB", result)

    def test_format_gigabytes(self):
        """Test formatting gigabytes."""
        result = format_size(1024 * 1024 * 1024)
        self.assertIn("GB", result)

    def test_format_with_decimals(self):
        """Test formatting with decimal values."""
        result = format_size(1536)  # 1.5 KB
        self.assertIn("1.5", result)


class TestSanitizeFilename(unittest.TestCase):
    """Tests for sanitize_filename function."""

    def test_remove_invalid_characters(self):
        """Test removing invalid filename characters."""
        result = sanitize_filename('file<>:"/\\|?*.txt')
        self.assertNotIn('<', result)
        self.assertNotIn('>', result)
        self.assertNotIn(':', result)
        self.assertNotIn('"', result)
        self.assertNotIn('|', result)
        self.assertNotIn('?', result)
        self.assertNotIn('*', result)

    def test_preserve_valid_characters(self):
        """Test preserving valid characters."""
        result = sanitize_filename('valid_file-name.txt')
        self.assertEqual(result, 'valid_file-name.txt')

    def test_trim_long_filename(self):
        """Test trimming very long filenames."""
        long_name = 'a' * 300 + '.txt'
        result = sanitize_filename(long_name)
        self.assertLessEqual(len(result), 255)
        self.assertTrue(result.endswith('.txt'))

    def test_empty_filename_fallback(self):
        """Test fallback for empty/invalid filename."""
        result = sanitize_filename('...')
        self.assertEqual(result, 'unnamed')

    def test_remove_control_characters(self):
        """Test removing control characters."""
        result = sanitize_filename('file\x00\x1fname.txt')
        self.assertNotIn('\x00', result)
        self.assertNotIn('\x1f', result)


class TestGenerateTimestamp(unittest.TestCase):
    """Tests for generate_timestamp function."""

    def test_format(self):
        """Test timestamp format."""
        result = generate_timestamp()
        # Format should be YYYYMMDD_HHMMSS
        self.assertEqual(len(result), 15)
        self.assertIn('_', result)

    def test_contains_digits(self):
        """Test that timestamp contains only digits and underscore."""
        result = generate_timestamp()
        cleaned = result.replace('_', '')
        self.assertTrue(cleaned.isdigit())


class TestCalculateFileHash(unittest.TestCase):
    """Tests for calculate_file_hash function."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_md5_hash(self):
        """Test calculating MD5 hash."""
        file_path = os.path.join(self.temp_dir, "test.txt")
        Path(file_path).write_text("Hello World")

        hash_value = calculate_file_hash(file_path, 'md5')
        self.assertEqual(len(hash_value), 32)  # MD5 is 32 hex chars

    def test_sha256_hash(self):
        """Test calculating SHA256 hash."""
        file_path = os.path.join(self.temp_dir, "test.txt")
        Path(file_path).write_text("Hello World")

        hash_value = calculate_file_hash(file_path, 'sha256')
        self.assertEqual(len(hash_value), 64)  # SHA256 is 64 hex chars

    def test_same_content_same_hash(self):
        """Test that same content produces same hash."""
        file1 = os.path.join(self.temp_dir, "file1.txt")
        file2 = os.path.join(self.temp_dir, "file2.txt")
        Path(file1).write_text("Same content")
        Path(file2).write_text("Same content")

        hash1 = calculate_file_hash(file1)
        hash2 = calculate_file_hash(file2)
        self.assertEqual(hash1, hash2)

    def test_different_content_different_hash(self):
        """Test that different content produces different hash."""
        file1 = os.path.join(self.temp_dir, "file1.txt")
        file2 = os.path.join(self.temp_dir, "file2.txt")
        Path(file1).write_text("Content A")
        Path(file2).write_text("Content B")

        hash1 = calculate_file_hash(file1)
        hash2 = calculate_file_hash(file2)
        self.assertNotEqual(hash1, hash2)


class TestReadFileSafe(unittest.TestCase):
    """Tests for read_file_safe function."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_read_small_file(self):
        """Test reading a small file."""
        file_path = os.path.join(self.temp_dir, "small.txt")
        Path(file_path).write_text("Small content")

        content, encoding = read_file_safe(file_path)
        self.assertEqual(content, "Small content")
        self.assertIsNotNone(encoding)

    def test_reject_large_file(self):
        """Test rejecting a file that's too large."""
        file_path = os.path.join(self.temp_dir, "large.txt")
        # Create file larger than default max (10MB)
        Path(file_path).write_bytes(b'x' * (11 * 1024 * 1024))

        content, error = read_file_safe(file_path)
        self.assertIsNone(content)
        self.assertIn("too large", error)

    def test_read_with_custom_max_size(self):
        """Test reading with custom max size."""
        file_path = os.path.join(self.temp_dir, "medium.txt")
        Path(file_path).write_bytes(b'x' * 1000)

        content, error = read_file_safe(file_path, max_size=500)
        self.assertIsNone(content)
        self.assertIn("too large", error)

    def test_handle_missing_file(self):
        """Test handling missing file."""
        content, error = read_file_safe("/nonexistent/file.txt")
        self.assertIsNone(content)
        self.assertIsNotNone(error)


class TestFindFiles(unittest.TestCase):
    """Tests for find_files function."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        # Create test structure
        os.makedirs(os.path.join(self.temp_dir, "subdir"))
        os.makedirs(os.path.join(self.temp_dir, "__pycache__"))
        Path(os.path.join(self.temp_dir, "file1.py")).touch()
        Path(os.path.join(self.temp_dir, "file2.txt")).touch()
        Path(os.path.join(self.temp_dir, "subdir", "file3.py")).touch()
        Path(os.path.join(self.temp_dir, "__pycache__", "cached.pyc")).touch()

    def tearDown(self):
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_find_by_extension(self):
        """Test finding files by extension."""
        files = find_files(self.temp_dir, ['*.py'])
        self.assertEqual(len(files), 2)  # file1.py and file3.py

    def test_exclude_directories(self):
        """Test excluding directories."""
        files = find_files(self.temp_dir, ['*.pyc'])
        # __pycache__ should be excluded by default
        self.assertEqual(len(files), 0)

    def test_multiple_patterns(self):
        """Test finding with multiple patterns."""
        files = find_files(self.temp_dir, ['*.py', '*.txt'])
        self.assertEqual(len(files), 3)

    def test_exclude_patterns(self):
        """Test excluding file patterns."""
        files = find_files(self.temp_dir, ['*.py'], exclude_patterns=['*3*'])
        self.assertEqual(len(files), 1)  # Only file1.py

    def test_max_depth(self):
        """Test max recursion depth."""
        files = find_files(self.temp_dir, ['*.py'], max_depth=0)
        self.assertEqual(len(files), 1)  # Only file1.py at root


class TestCreateDirectoryTree(unittest.TestCase):
    """Tests for create_directory_tree function."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        # Create test structure
        os.makedirs(os.path.join(self.temp_dir, "subdir"))
        Path(os.path.join(self.temp_dir, "file.txt")).touch()
        Path(os.path.join(self.temp_dir, "subdir", "nested.txt")).touch()

    def tearDown(self):
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_includes_files(self):
        """Test that tree includes files."""
        tree = create_directory_tree(self.temp_dir)
        self.assertIn("file.txt", tree)
        self.assertIn("nested.txt", tree)

    def test_includes_directories(self):
        """Test that tree includes directories."""
        tree = create_directory_tree(self.temp_dir)
        self.assertIn("subdir", tree)

    def test_uses_tree_characters(self):
        """Test that tree uses proper characters."""
        tree = create_directory_tree(self.temp_dir)
        self.assertTrue("├── " in tree or "└── " in tree)


class TestMergeDicts(unittest.TestCase):
    """Tests for merge_dicts function."""

    def test_simple_merge(self):
        """Test merging simple dictionaries."""
        d1 = {"a": 1, "b": 2}
        d2 = {"c": 3, "d": 4}
        result = merge_dicts(d1, d2)
        self.assertEqual(result, {"a": 1, "b": 2, "c": 3, "d": 4})

    def test_override_values(self):
        """Test that later dicts override earlier values."""
        d1 = {"a": 1, "b": 2}
        d2 = {"b": 3, "c": 4}
        result = merge_dicts(d1, d2)
        self.assertEqual(result["b"], 3)

    def test_deep_merge(self):
        """Test deep merging nested dictionaries."""
        d1 = {"nested": {"a": 1, "b": 2}}
        d2 = {"nested": {"b": 3, "c": 4}}
        result = merge_dicts(d1, d2)
        self.assertEqual(result["nested"], {"a": 1, "b": 3, "c": 4})

    def test_multiple_dicts(self):
        """Test merging multiple dictionaries."""
        d1 = {"a": 1}
        d2 = {"b": 2}
        d3 = {"c": 3}
        result = merge_dicts(d1, d2, d3)
        self.assertEqual(result, {"a": 1, "b": 2, "c": 3})


class TestTruncateString(unittest.TestCase):
    """Tests for truncate_string function."""

    def test_no_truncation_needed(self):
        """Test that short strings aren't truncated."""
        result = truncate_string("short", max_length=100)
        self.assertEqual(result, "short")

    def test_truncation_with_suffix(self):
        """Test truncation with suffix."""
        result = truncate_string("a" * 50, max_length=20)
        self.assertEqual(len(result), 20)
        self.assertTrue(result.endswith("..."))

    def test_custom_suffix(self):
        """Test truncation with custom suffix."""
        result = truncate_string("a" * 50, max_length=20, suffix="[...]")
        self.assertTrue(result.endswith("[...]"))

    def test_exact_length(self):
        """Test string at exact max length."""
        result = truncate_string("a" * 10, max_length=10)
        self.assertEqual(len(result), 10)
        self.assertFalse(result.endswith("..."))


if __name__ == '__main__':
    unittest.main()
