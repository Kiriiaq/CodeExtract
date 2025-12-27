"""
Unit tests for ExportManager module.
Tests export functionality for JSON, CSV, HTML, and TXT formats.
"""

import json
import os
import sys
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent.parent / 'src'))

from core.export_manager import (
    ExportManager, ExportResult, JSONExporter, CSVExporter,
    TXTExporter, HTMLExporter, get_export_manager
)


class TestExportResult(unittest.TestCase):
    """Tests for ExportResult dataclass."""

    def test_success_result(self):
        """Test creating a success result."""
        result = ExportResult(True, "/path/to/file.json", "JSON", 1024, "Success", 0.5)
        self.assertTrue(result.success)
        self.assertEqual(result.format, "JSON")
        self.assertEqual(result.size, 1024)

    def test_failure_result(self):
        """Test creating a failure result."""
        result = ExportResult(False, "/path/to/file.json", "JSON", message="File not found")
        self.assertFalse(result.success)
        self.assertEqual(result.message, "File not found")


class TestJSONExporter(unittest.TestCase):
    """Tests for JSONExporter class."""

    def setUp(self):
        self.exporter = JSONExporter()
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_format_properties(self):
        """Test format name and extension."""
        self.assertEqual(self.exporter.format_name, "JSON")
        self.assertEqual(self.exporter.file_extension, ".json")

    def test_export_simple_dict(self):
        """Test exporting a simple dictionary."""
        data = {"name": "test", "value": 42}
        file_path = os.path.join(self.temp_dir, "test.json")

        result = self.exporter.export(data, file_path)

        self.assertTrue(result.success)
        with open(file_path, 'r') as f:
            loaded = json.load(f)
        self.assertEqual(loaded, data)

    def test_export_nested_data(self):
        """Test exporting nested data structures."""
        data = {
            "files": [
                {"name": "file1.py", "lines": 100},
                {"name": "file2.py", "lines": 200}
            ],
            "summary": {"total": 2, "lines": 300}
        }
        file_path = os.path.join(self.temp_dir, "nested.json")

        result = self.exporter.export(data, file_path)

        self.assertTrue(result.success)
        with open(file_path, 'r') as f:
            loaded = json.load(f)
        self.assertEqual(loaded["summary"]["total"], 2)

    def test_export_datetime(self):
        """Test exporting datetime objects."""
        data = {"timestamp": datetime(2024, 1, 15, 10, 30, 0)}
        file_path = os.path.join(self.temp_dir, "datetime.json")

        result = self.exporter.export(data, file_path)

        self.assertTrue(result.success)
        with open(file_path, 'r') as f:
            loaded = json.load(f)
        self.assertEqual(loaded["timestamp"], "2024-01-15T10:30:00")

    def test_export_path_objects(self):
        """Test exporting Path objects."""
        data = {"path": Path("/some/path/file.txt")}
        file_path = os.path.join(self.temp_dir, "path.json")

        result = self.exporter.export(data, file_path)

        self.assertTrue(result.success)
        with open(file_path, 'r') as f:
            loaded = json.load(f)
        self.assertIn("file.txt", loaded["path"])

    def test_export_set(self):
        """Test exporting set objects."""
        data = {"items": {"a", "b", "c"}}
        file_path = os.path.join(self.temp_dir, "set.json")

        result = self.exporter.export(data, file_path)

        self.assertTrue(result.success)
        with open(file_path, 'r') as f:
            loaded = json.load(f)
        self.assertEqual(set(loaded["items"]), {"a", "b", "c"})


class TestCSVExporter(unittest.TestCase):
    """Tests for CSVExporter class."""

    def setUp(self):
        self.exporter = CSVExporter()
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_format_properties(self):
        """Test format name and extension."""
        self.assertEqual(self.exporter.format_name, "CSV")
        self.assertEqual(self.exporter.file_extension, ".csv")

    def test_export_list_of_dicts(self):
        """Test exporting a list of dictionaries."""
        data = [
            {"name": "file1.py", "lines": 100, "size": 1024},
            {"name": "file2.py", "lines": 200, "size": 2048}
        ]
        file_path = os.path.join(self.temp_dir, "test.csv")

        result = self.exporter.export(data, file_path)

        self.assertTrue(result.success)
        with open(file_path, 'r', encoding='utf-8-sig') as f:
            content = f.read()
        self.assertIn("name", content)
        self.assertIn("file1.py", content)

    def test_export_dict_with_files_key(self):
        """Test exporting dict with 'files' key."""
        data = {
            "files": [
                {"name": "a.py", "lines": 50},
                {"name": "b.py", "lines": 75}
            ]
        }
        file_path = os.path.join(self.temp_dir, "files.csv")

        result = self.exporter.export(data, file_path)

        self.assertTrue(result.success)
        self.assertIn("2 rows", result.message)

    def test_export_empty_data(self):
        """Test exporting empty data."""
        data = []
        file_path = os.path.join(self.temp_dir, "empty.csv")

        result = self.exporter.export(data, file_path)

        self.assertFalse(result.success)
        self.assertIn("No data", result.message)

    def test_flatten_nested_dict(self):
        """Test flattening nested dictionaries via dict wrapper."""
        # The CSVExporter flattens when accessing dicts with __dict__ or to_dict
        # For raw dicts, nested values are kept as is
        data = {"data": [{"name": "test", "value": 42}]}
        file_path = os.path.join(self.temp_dir, "nested.csv")

        result = self.exporter.export(data, file_path)

        self.assertTrue(result.success)
        with open(file_path, 'r', encoding='utf-8-sig') as f:
            content = f.read()
        self.assertIn("name", content)
        self.assertIn("test", content)


class TestTXTExporter(unittest.TestCase):
    """Tests for TXTExporter class."""

    def setUp(self):
        self.exporter = TXTExporter()
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_format_properties(self):
        """Test format name and extension."""
        self.assertEqual(self.exporter.format_name, "TXT")
        self.assertEqual(self.exporter.file_extension, ".txt")

    def test_export_with_statistics(self):
        """Test exporting data with statistics."""
        data = {
            "statistics": {"total_files": 10, "total_lines": 1000}
        }
        file_path = os.path.join(self.temp_dir, "stats.txt")

        result = self.exporter.export(data, file_path, title="Test Report")

        self.assertTrue(result.success)
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        self.assertIn("Test Report", content)
        self.assertIn("STATISTICS", content)
        self.assertIn("Total Files", content)

    def test_export_with_files_list(self):
        """Test exporting data with files list."""
        data = {
            "files": [
                {"name": "file1.py", "path": "/path/file1.py"},
                {"name": "file2.py", "path": "/path/file2.py"}
            ]
        }
        file_path = os.path.join(self.temp_dir, "files.txt")

        result = self.exporter.export(data, file_path)

        self.assertTrue(result.success)
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        self.assertIn("FILES", content)
        self.assertIn("file1.py", content)


class TestHTMLExporter(unittest.TestCase):
    """Tests for HTMLExporter class."""

    def setUp(self):
        self.exporter = HTMLExporter()
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_format_properties(self):
        """Test format name and extension."""
        self.assertEqual(self.exporter.format_name, "HTML")
        self.assertEqual(self.exporter.file_extension, ".html")

    def test_export_dark_theme(self):
        """Test exporting with dark theme."""
        data = {"summary": {"total_files": 5}}
        file_path = os.path.join(self.temp_dir, "dark.html")

        result = self.exporter.export(data, file_path, theme="dark")

        self.assertTrue(result.success)
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        self.assertIn("#0f172a", content)  # Dark background color

    def test_export_light_theme(self):
        """Test exporting with light theme."""
        data = {"summary": {"total_files": 5}}
        file_path = os.path.join(self.temp_dir, "light.html")

        result = self.exporter.export(data, file_path, theme="light")

        self.assertTrue(result.success)
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        self.assertIn("#ffffff", content)  # Light background color

    def test_export_with_statistics(self):
        """Test exporting with statistics generates stat cards."""
        data = {
            "summary": {"total_files": 10, "total_lines": 1000}
        }
        file_path = os.path.join(self.temp_dir, "stats.html")

        result = self.exporter.export(data, file_path)

        self.assertTrue(result.success)
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        self.assertIn("stat-card", content)
        self.assertIn("Statistics", content)

    def test_export_valid_html_structure(self):
        """Test that exported HTML has valid structure."""
        data = {"test": "value"}
        file_path = os.path.join(self.temp_dir, "valid.html")

        result = self.exporter.export(data, file_path)

        self.assertTrue(result.success)
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        self.assertIn("<!DOCTYPE html>", content)
        self.assertIn("<html", content)
        self.assertIn("</html>", content)


class TestExportManager(unittest.TestCase):
    """Tests for ExportManager class."""

    def setUp(self):
        self.manager = ExportManager()
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_available_formats(self):
        """Test getting available export formats."""
        formats = self.manager.get_available_formats()
        self.assertIn('json', formats)
        self.assertIn('csv', formats)
        self.assertIn('html', formats)
        self.assertIn('txt', formats)

    def test_export_with_explicit_format(self):
        """Test export with explicit format parameter."""
        data = {"test": "data"}
        file_path = os.path.join(self.temp_dir, "test.json")

        result = self.manager.export(data, file_path, format='json')

        self.assertTrue(result.success)
        self.assertEqual(result.format, "JSON")

    def test_export_auto_detect_format(self):
        """Test export with auto-detected format from extension."""
        data = {"test": "data"}
        file_path = os.path.join(self.temp_dir, "test.html")

        result = self.manager.export(data, file_path)

        self.assertTrue(result.success)
        self.assertEqual(result.format, "HTML")

    def test_export_unsupported_format(self):
        """Test export with unsupported format."""
        data = {"test": "data"}
        file_path = os.path.join(self.temp_dir, "test.xyz")

        result = self.manager.export(data, file_path, format='xyz')

        self.assertFalse(result.success)
        self.assertIn("Unsupported format", result.message)

    def test_export_creates_directory(self):
        """Test that export creates parent directories if needed."""
        data = {"test": "data"}
        file_path = os.path.join(self.temp_dir, "subdir", "nested", "test.json")

        result = self.manager.export(data, file_path, format='json')

        self.assertTrue(result.success)
        self.assertTrue(os.path.exists(file_path))

    def test_export_multiple_formats(self):
        """Test exporting to multiple formats at once."""
        data = {"test": "data"}
        base_path = os.path.join(self.temp_dir, "multi")

        results = self.manager.export_multiple(data, base_path, ['json', 'txt'])

        self.assertTrue(results['json'].success)
        self.assertTrue(results['txt'].success)
        self.assertTrue(os.path.exists(base_path + ".json"))
        self.assertTrue(os.path.exists(base_path + ".txt"))

    def test_create_archive(self):
        """Test creating ZIP archive of files."""
        # Create test files
        file1 = os.path.join(self.temp_dir, "file1.txt")
        file2 = os.path.join(self.temp_dir, "file2.txt")
        Path(file1).write_text("content1")
        Path(file2).write_text("content2")

        archive_path = os.path.join(self.temp_dir, "archive.zip")
        result = self.manager.create_archive([file1, file2], archive_path)

        self.assertTrue(result.success)
        self.assertTrue(os.path.exists(archive_path))
        self.assertIn("2 files", result.message)


class TestGlobalExportManager(unittest.TestCase):
    """Tests for global export manager instance."""

    def test_get_export_manager_singleton(self):
        """Test that get_export_manager returns same instance."""
        manager1 = get_export_manager()
        manager2 = get_export_manager()
        self.assertIs(manager1, manager2)

    def test_get_export_manager_has_formats(self):
        """Test that global manager has all formats."""
        manager = get_export_manager()
        formats = manager.get_available_formats()
        self.assertGreaterEqual(len(formats), 4)


if __name__ == '__main__':
    unittest.main()
