"""
Unit tests for ConfigManager module.
Tests configuration loading, saving, access, and persistence.
"""

import json
import os
import sys
import tempfile
import unittest
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent.parent / 'src'))

from core.config_manager import (
    ConfigManager, AppConfig, VBAExtractorConfig, PythonAnalyzerConfig,
    FolderScannerConfig, VBAOptimizerConfig, UIConfig, ExportConfig
)


class TestConfigManager(unittest.TestCase):
    """Tests for ConfigManager class."""

    def setUp(self):
        """Create a temporary config directory for each test."""
        self.temp_dir = tempfile.mkdtemp()
        self.config_manager = ConfigManager(config_dir=Path(self.temp_dir))

    def tearDown(self):
        """Clean up temporary directory."""
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_init_creates_default_config(self):
        """Test that initialization creates default config file."""
        config_file = Path(self.temp_dir) / "config.json"
        self.assertTrue(config_file.exists())

    def test_default_values(self):
        """Test default configuration values."""
        cfg = self.config_manager.config

        self.assertEqual(cfg.version, "1.0.0")
        self.assertEqual(cfg.ui.theme, "dark")
        self.assertEqual(cfg.export.default_format, "html")
        self.assertTrue(cfg.vba_extractor.create_individual_files)
        self.assertTrue(cfg.python_analyzer.include_subdirs)

    def test_save_and_load(self):
        """Test saving and loading configuration."""
        # Modify config
        self.config_manager.config.ui.theme = "light"
        self.config_manager.config.export.open_after_export = False
        self.config_manager.save()

        # Create new manager and load
        new_manager = ConfigManager(config_dir=Path(self.temp_dir))

        self.assertEqual(new_manager.config.ui.theme, "light")
        self.assertFalse(new_manager.config.export.open_after_export)

    def test_get_simple_key(self):
        """Test getting values with dot-notation for simple keys."""
        value = self.config_manager.get("ui.theme")
        self.assertEqual(value, "dark")

    def test_get_nested_key(self):
        """Test getting nested configuration values."""
        value = self.config_manager.get("vba_extractor.extraction_method")
        self.assertEqual(value, "auto")

    def test_get_with_default(self):
        """Test getting non-existent key returns default."""
        value = self.config_manager.get("nonexistent.key", "default_value")
        self.assertEqual(value, "default_value")

    def test_set_simple_key(self):
        """Test setting values with dot-notation."""
        self.config_manager.set("ui.theme", "light", auto_save=False)
        self.assertEqual(self.config_manager.config.ui.theme, "light")

    def test_set_nested_key(self):
        """Test setting nested configuration values."""
        self.config_manager.set("vba_extractor.extraction_method", "oletools", auto_save=False)
        self.assertEqual(self.config_manager.config.vba_extractor.extraction_method, "oletools")

    def test_reset_to_defaults_all(self):
        """Test resetting all configuration to defaults."""
        # Modify some values
        self.config_manager.config.ui.theme = "light"
        self.config_manager.config.vba_extractor.extraction_method = "win32com"

        # Reset all
        self.config_manager.reset_to_defaults()

        self.assertEqual(self.config_manager.config.ui.theme, "dark")
        self.assertEqual(self.config_manager.config.vba_extractor.extraction_method, "auto")

    def test_reset_to_defaults_section(self):
        """Test resetting specific section to defaults."""
        # Modify values
        self.config_manager.config.ui.theme = "light"
        self.config_manager.config.vba_extractor.extraction_method = "win32com"

        # Reset only ui section
        self.config_manager.reset_to_defaults("ui")

        # ui should be reset, vba_extractor should not
        self.assertEqual(self.config_manager.config.ui.theme, "dark")
        self.assertEqual(self.config_manager.config.vba_extractor.extraction_method, "win32com")

    def test_add_recent_file(self):
        """Test adding files to recent files list."""
        self.config_manager.add_recent_file("/path/to/file1.py")
        self.config_manager.add_recent_file("/path/to/file2.py")

        recent = self.config_manager.config.ui.recent_files
        self.assertEqual(len(recent), 2)
        self.assertEqual(recent[0], "/path/to/file2.py")  # Most recent first

    def test_recent_files_no_duplicates(self):
        """Test that recent files list doesn't have duplicates."""
        self.config_manager.add_recent_file("/path/to/file1.py")
        self.config_manager.add_recent_file("/path/to/file2.py")
        self.config_manager.add_recent_file("/path/to/file1.py")  # Duplicate

        recent = self.config_manager.config.ui.recent_files
        self.assertEqual(len(recent), 2)
        self.assertEqual(recent[0], "/path/to/file1.py")  # Moved to front

    def test_recent_files_max_limit(self):
        """Test that recent files list respects max limit."""
        for i in range(15):
            self.config_manager.add_recent_file(f"/path/to/file{i}.py")

        recent = self.config_manager.config.ui.recent_files
        self.assertEqual(len(recent), 10)  # max_recent_files default is 10

    def test_observer_pattern(self):
        """Test observer notification on config save."""
        callback_called = [False]

        def observer(config):
            callback_called[0] = True

        self.config_manager.add_observer(observer)
        self.config_manager.save()

        self.assertTrue(callback_called[0])

    def test_remove_observer(self):
        """Test removing an observer."""
        callback_count = [0]

        def observer(config):
            callback_count[0] += 1

        self.config_manager.add_observer(observer)
        self.config_manager.save()
        self.assertEqual(callback_count[0], 1)

        self.config_manager.remove_observer(observer)
        self.config_manager.save()
        self.assertEqual(callback_count[0], 1)  # Should not increase

    def test_export_config(self):
        """Test exporting configuration to file."""
        export_path = Path(self.temp_dir) / "exported_config.json"

        result = self.config_manager.export_config(str(export_path))

        self.assertTrue(result)
        self.assertTrue(export_path.exists())

        with open(export_path, 'r') as f:
            data = json.load(f)
        self.assertIn('ui', data)
        self.assertIn('export', data)

    def test_import_config(self):
        """Test importing configuration from file."""
        # Create config file to import
        import_data = {
            "ui": {"theme": "light", "color_scheme": "green", "window_width": 1400,
                   "window_height": 900, "show_tooltips": True, "auto_save_config": True,
                   "confirm_on_exit": False, "recent_files": [], "max_recent_files": 10,
                   "language": "en"},
            "export": {"default_format": "json", "open_after_export": False,
                      "include_timestamp": True, "compress_output": False, "pdf_page_size": "A4"}
        }
        import_path = Path(self.temp_dir) / "import_config.json"
        with open(import_path, 'w') as f:
            json.dump(import_data, f)

        result = self.config_manager.import_config(str(import_path))

        self.assertTrue(result)
        self.assertEqual(self.config_manager.config.ui.theme, "light")
        self.assertEqual(self.config_manager.config.export.default_format, "json")

    def test_backup_rotation(self):
        """Test that backups are rotated on save."""
        # Save multiple times to trigger backup rotation
        for i in range(5):
            self.config_manager.config.ui.window_width = 1000 + i
            self.config_manager.save()

        backup1 = Path(self.temp_dir) / "config.backup.1.json"
        backup2 = Path(self.temp_dir) / "config.backup.2.json"

        self.assertTrue(backup1.exists())
        self.assertTrue(backup2.exists())

    def test_corrupted_config_recovery(self):
        """Test recovery from corrupted config file."""
        config_file = Path(self.temp_dir) / "config.json"

        # Write corrupted data
        with open(config_file, 'w') as f:
            f.write("{ invalid json }")

        # Create new manager - should recover
        new_manager = ConfigManager(config_dir=Path(self.temp_dir))

        # Should have default values after recovery
        self.assertEqual(new_manager.config.version, "1.0.0")


class TestConfigDataclasses(unittest.TestCase):
    """Tests for configuration dataclasses."""

    def test_vba_extractor_config_defaults(self):
        """Test VBAExtractorConfig default values."""
        cfg = VBAExtractorConfig()
        self.assertEqual(cfg.extraction_method, "auto")
        self.assertTrue(cfg.create_individual_files)
        self.assertEqual(cfg.output_encoding, "utf-8")

    def test_python_analyzer_config_defaults(self):
        """Test PythonAnalyzerConfig default values."""
        cfg = PythonAnalyzerConfig()
        self.assertTrue(cfg.include_subdirs)
        self.assertEqual(cfg.max_workers, 4)
        self.assertIn("test_*", cfg.exclude_patterns)

    def test_folder_scanner_config_defaults(self):
        """Test FolderScannerConfig default values."""
        cfg = FolderScannerConfig()
        self.assertTrue(cfg.include_content)
        self.assertFalse(cfg.include_binary)
        self.assertIn("__pycache__", cfg.excluded_dirs)

    def test_ui_config_defaults(self):
        """Test UIConfig default values."""
        cfg = UIConfig()
        self.assertEqual(cfg.theme, "dark")
        self.assertEqual(cfg.color_scheme, "blue")
        self.assertTrue(cfg.show_tooltips)
        self.assertEqual(cfg.max_recent_files, 10)


if __name__ == '__main__':
    unittest.main()
