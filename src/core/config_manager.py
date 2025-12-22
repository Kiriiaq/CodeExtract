"""
Configuration Manager - Centralized configuration system with JSON persistence.
Handles user preferences, tool settings, and automatic save/load.
"""

import json
import os
from dataclasses import dataclass, field, asdict
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional
import shutil


@dataclass
class VBAExtractorConfig:
    """Configuration for VBA Extractor tool."""
    extraction_method: str = "auto"  # auto, win32com, oletools
    create_individual_files: bool = True
    create_concatenated_file: bool = True
    include_metadata: bool = True
    output_encoding: str = "utf-8"
    last_source_path: str = ""
    last_output_path: str = ""


@dataclass
class PythonAnalyzerConfig:
    """Configuration for Python Analyzer tool."""
    include_subdirs: bool = True
    analyze_quality: bool = True
    max_workers: int = 4
    exclude_patterns: List[str] = field(default_factory=lambda: ["test_*", "*_test.py"])
    report_format: str = "html"  # html, json, markdown
    last_source_path: str = ""
    last_output_path: str = ""


@dataclass
class FolderScannerConfig:
    """Configuration for Folder Scanner tool."""
    include_content: bool = True
    include_binary: bool = False
    max_file_size_kb: int = 1024
    excluded_dirs: List[str] = field(default_factory=lambda: [
        "__pycache__", ".git", "node_modules", ".venv", "venv", ".idea", ".vscode"
    ])
    excluded_extensions: List[str] = field(default_factory=lambda: [
        ".exe", ".dll", ".so", ".pyc", ".pyo", ".jpg", ".png", ".gif", ".mp3", ".mp4"
    ])
    output_format: str = "txt"  # txt, json, html
    last_source_path: str = ""
    last_output_path: str = ""


@dataclass
class VBAOptimizerConfig:
    """Configuration for VBA Optimizer tool."""
    remove_comments: bool = False
    auto_indent: bool = True
    remove_empty_lines: bool = True
    rename_unused_vars: bool = False
    minify: bool = False
    indent_size: int = 4
    create_backup: bool = True
    last_source_path: str = ""
    last_output_path: str = ""


@dataclass
class UIConfig:
    """Configuration for UI appearance and behavior."""
    theme: str = "dark"  # dark, light, system
    color_scheme: str = "blue"  # blue, green, dark-blue
    window_width: int = 1400
    window_height: int = 900
    window_x: Optional[int] = None
    window_y: Optional[int] = None
    show_tooltips: bool = True
    auto_save_config: bool = True
    confirm_on_exit: bool = False
    recent_files: List[str] = field(default_factory=list)
    max_recent_files: int = 10
    language: str = "fr"  # fr, en


@dataclass
class ExportConfig:
    """Configuration for export settings."""
    default_format: str = "html"  # html, json, csv, txt, pdf
    open_after_export: bool = True
    include_timestamp: bool = True
    compress_output: bool = False
    pdf_page_size: str = "A4"


@dataclass
class AppConfig:
    """Main application configuration."""
    version: str = "2.0.0"
    first_run: bool = True
    last_used: str = ""

    # Tool configs
    vba_extractor: VBAExtractorConfig = field(default_factory=VBAExtractorConfig)
    python_analyzer: PythonAnalyzerConfig = field(default_factory=PythonAnalyzerConfig)
    folder_scanner: FolderScannerConfig = field(default_factory=FolderScannerConfig)
    vba_optimizer: VBAOptimizerConfig = field(default_factory=VBAOptimizerConfig)

    # UI and export
    ui: UIConfig = field(default_factory=UIConfig)
    export: ExportConfig = field(default_factory=ExportConfig)


class ConfigManager:
    """
    Manages application configuration with JSON persistence.
    Features: auto-save, reset to defaults, backup, migration.
    """

    DEFAULT_CONFIG_DIR = Path.home() / ".codeextractpro"
    DEFAULT_CONFIG_FILE = "config.json"
    BACKUP_COUNT = 3

    def __init__(self, config_dir: Optional[Path] = None):
        self.config_dir = config_dir or self.DEFAULT_CONFIG_DIR
        self.config_file = self.config_dir / self.DEFAULT_CONFIG_FILE
        self.config: AppConfig = AppConfig()
        self._observers: List[callable] = []

        # Ensure config directory exists
        self.config_dir.mkdir(parents=True, exist_ok=True)

        # Load existing config or create default
        self.load()

    def load(self) -> bool:
        """Load configuration from file."""
        if not self.config_file.exists():
            self.config = AppConfig()
            self.save()
            return False

        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                data = json.load(f)

            self.config = self._dict_to_config(data)
            self.config.last_used = datetime.now().isoformat()
            self.config.first_run = False
            return True

        except Exception as e:
            print(f"Error loading config: {e}")
            self._backup_corrupted()
            self.config = AppConfig()
            return False

    def save(self) -> bool:
        """Save configuration to file."""
        try:
            # Create backup before saving
            if self.config_file.exists():
                self._rotate_backups()

            data = self._config_to_dict(self.config)
            data['last_saved'] = datetime.now().isoformat()

            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)

            self._notify_observers()
            return True

        except Exception as e:
            print(f"Error saving config: {e}")
            return False

    def reset_to_defaults(self, section: Optional[str] = None) -> None:
        """Reset configuration to defaults."""
        if section is None:
            self.config = AppConfig()
        elif section == "vba_extractor":
            self.config.vba_extractor = VBAExtractorConfig()
        elif section == "python_analyzer":
            self.config.python_analyzer = PythonAnalyzerConfig()
        elif section == "folder_scanner":
            self.config.folder_scanner = FolderScannerConfig()
        elif section == "vba_optimizer":
            self.config.vba_optimizer = VBAOptimizerConfig()
        elif section == "ui":
            self.config.ui = UIConfig()
        elif section == "export":
            self.config.export = ExportConfig()

        self.save()

    def get(self, key: str, default: Any = None) -> Any:
        """Get a configuration value by dot-notation key."""
        try:
            parts = key.split('.')
            value = self.config
            for part in parts:
                if hasattr(value, part):
                    value = getattr(value, part)
                elif isinstance(value, dict):
                    value = value.get(part)
                else:
                    return default
            return value
        except Exception:
            return default

    def set(self, key: str, value: Any, auto_save: bool = True) -> None:
        """Set a configuration value by dot-notation key."""
        try:
            parts = key.split('.')
            obj = self.config

            for part in parts[:-1]:
                obj = getattr(obj, part)

            setattr(obj, parts[-1], value)

            if auto_save and self.config.ui.auto_save_config:
                self.save()

        except Exception as e:
            print(f"Error setting config {key}: {e}")

    def add_recent_file(self, file_path: str) -> None:
        """Add a file to recent files list."""
        recent = self.config.ui.recent_files

        # Remove if already exists
        if file_path in recent:
            recent.remove(file_path)

        # Add to front
        recent.insert(0, file_path)

        # Trim to max
        self.config.ui.recent_files = recent[:self.config.ui.max_recent_files]
        self.save()

    def add_observer(self, callback: callable) -> None:
        """Add an observer to be notified of config changes."""
        self._observers.append(callback)

    def remove_observer(self, callback: callable) -> None:
        """Remove an observer."""
        if callback in self._observers:
            self._observers.remove(callback)

    def _notify_observers(self) -> None:
        """Notify all observers of config change."""
        for callback in self._observers:
            try:
                callback(self.config)
            except Exception:
                pass

    def _rotate_backups(self) -> None:
        """Rotate backup files."""
        for i in range(self.BACKUP_COUNT - 1, 0, -1):
            old_backup = self.config_dir / f"config.backup.{i}.json"
            new_backup = self.config_dir / f"config.backup.{i + 1}.json"
            if old_backup.exists():
                shutil.move(str(old_backup), str(new_backup))

        # Create new backup
        backup_file = self.config_dir / "config.backup.1.json"
        shutil.copy2(str(self.config_file), str(backup_file))

    def _backup_corrupted(self) -> None:
        """Backup corrupted config file."""
        if self.config_file.exists():
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            corrupted_file = self.config_dir / f"config.corrupted.{timestamp}.json"
            shutil.move(str(self.config_file), str(corrupted_file))

    def _config_to_dict(self, config: Any) -> Dict:
        """Convert config dataclass to dictionary."""
        if hasattr(config, '__dataclass_fields__'):
            return {k: self._config_to_dict(v) for k, v in asdict(config).items()}
        elif isinstance(config, list):
            return [self._config_to_dict(item) for item in config]
        elif isinstance(config, dict):
            return {k: self._config_to_dict(v) for k, v in config.items()}
        return config

    def _dict_to_config(self, data: Dict) -> AppConfig:
        """Convert dictionary to config dataclass."""
        config = AppConfig()

        if 'vba_extractor' in data:
            config.vba_extractor = VBAExtractorConfig(**data['vba_extractor'])
        if 'python_analyzer' in data:
            config.python_analyzer = PythonAnalyzerConfig(**data['python_analyzer'])
        if 'folder_scanner' in data:
            config.folder_scanner = FolderScannerConfig(**data['folder_scanner'])
        if 'vba_optimizer' in data:
            config.vba_optimizer = VBAOptimizerConfig(**data['vba_optimizer'])
        if 'ui' in data:
            config.ui = UIConfig(**data['ui'])
        if 'export' in data:
            config.export = ExportConfig(**data['export'])

        config.version = data.get('version', '2.0.0')
        config.first_run = data.get('first_run', False)
        config.last_used = data.get('last_used', '')

        return config

    def export_config(self, file_path: str) -> bool:
        """Export configuration to a file."""
        try:
            data = self._config_to_dict(self.config)
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            return True
        except Exception:
            return False

    def import_config(self, file_path: str) -> bool:
        """Import configuration from a file."""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            self.config = self._dict_to_config(data)
            self.save()
            return True
        except Exception:
            return False


# Global config instance
_config_manager: Optional[ConfigManager] = None


def get_config() -> ConfigManager:
    """Get the global configuration manager instance."""
    global _config_manager
    if _config_manager is None:
        _config_manager = ConfigManager()
    return _config_manager
