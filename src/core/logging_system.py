"""
Logging System - Real-time logging with multiple outputs.
Supports console, file, and GUI callbacks with colored output.
"""

from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum
from pathlib import Path
from typing import Callable, Dict, List, Optional
import threading
import queue
import sys
import io


class LogLevel(Enum):
    """Log severity levels."""
    DEBUG = 0
    INFO = 1
    SUCCESS = 2
    WARNING = 3
    ERROR = 4
    CRITICAL = 5


@dataclass
class LogEntry:
    """A single log entry."""
    timestamp: datetime
    level: LogLevel
    message: str
    source: str = ""
    extra: Dict = field(default_factory=dict)

    def formatted(self, include_timestamp: bool = True, include_level: bool = True) -> str:
        """Format the log entry as a string."""
        parts = []
        if include_timestamp:
            parts.append(f"[{self.timestamp.strftime('%H:%M:%S')}]")
        if include_level:
            parts.append(f"[{self.level.name}]")
        if self.source:
            parts.append(f"[{self.source}]")
        parts.append(self.message)
        return " ".join(parts)


class LogManager:
    """
    Centralized logging system with real-time output capabilities.
    Supports multiple handlers: file, console, GUI callbacks.
    """

    # Color codes for terminal output
    COLORS = {
        LogLevel.DEBUG: "\033[90m",      # Gray
        LogLevel.INFO: "\033[97m",       # White
        LogLevel.SUCCESS: "\033[92m",    # Green
        LogLevel.WARNING: "\033[93m",    # Yellow
        LogLevel.ERROR: "\033[91m",      # Red
        LogLevel.CRITICAL: "\033[95m",   # Magenta
    }
    RESET = "\033[0m"

    def __init__(self, name: str = "CodeExtractPro"):
        self.name = name
        self.min_level = LogLevel.INFO
        self.entries: List[LogEntry] = []
        self.max_entries = 10000

        # Thread-safe queue for async logging
        self.log_queue: queue.Queue = queue.Queue()
        self._running = True

        # Output handlers
        self._file_path: Optional[Path] = None
        self._file_handle: Optional[io.TextIOWrapper] = None
        self._console_enabled = True
        self._use_colors = True
        self._callbacks: List[Callable[[LogEntry], None]] = []

        # Lock for thread safety
        self._lock = threading.Lock()

        # Start async logger thread
        self._logger_thread = threading.Thread(target=self._process_queue, daemon=True)
        self._logger_thread.start()

    def set_level(self, level: LogLevel) -> None:
        """Set minimum log level."""
        self.min_level = level

    def enable_console(self, enabled: bool = True, use_colors: bool = True) -> None:
        """Enable or disable console output."""
        self._console_enabled = enabled
        self._use_colors = use_colors

    def set_file(self, file_path: Optional[str]) -> None:
        """Set the log file path."""
        with self._lock:
            if self._file_handle:
                self._file_handle.close()
                self._file_handle = None

            if file_path:
                self._file_path = Path(file_path)
                self._file_path.parent.mkdir(parents=True, exist_ok=True)
                self._file_handle = open(self._file_path, 'a', encoding='utf-8')

    def add_callback(self, callback: Callable[[LogEntry], None]) -> None:
        """Add a callback for log entries."""
        self._callbacks.append(callback)

    def remove_callback(self, callback: Callable[[LogEntry], None]) -> None:
        """Remove a callback."""
        if callback in self._callbacks:
            self._callbacks.remove(callback)

    def clear_callbacks(self) -> None:
        """Remove all callbacks."""
        self._callbacks.clear()

    def _process_queue(self) -> None:
        """Background thread to process log entries."""
        while self._running:
            try:
                entry = self.log_queue.get(timeout=0.1)
                self._write_entry(entry)
            except queue.Empty:
                continue

    def _write_entry(self, entry: LogEntry) -> None:
        """Write a log entry to all outputs."""
        with self._lock:
            # Store entry
            self.entries.append(entry)
            if len(self.entries) > self.max_entries:
                self.entries = self.entries[-self.max_entries:]

            # Console output
            if self._console_enabled and sys.stdout:
                try:
                    if self._use_colors and sys.stdout.isatty():
                        color = self.COLORS.get(entry.level, "")
                        print(f"{color}{entry.formatted()}{self.RESET}")
                    else:
                        print(entry.formatted())
                except:
                    pass

            # File output
            if self._file_handle:
                try:
                    self._file_handle.write(entry.formatted() + "\n")
                    self._file_handle.flush()
                except:
                    pass

            # Callbacks
            for callback in self._callbacks:
                try:
                    callback(entry)
                except:
                    pass

    def log(self, message: str, level: LogLevel = LogLevel.INFO, source: str = "") -> None:
        """Log a message."""
        if level.value < self.min_level.value:
            return

        entry = LogEntry(
            timestamp=datetime.now(),
            level=level,
            message=message,
            source=source
        )
        self.log_queue.put(entry)

    def debug(self, message: str, source: str = "") -> None:
        """Log a debug message."""
        self.log(message, LogLevel.DEBUG, source)

    def info(self, message: str, source: str = "") -> None:
        """Log an info message."""
        self.log(message, LogLevel.INFO, source)

    def success(self, message: str, source: str = "") -> None:
        """Log a success message."""
        self.log(message, LogLevel.SUCCESS, source)

    def warning(self, message: str, source: str = "") -> None:
        """Log a warning message."""
        self.log(message, LogLevel.WARNING, source)

    def error(self, message: str, source: str = "") -> None:
        """Log an error message."""
        self.log(message, LogLevel.ERROR, source)

    def critical(self, message: str, source: str = "") -> None:
        """Log a critical message."""
        self.log(message, LogLevel.CRITICAL, source)

    def get_entries(self, level: Optional[LogLevel] = None,
                    source: Optional[str] = None,
                    limit: Optional[int] = None) -> List[LogEntry]:
        """Get log entries with optional filtering."""
        with self._lock:
            entries = self.entries.copy()

        if level:
            entries = [e for e in entries if e.level.value >= level.value]
        if source:
            entries = [e for e in entries if e.source == source]
        if limit:
            entries = entries[-limit:]

        return entries

    def clear(self) -> None:
        """Clear all log entries."""
        with self._lock:
            self.entries.clear()

    def export_to_file(self, file_path: str, level: Optional[LogLevel] = None) -> None:
        """Export logs to a file."""
        entries = self.get_entries(level=level)
        path = Path(file_path)
        path.parent.mkdir(parents=True, exist_ok=True)

        with open(path, 'w', encoding='utf-8') as f:
            f.write(f"# Log Export - {self.name}\n")
            f.write(f"# Generated: {datetime.now().isoformat()}\n")
            f.write(f"# Entries: {len(entries)}\n")
            f.write("=" * 80 + "\n\n")
            for entry in entries:
                f.write(entry.formatted() + "\n")

    def close(self) -> None:
        """Close the logger and release resources."""
        self._running = False
        self._logger_thread.join(timeout=1.0)
        if self._file_handle:
            self._file_handle.close()


# Global logger instance
_global_logger: Optional[LogManager] = None


def get_logger() -> LogManager:
    """Get the global logger instance."""
    global _global_logger
    if _global_logger is None:
        _global_logger = LogManager()
    return _global_logger


def set_logger(logger: LogManager) -> None:
    """Set the global logger instance."""
    global _global_logger
    _global_logger = logger
