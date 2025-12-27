"""
Unit tests for LoggingSystem module.
Tests logging, filtering, callbacks, and file output.
"""

import os
import sys
import tempfile
import threading
import time
import unittest
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent.parent / 'src'))

from core.logging_system import LogManager, LogLevel, LogEntry, get_logger


class TestLogLevel(unittest.TestCase):
    """Tests for LogLevel enum."""

    def test_level_ordering(self):
        """Test that log levels are ordered correctly."""
        self.assertLess(LogLevel.DEBUG.value, LogLevel.INFO.value)
        self.assertLess(LogLevel.INFO.value, LogLevel.WARNING.value)
        self.assertLess(LogLevel.WARNING.value, LogLevel.ERROR.value)
        self.assertLess(LogLevel.ERROR.value, LogLevel.CRITICAL.value)

    def test_all_levels_exist(self):
        """Test that all expected log levels exist."""
        levels = [LogLevel.DEBUG, LogLevel.INFO, LogLevel.SUCCESS,
                  LogLevel.WARNING, LogLevel.ERROR, LogLevel.CRITICAL]
        self.assertEqual(len(levels), 6)


class TestLogEntry(unittest.TestCase):
    """Tests for LogEntry dataclass."""

    def test_create_entry(self):
        """Test creating a log entry."""
        from datetime import datetime
        entry = LogEntry(
            timestamp=datetime.now(),
            level=LogLevel.INFO,
            message="Test message",
            source="test"
        )
        self.assertEqual(entry.message, "Test message")
        self.assertEqual(entry.level, LogLevel.INFO)
        self.assertEqual(entry.source, "test")

    def test_formatted_output(self):
        """Test formatted string output."""
        from datetime import datetime
        entry = LogEntry(
            timestamp=datetime(2024, 1, 15, 10, 30, 0),
            level=LogLevel.WARNING,
            message="Warning message"
        )
        formatted = entry.formatted()
        self.assertIn("[10:30:00]", formatted)
        self.assertIn("[WARNING]", formatted)
        self.assertIn("Warning message", formatted)

    def test_formatted_without_timestamp(self):
        """Test formatted output without timestamp."""
        from datetime import datetime
        entry = LogEntry(
            timestamp=datetime.now(),
            level=LogLevel.INFO,
            message="Test"
        )
        formatted = entry.formatted(include_timestamp=False)
        self.assertNotIn(":", formatted.split("]")[0])  # No time separator

    def test_formatted_with_source(self):
        """Test formatted output with source."""
        from datetime import datetime
        entry = LogEntry(
            timestamp=datetime.now(),
            level=LogLevel.INFO,
            message="Test",
            source="MyModule"
        )
        formatted = entry.formatted()
        self.assertIn("[MyModule]", formatted)


class TestLogManager(unittest.TestCase):
    """Tests for LogManager class."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.logger = LogManager(name="TestLogger")
        self.logger.enable_console(False)  # Disable console output for tests

    def tearDown(self):
        self.logger.close()
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_log_info(self):
        """Test logging info message."""
        self.logger.info("Test info message")
        time.sleep(0.2)  # Wait for async processing
        entries = self.logger.get_entries()
        self.assertTrue(any(e.message == "Test info message" for e in entries))

    def test_log_all_levels(self):
        """Test logging at all levels."""
        self.logger.set_level(LogLevel.DEBUG)
        self.logger.debug("Debug message")
        self.logger.info("Info message")
        self.logger.success("Success message")
        self.logger.warning("Warning message")
        self.logger.error("Error message")
        self.logger.critical("Critical message")
        time.sleep(0.3)

        entries = self.logger.get_entries()
        levels = {e.level for e in entries}
        self.assertIn(LogLevel.DEBUG, levels)
        self.assertIn(LogLevel.INFO, levels)
        self.assertIn(LogLevel.CRITICAL, levels)

    def test_level_filtering(self):
        """Test that messages below min level are filtered."""
        self.logger.set_level(LogLevel.WARNING)
        self.logger.debug("Debug - should not appear")
        self.logger.info("Info - should not appear")
        self.logger.warning("Warning - should appear")
        time.sleep(0.2)

        entries = self.logger.get_entries()
        messages = [e.message for e in entries]
        self.assertNotIn("Debug - should not appear", messages)
        self.assertNotIn("Info - should not appear", messages)
        self.assertIn("Warning - should appear", messages)

    def test_get_entries_by_level(self):
        """Test filtering entries by level."""
        self.logger.set_level(LogLevel.DEBUG)
        self.logger.debug("Debug")
        self.logger.info("Info")
        self.logger.error("Error")
        time.sleep(0.2)

        entries = self.logger.get_entries(level=LogLevel.ERROR)
        self.assertTrue(all(e.level.value >= LogLevel.ERROR.value for e in entries))

    def test_get_entries_by_source(self):
        """Test filtering entries by source."""
        self.logger.info("Message 1", source="ModuleA")
        self.logger.info("Message 2", source="ModuleB")
        self.logger.info("Message 3", source="ModuleA")
        time.sleep(0.2)

        entries = self.logger.get_entries(source="ModuleA")
        self.assertEqual(len(entries), 2)
        self.assertTrue(all(e.source == "ModuleA" for e in entries))

    def test_get_entries_with_limit(self):
        """Test limiting returned entries."""
        for i in range(10):
            self.logger.info(f"Message {i}")
        time.sleep(0.3)

        entries = self.logger.get_entries(limit=5)
        self.assertEqual(len(entries), 5)

    def test_clear_entries(self):
        """Test clearing all entries."""
        self.logger.info("Message 1")
        self.logger.info("Message 2")
        time.sleep(0.2)

        self.logger.clear()
        entries = self.logger.get_entries()
        self.assertEqual(len(entries), 0)

    def test_callback_invocation(self):
        """Test that callbacks are invoked for each log entry."""
        callback_entries = []

        def callback(entry):
            callback_entries.append(entry)

        self.logger.add_callback(callback)
        self.logger.info("Test callback")
        time.sleep(0.2)

        self.assertTrue(any(e.message == "Test callback" for e in callback_entries))

    def test_remove_callback(self):
        """Test removing a callback."""
        callback_count = [0]

        def callback(entry):
            callback_count[0] += 1

        self.logger.add_callback(callback)
        self.logger.info("Message 1")
        time.sleep(0.2)
        initial_count = callback_count[0]

        self.logger.remove_callback(callback)
        self.logger.info("Message 2")
        time.sleep(0.2)

        self.assertEqual(callback_count[0], initial_count)

    def test_clear_callbacks(self):
        """Test clearing all callbacks."""
        callback_called = [False]

        def callback(entry):
            callback_called[0] = True

        self.logger.add_callback(callback)
        self.logger.clear_callbacks()
        self.logger.info("Test")
        time.sleep(0.2)

        self.assertFalse(callback_called[0])

    def test_file_logging(self):
        """Test logging to file."""
        log_file = os.path.join(self.temp_dir, "test.log")
        self.logger.set_file(log_file)
        self.logger.info("File log message")
        time.sleep(0.2)

        self.assertTrue(os.path.exists(log_file))
        with open(log_file, 'r') as f:
            content = f.read()
        self.assertIn("File log message", content)

    def test_export_to_file(self):
        """Test exporting logs to file."""
        self.logger.info("Export message 1")
        self.logger.warning("Export message 2")
        time.sleep(0.2)

        export_file = os.path.join(self.temp_dir, "export.log")
        self.logger.export_to_file(export_file)

        self.assertTrue(os.path.exists(export_file))
        with open(export_file, 'r') as f:
            content = f.read()
        self.assertIn("Export message 1", content)
        self.assertIn("Export message 2", content)
        self.assertIn("# Log Export", content)

    def test_max_entries_limit(self):
        """Test that entries are trimmed to max limit."""
        self.logger.max_entries = 5
        for i in range(10):
            self.logger.info(f"Message {i}")
        time.sleep(0.3)

        entries = self.logger.get_entries()
        self.assertLessEqual(len(entries), 5)

    def test_thread_safety(self):
        """Test thread-safe logging."""
        num_threads = 5
        messages_per_thread = 10

        def log_messages(thread_id):
            for i in range(messages_per_thread):
                self.logger.info(f"Thread {thread_id} message {i}")

        threads = [threading.Thread(target=log_messages, args=(i,))
                   for i in range(num_threads)]
        for t in threads:
            t.start()
        for t in threads:
            t.join()

        time.sleep(0.5)  # Wait for async processing

        entries = self.logger.get_entries()
        # Should have all messages (threads * messages_per_thread)
        self.assertEqual(len(entries), num_threads * messages_per_thread)


class TestGlobalLogger(unittest.TestCase):
    """Tests for global logger instance."""

    def test_get_logger_returns_instance(self):
        """Test that get_logger returns a LogManager."""
        logger = get_logger()
        self.assertIsInstance(logger, LogManager)

    def test_get_logger_singleton(self):
        """Test that get_logger returns same instance."""
        logger1 = get_logger()
        logger2 = get_logger()
        self.assertIs(logger1, logger2)


if __name__ == '__main__':
    unittest.main()
