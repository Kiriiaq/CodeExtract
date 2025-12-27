"""
Integration tests for CodeExtractPro.
Tests interactions between modules.
"""

import os
import sys
import tempfile
import time
import unittest
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent.parent / 'src'))

from core.config_manager import ConfigManager
from core.export_manager import ExportManager
from core.logging_system import LogManager, LogLevel
from core.workflow import WorkflowManager, WorkflowStep, StepResult
from modules.python_analyzer import PythonAnalyzer
from modules.folder_scanner import FolderScanner
from modules.vba_optimizer import VBAOptimizer, OptimizationOptions


class TestConfigAndLoggingIntegration(unittest.TestCase):
    """Tests for ConfigManager and LogManager integration."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.config = ConfigManager(config_dir=Path(self.temp_dir))
        self.logger = LogManager(name="IntegrationTest")
        self.logger.enable_console(False)

    def tearDown(self):
        self.logger.close()
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_log_config_changes(self):
        """Test logging configuration changes."""
        log_entries = []

        def callback(entry):
            log_entries.append(entry)

        self.logger.add_callback(callback)

        # Simulate logging config changes
        self.config.set("ui.theme", "light", auto_save=False)
        self.logger.info(f"Config changed: ui.theme = light")
        time.sleep(0.2)

        self.assertTrue(any("theme" in e.message for e in log_entries))

    def test_config_observer_with_logging(self):
        """Test config observer that logs changes."""
        log_entries = []

        def log_callback(entry):
            log_entries.append(entry)

        def config_observer(config):
            self.logger.info("Configuration saved")

        self.logger.add_callback(log_callback)
        self.config.add_observer(config_observer)

        self.config.save()
        time.sleep(0.2)

        self.assertTrue(any("saved" in e.message for e in log_entries))


class TestAnalyzerAndExportIntegration(unittest.TestCase):
    """Tests for PythonAnalyzer and ExportManager integration."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.analyzer = PythonAnalyzer()
        self.exporter = ExportManager()

    def tearDown(self):
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_analyze_and_export_json(self):
        """Test analyzing Python files and exporting to JSON."""
        # Create a Python file
        py_file = os.path.join(self.temp_dir, "sample.py")
        Path(py_file).write_text('''
def hello():
    """Greet the world."""
    print("Hello")

class Greeter:
    def greet(self, name):
        return f"Hello, {name}"
''')

        # Analyze
        analyses = self.analyzer.analyze_directory(self.temp_dir)
        summary = self.analyzer.generate_summary(analyses)

        # Export
        export_path = os.path.join(self.temp_dir, "analysis.json")
        export_data = {"summary": summary, "files": len(analyses)}
        result = self.exporter.export(export_data, export_path, 'json')

        self.assertTrue(result.success)
        self.assertTrue(os.path.exists(export_path))

    def test_analyze_and_export_html(self):
        """Test analyzing and exporting to HTML."""
        # Create test files
        py_file = os.path.join(self.temp_dir, "module.py")
        Path(py_file).write_text('''
import os

class FileHandler:
    def read(self, path):
        return open(path).read()
''')

        analyses = self.analyzer.analyze_directory(self.temp_dir)
        summary = self.analyzer.generate_summary(analyses)

        export_path = os.path.join(self.temp_dir, "report.html")
        result = self.exporter.export(
            {"summary": summary, "files": [{"name": a.name, "lines": a.line_count} for a in analyses]},
            export_path,
            'html'
        )

        self.assertTrue(result.success)
        with open(export_path, 'r') as f:
            content = f.read()
        self.assertIn("<!DOCTYPE html>", content)


class TestScannerAndExportIntegration(unittest.TestCase):
    """Tests for FolderScanner and ExportManager integration."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.scanner = FolderScanner()
        self.exporter = ExportManager()

    def tearDown(self):
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_scan_and_export_csv(self):
        """Test scanning folder and exporting to CSV."""
        # Create test structure
        os.makedirs(os.path.join(self.temp_dir, "subdir"))
        Path(os.path.join(self.temp_dir, "file1.txt")).write_text("content1")
        Path(os.path.join(self.temp_dir, "file2.py")).write_text("print('test')")
        Path(os.path.join(self.temp_dir, "subdir", "file3.md")).write_text("# Title")

        # Scan
        scan_result = self.scanner.scan(self.temp_dir)

        # Export
        export_data = {
            "files": [{"name": f"file{i}", "size": 100} for i in range(scan_result.total_files)]
        }
        export_path = os.path.join(self.temp_dir, "scan.csv")
        result = self.exporter.export(export_data, export_path, 'csv')

        self.assertTrue(result.success)
        self.assertTrue(os.path.exists(export_path))


class TestWorkflowIntegration(unittest.TestCase):
    """Tests for WorkflowManager with other modules."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.workflow = WorkflowManager()
        self.workflow.output_base_dir = Path(self.temp_dir)

    def tearDown(self):
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_workflow_with_analyzer(self):
        """Test workflow that uses PythonAnalyzer."""
        analyzer = PythonAnalyzer()
        results = {}

        # Create test file
        py_file = os.path.join(self.temp_dir, "test.py")
        Path(py_file).write_text("def func(): pass")

        def analyze_step(ctx):
            analyses = analyzer.analyze_directory(self.temp_dir)
            results['count'] = len(analyses)
            return StepResult(True, f"Analyzed {len(analyses)} files")

        self.workflow.add_step(WorkflowStep(
            id="analyze",
            name="Analyze Python",
            description="Analyze Python files",
            function=analyze_step
        ))

        self.workflow.run()

        self.assertEqual(results['count'], 1)

    def test_workflow_with_scanner(self):
        """Test workflow that uses FolderScanner."""
        scanner = FolderScanner()
        results = {}

        # Create test files
        Path(os.path.join(self.temp_dir, "a.txt")).write_text("a")
        Path(os.path.join(self.temp_dir, "b.txt")).write_text("b")

        def scan_step(ctx):
            result = scanner.scan(self.temp_dir)
            results['files'] = result.total_files
            return StepResult(True, f"Scanned {result.total_files} files")

        self.workflow.add_step(WorkflowStep(
            id="scan",
            name="Scan Folder",
            description="Scan folder structure",
            function=scan_step
        ))

        self.workflow.run()

        self.assertGreaterEqual(results['files'], 2)

    def test_multi_step_workflow(self):
        """Test workflow with multiple dependent steps."""
        context_data = {}

        def step1(ctx):
            context_data['step1'] = True
            return StepResult(True, "Step 1 complete")

        def step2(ctx):
            if not context_data.get('step1'):
                return StepResult(False, "Step 1 not complete")
            context_data['step2'] = True
            return StepResult(True, "Step 2 complete")

        def step3(ctx):
            if not context_data.get('step2'):
                return StepResult(False, "Step 2 not complete")
            context_data['step3'] = True
            return StepResult(True, "Step 3 complete")

        self.workflow.add_step(WorkflowStep("s1", "Step 1", "First step", step1))
        self.workflow.add_step(WorkflowStep("s2", "Step 2", "Second step", step2))
        self.workflow.add_step(WorkflowStep("s3", "Step 3", "Third step", step3))

        self.workflow.run()

        self.assertTrue(context_data.get('step1'))
        self.assertTrue(context_data.get('step2'))
        self.assertTrue(context_data.get('step3'))


class TestOptimizerAndExportIntegration(unittest.TestCase):
    """Tests for VBAOptimizer and ExportManager integration."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.optimizer = VBAOptimizer()
        self.exporter = ExportManager()

    def tearDown(self):
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_optimize_and_export(self):
        """Test optimizing VBA code and exporting results."""
        code = '''
' Comment to remove
Sub TestProcedure()
Dim x As Integer


x = 5
End Sub
'''
        options = OptimizationOptions(remove_comments=True, remove_empty_lines=True)
        result = self.optimizer.optimize(code, options)

        # Export result
        export_data = {
            "original_lines": result.original_lines,
            "optimized_lines": result.optimized_lines,
            "modifications": result.modifications
        }
        export_path = os.path.join(self.temp_dir, "optimization.json")
        export_result = self.exporter.export(export_data, export_path, 'json')

        self.assertTrue(result.success)
        self.assertTrue(export_result.success)


class TestEndToEndAnalysisWorkflow(unittest.TestCase):
    """End-to-end test for complete analysis workflow."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        # Create a mini project structure
        src_dir = os.path.join(self.temp_dir, "src")
        tests_dir = os.path.join(self.temp_dir, "tests")
        os.makedirs(src_dir)
        os.makedirs(tests_dir)

        # Create source files
        Path(os.path.join(src_dir, "main.py")).write_text('''
"""Main module."""
import os

def main():
    print("Hello")

if __name__ == "__main__":
    main()
''')

        Path(os.path.join(src_dir, "utils.py")).write_text('''
"""Utility functions."""

def helper(x):
    return x * 2

class DataHandler:
    def process(self, data):
        return data
''')

        Path(os.path.join(tests_dir, "test_main.py")).write_text('''
import unittest

class TestMain(unittest.TestCase):
    def test_pass(self):
        self.assertTrue(True)
''')

    def tearDown(self):
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_complete_analysis_pipeline(self):
        """Test complete analysis pipeline: scan, analyze, export."""
        # Step 1: Scan directory
        scanner = FolderScanner()
        scan_result = scanner.scan(self.temp_dir)
        self.assertGreater(scan_result.total_files, 0)

        # Step 2: Analyze Python files
        analyzer = PythonAnalyzer()
        analyses = analyzer.analyze_directory(self.temp_dir, include_subdirs=True)
        summary = analyzer.generate_summary(analyses)
        self.assertGreater(summary['total_files'], 0)

        # Step 3: Export results
        exporter = ExportManager()

        # Export JSON
        json_path = os.path.join(self.temp_dir, "report.json")
        json_result = exporter.export({
            "scan": {"files": scan_result.total_files, "dirs": scan_result.total_directories},
            "analysis": summary
        }, json_path, 'json')
        self.assertTrue(json_result.success)

        # Export HTML
        html_path = os.path.join(self.temp_dir, "report.html")
        html_result = exporter.export({
            "summary": summary,
            "files": [{"name": a.name, "lines": a.line_count} for a in analyses]
        }, html_path, 'html')
        self.assertTrue(html_result.success)

        # Verify exports
        self.assertTrue(os.path.exists(json_path))
        self.assertTrue(os.path.exists(html_path))


if __name__ == '__main__':
    unittest.main()
