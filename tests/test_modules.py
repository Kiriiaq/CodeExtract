"""
Unit tests for CodeExtractPro modules.
"""

import os
import sys
import tempfile
import unittest
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent.parent / 'src'))

from modules.python_analyzer import PythonAnalyzer
from modules.folder_scanner import FolderScanner
from modules.vba_optimizer import VBAOptimizer, OptimizationOptions
from core.workflow import WorkflowManager, WorkflowStep, StepResult


class TestPythonAnalyzer(unittest.TestCase):
    """Tests for PythonAnalyzer module."""

    def setUp(self):
        self.analyzer = PythonAnalyzer()

    def test_analyze_simple_file(self):
        """Test analyzing a simple Python file."""
        with tempfile.NamedTemporaryFile(mode='w', suffix='.py', delete=False) as f:
            f.write('''
def hello():
    """Say hello."""
    print("Hello, World!")

class Greeter:
    """A greeter class."""
    def greet(self, name):
        return f"Hello, {name}!"
''')
            f.flush()
            temp_path = f.name

        try:
            analysis = self.analyzer.analyze_file(temp_path)

            self.assertIsNotNone(analysis)
            self.assertEqual(len(analysis.functions), 1)
            self.assertEqual(len(analysis.classes), 1)
            self.assertEqual(analysis.classes[0].name, 'Greeter')
            self.assertTrue(analysis.line_count > 0)

        finally:
            os.unlink(temp_path)

    def test_detect_imports(self):
        """Test import detection."""
        with tempfile.NamedTemporaryFile(mode='w', suffix='.py', delete=False) as f:
            f.write('''
import os
import sys
from pathlib import Path
from typing import List, Dict
''')
            f.flush()
            temp_path = f.name

        try:
            analysis = self.analyzer.analyze_file(temp_path)

            self.assertIn('os', analysis.dependencies)
            self.assertIn('sys', analysis.dependencies)
            self.assertIn('pathlib', analysis.dependencies)
            self.assertIn('typing', analysis.dependencies)

        finally:
            os.unlink(temp_path)


class TestFolderScanner(unittest.TestCase):
    """Tests for FolderScanner module."""

    def setUp(self):
        self.scanner = FolderScanner()

    def test_scan_empty_directory(self):
        """Test scanning an empty directory."""
        with tempfile.TemporaryDirectory() as tmpdir:
            result = self.scanner.scan(tmpdir)

            self.assertTrue(result.root_entry is not None)
            self.assertEqual(result.total_files, 0)

    def test_scan_with_files(self):
        """Test scanning a directory with files."""
        with tempfile.TemporaryDirectory() as tmpdir:
            # Create test files
            Path(tmpdir, 'test.txt').write_text('Hello')
            Path(tmpdir, 'test.py').write_text('print("test")')

            result = self.scanner.scan(tmpdir)

            self.assertEqual(result.total_files, 2)

    def test_generate_tree(self):
        """Test tree generation."""
        with tempfile.TemporaryDirectory() as tmpdir:
            Path(tmpdir, 'file.txt').write_text('content')
            subdir = Path(tmpdir, 'subdir')
            subdir.mkdir()
            Path(subdir, 'nested.txt').write_text('nested content')

            result = self.scanner.scan(tmpdir)
            tree = self.scanner.generate_tree(result)

            self.assertIn('file.txt', tree)
            self.assertIn('subdir', tree)
            self.assertIn('nested.txt', tree)


class TestVBAOptimizer(unittest.TestCase):
    """Tests for VBAOptimizer module."""

    def setUp(self):
        self.optimizer = VBAOptimizer()

    def test_remove_comments(self):
        """Test comment removal."""
        code = """' This is a comment
Dim x As Integer ' inline comment
x = 5"""

        options = OptimizationOptions(remove_comments=True)
        result = self.optimizer.optimize(code, options)

        self.assertTrue(result.success)
        self.assertNotIn("This is a comment", result.optimized_code)
        self.assertIn("Dim x As Integer", result.optimized_code)

    def test_auto_indent(self):
        """Test auto-indentation."""
        code = """Sub Test()
Dim x As Integer
If x > 0 Then
MsgBox "Positive"
End If
End Sub"""

        options = OptimizationOptions(auto_indent=True)
        result = self.optimizer.optimize(code, options)

        self.assertTrue(result.success)
        lines = result.optimized_code.split('\n')
        # Check that content inside Sub is indented
        self.assertTrue(any(line.startswith('    ') for line in lines))

    def test_remove_empty_lines(self):
        """Test empty line removal."""
        code = """Line 1



Line 2

Line 3"""

        options = OptimizationOptions(remove_empty_lines=True)
        result = self.optimizer.optimize(code, options)

        self.assertTrue(result.success)
        # Should have max 1 consecutive empty line
        self.assertNotIn('\n\n\n', result.optimized_code)


class TestWorkflowManager(unittest.TestCase):
    """Tests for WorkflowManager."""

    def test_add_step(self):
        """Test adding steps to workflow."""
        workflow = WorkflowManager()

        step = WorkflowStep(
            id="test_step",
            name="Test Step",
            description="A test step",
            function=lambda ctx: StepResult(True, "Success")
        )

        workflow.add_step(step)

        self.assertEqual(len(workflow.get_steps()), 1)
        self.assertEqual(workflow.get_step("test_step").name, "Test Step")

    def test_run_workflow(self):
        """Test running a simple workflow."""
        workflow = WorkflowManager()

        results = []

        def step1(ctx):
            results.append(1)
            return StepResult(True, "Step 1 done")

        def step2(ctx):
            results.append(2)
            return StepResult(True, "Step 2 done")

        workflow.add_step(WorkflowStep("s1", "Step 1", "First", step1))
        workflow.add_step(WorkflowStep("s2", "Step 2", "Second", step2))

        with tempfile.TemporaryDirectory() as tmpdir:
            workflow.output_base_dir = Path(tmpdir)
            workflow.run()

        self.assertEqual(results, [1, 2])

    def test_stop_workflow(self):
        """Test stopping a workflow."""
        workflow = WorkflowManager()

        def slow_step(ctx):
            import time
            for _ in range(10):
                if workflow.should_stop:
                    return StepResult(False, "Stopped")
                time.sleep(0.1)
            return StepResult(True, "Done")

        workflow.add_step(WorkflowStep("slow", "Slow Step", "Takes time", slow_step))

        import threading
        with tempfile.TemporaryDirectory() as tmpdir:
            workflow.output_base_dir = Path(tmpdir)

            thread = threading.Thread(target=workflow.run)
            thread.start()

            import time
            time.sleep(0.2)
            workflow.stop()

            thread.join(timeout=2.0)

        self.assertTrue(workflow.should_stop)


if __name__ == '__main__':
    unittest.main()
