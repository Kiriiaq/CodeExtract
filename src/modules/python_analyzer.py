"""
Python Analyzer Module - Analyze Python code structure and quality.
Extracts classes, functions, imports, and generates comprehensive reports.
"""

import ast
import os
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Set, Any
from concurrent.futures import ThreadPoolExecutor, as_completed


@dataclass
class FunctionInfo:
    """Information about a Python function."""
    name: str
    lineno: int
    args: List[str]
    return_type: Optional[str] = None
    decorators: List[str] = field(default_factory=list)
    docstring: Optional[str] = None
    is_async: bool = False
    is_method: bool = False
    complexity: int = 1


@dataclass
class ClassInfo:
    """Information about a Python class."""
    name: str
    lineno: int
    bases: List[str] = field(default_factory=list)
    methods: List[FunctionInfo] = field(default_factory=list)
    attributes: List[str] = field(default_factory=list)
    docstring: Optional[str] = None
    decorators: List[str] = field(default_factory=list)


@dataclass
class FileAnalysis:
    """Complete analysis of a Python file."""
    path: str
    name: str
    size: int
    modified: datetime
    line_count: int
    code_lines: int
    comment_lines: int
    blank_lines: int
    docstring_lines: int
    imports: List[str] = field(default_factory=list)
    from_imports: List[str] = field(default_factory=list)
    classes: List[ClassInfo] = field(default_factory=list)
    functions: List[FunctionInfo] = field(default_factory=list)
    global_variables: List[str] = field(default_factory=list)
    dependencies: Set[str] = field(default_factory=set)
    has_main: bool = False
    encoding: str = "utf-8"
    parse_error: Optional[str] = None

    @property
    def documentation_ratio(self) -> float:
        """Calculate the documentation ratio."""
        total = self.code_lines + self.comment_lines + self.docstring_lines
        if total == 0:
            return 0.0
        return (self.comment_lines + self.docstring_lines) / total * 100

    @property
    def total_functions(self) -> int:
        """Total number of functions including methods."""
        return len(self.functions) + sum(len(c.methods) for c in self.classes)


class PythonAnalyzer:
    """
    Analyze Python source files for structure, quality, and documentation.
    """

    STDLIB_MODULES = {
        'abc', 'argparse', 'ast', 'asyncio', 'base64', 'collections',
        'contextlib', 'copy', 'csv', 'dataclasses', 'datetime', 'decimal',
        'email', 'enum', 'functools', 'glob', 'hashlib', 'html', 'http',
        'importlib', 'inspect', 'io', 'itertools', 'json', 'logging',
        'math', 'multiprocessing', 'operator', 'os', 'pathlib', 'pickle',
        'platform', 'queue', 're', 'shutil', 'socket', 'sqlite3', 'string',
        'subprocess', 'sys', 'tempfile', 'threading', 'time', 'typing',
        'unittest', 'urllib', 'uuid', 'warnings', 'weakref', 'xml', 'zipfile'
    }

    def __init__(self, max_workers: int = 4):
        self.max_workers = max_workers

    def analyze_file(self, file_path: str) -> FileAnalysis:
        """
        Analyze a single Python file.

        Args:
            file_path: Path to the Python file

        Returns:
            FileAnalysis object with complete analysis
        """
        path = Path(file_path)
        stats = path.stat()

        # Read file content
        encoding = self._detect_encoding(file_path)
        try:
            with open(file_path, 'r', encoding=encoding, errors='replace') as f:
                content = f.read()
        except Exception as e:
            return FileAnalysis(
                path=str(path),
                name=path.name,
                size=stats.st_size,
                modified=datetime.fromtimestamp(stats.st_mtime),
                line_count=0,
                code_lines=0,
                comment_lines=0,
                blank_lines=0,
                docstring_lines=0,
                encoding=encoding,
                parse_error=str(e)
            )

        # Count line types
        lines = content.split('\n')
        line_count = len(lines)
        blank_lines = sum(1 for line in lines if not line.strip())
        comment_lines = sum(1 for line in lines if line.strip().startswith('#'))

        # Parse AST
        analysis = FileAnalysis(
            path=str(path),
            name=path.name,
            size=stats.st_size,
            modified=datetime.fromtimestamp(stats.st_mtime),
            line_count=line_count,
            code_lines=0,
            comment_lines=comment_lines,
            blank_lines=blank_lines,
            docstring_lines=0,
            encoding=encoding
        )

        try:
            tree = ast.parse(content)
            self._analyze_ast(tree, analysis, content)
            analysis.code_lines = line_count - blank_lines - comment_lines - analysis.docstring_lines
        except SyntaxError as e:
            analysis.parse_error = f"Syntax error at line {e.lineno}: {e.msg}"
            analysis.code_lines = line_count - blank_lines - comment_lines

        return analysis

    def _detect_encoding(self, file_path: str) -> str:
        """Detect file encoding."""
        encodings = ['utf-8', 'utf-8-sig', 'latin-1', 'cp1252']
        for encoding in encodings:
            try:
                with open(file_path, 'r', encoding=encoding) as f:
                    f.read(1024)
                return encoding
            except UnicodeDecodeError:
                continue
        return 'latin-1'

    def _analyze_ast(self, tree: ast.AST, analysis: FileAnalysis, content: str) -> None:
        """Analyze the AST of a Python file."""
        for node in ast.walk(tree):
            # Imports
            if isinstance(node, ast.Import):
                for alias in node.names:
                    analysis.imports.append(alias.name)
                    analysis.dependencies.add(alias.name.split('.')[0])

            elif isinstance(node, ast.ImportFrom):
                if node.module:
                    analysis.from_imports.append(node.module)
                    analysis.dependencies.add(node.module.split('.')[0])

            # Classes
            elif isinstance(node, ast.ClassDef):
                class_info = self._analyze_class(node)
                analysis.classes.append(class_info)

            # Top-level functions
            elif isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                if isinstance(node, ast.AsyncFunctionDef) or not self._is_method(node, tree):
                    func_info = self._analyze_function(node)
                    analysis.functions.append(func_info)

            # Global variables
            elif isinstance(node, ast.Assign):
                for target in node.targets:
                    if isinstance(target, ast.Name):
                        analysis.global_variables.append(target.id)

        # Check for __main__
        analysis.has_main = 'if __name__' in content

        # Count docstring lines
        analysis.docstring_lines = self._count_docstring_lines(tree)

    def _analyze_class(self, node: ast.ClassDef) -> ClassInfo:
        """Analyze a class definition."""
        bases = []
        for base in node.bases:
            try:
                bases.append(ast.unparse(base))
            except:
                bases.append("Unknown")

        methods = []
        attributes = []

        for item in node.body:
            if isinstance(item, (ast.FunctionDef, ast.AsyncFunctionDef)):
                func_info = self._analyze_function(item, is_method=True)
                methods.append(func_info)
            elif isinstance(item, ast.AnnAssign) and isinstance(item.target, ast.Name):
                attributes.append(item.target.id)
            elif isinstance(item, ast.Assign):
                for target in item.targets:
                    if isinstance(target, ast.Name):
                        attributes.append(target.id)

        decorators = [ast.unparse(d) for d in node.decorator_list]
        docstring = ast.get_docstring(node)

        return ClassInfo(
            name=node.name,
            lineno=node.lineno,
            bases=bases,
            methods=methods,
            attributes=attributes,
            docstring=docstring,
            decorators=decorators
        )

    def _analyze_function(self, node, is_method: bool = False) -> FunctionInfo:
        """Analyze a function definition."""
        args = []
        for arg in node.args.args:
            arg_str = arg.arg
            if arg.annotation:
                try:
                    arg_str += f": {ast.unparse(arg.annotation)}"
                except:
                    pass
            args.append(arg_str)

        return_type = None
        if node.returns:
            try:
                return_type = ast.unparse(node.returns)
            except:
                pass

        decorators = [ast.unparse(d) for d in node.decorator_list]
        docstring = ast.get_docstring(node)
        complexity = self._calculate_complexity(node)

        return FunctionInfo(
            name=node.name,
            lineno=node.lineno,
            args=args,
            return_type=return_type,
            decorators=decorators,
            docstring=docstring,
            is_async=isinstance(node, ast.AsyncFunctionDef),
            is_method=is_method,
            complexity=complexity
        )

    def _calculate_complexity(self, node: ast.AST) -> int:
        """Calculate cyclomatic complexity of a function."""
        complexity = 1
        for child in ast.walk(node):
            if isinstance(child, (ast.If, ast.While, ast.For, ast.ExceptHandler,
                                  ast.With, ast.Assert, ast.comprehension)):
                complexity += 1
            elif isinstance(child, ast.BoolOp):
                complexity += len(child.values) - 1
        return complexity

    def _is_method(self, node: ast.FunctionDef, tree: ast.AST) -> bool:
        """Check if a function is a method inside a class."""
        for item in ast.walk(tree):
            if isinstance(item, ast.ClassDef):
                for child in item.body:
                    if child is node:
                        return True
        return False

    def _count_docstring_lines(self, tree: ast.AST) -> int:
        """Count total docstring lines in the file."""
        count = 0
        for node in ast.walk(tree):
            if isinstance(node, (ast.Module, ast.ClassDef, ast.FunctionDef, ast.AsyncFunctionDef)):
                docstring = ast.get_docstring(node)
                if docstring:
                    count += len(docstring.split('\n'))
        return count

    def analyze_directory(self, directory: str, include_subdirs: bool = True,
                         pattern: Optional[str] = None,
                         exclude_dirs: Optional[List[str]] = None) -> List[FileAnalysis]:
        """
        Analyze all Python files in a directory.

        Args:
            directory: Root directory to analyze
            include_subdirs: Include subdirectories
            pattern: Regex pattern to filter files
            exclude_dirs: Directories to exclude

        Returns:
            List of FileAnalysis objects
        """
        exclude_dirs = exclude_dirs or ['__pycache__', '.git', 'venv', '.venv', 'node_modules']
        files = []

        if include_subdirs:
            for root, dirs, filenames in os.walk(directory):
                dirs[:] = [d for d in dirs if d not in exclude_dirs]
                for filename in filenames:
                    if filename.endswith('.py'):
                        if pattern is None or re.search(pattern, filename):
                            files.append(os.path.join(root, filename))
        else:
            for filename in os.listdir(directory):
                if filename.endswith('.py'):
                    if pattern is None or re.search(pattern, filename):
                        files.append(os.path.join(directory, filename))

        # Analyze files in parallel
        results = []
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            futures = {executor.submit(self.analyze_file, f): f for f in files}
            for future in as_completed(futures):
                try:
                    results.append(future.result())
                except Exception as e:
                    file_path = futures[future]
                    results.append(FileAnalysis(
                        path=file_path,
                        name=os.path.basename(file_path),
                        size=0,
                        modified=datetime.now(),
                        line_count=0,
                        code_lines=0,
                        comment_lines=0,
                        blank_lines=0,
                        docstring_lines=0,
                        parse_error=str(e)
                    ))

        return sorted(results, key=lambda x: x.path)

    def get_external_dependencies(self, analyses: List[FileAnalysis]) -> Set[str]:
        """Get all external (non-stdlib) dependencies."""
        all_deps = set()
        for analysis in analyses:
            all_deps.update(analysis.dependencies)
        return all_deps - self.STDLIB_MODULES

    def generate_summary(self, analyses: List[FileAnalysis]) -> Dict[str, Any]:
        """Generate a summary of the analysis."""
        total_lines = sum(a.line_count for a in analyses)
        total_code = sum(a.code_lines for a in analyses)
        total_comments = sum(a.comment_lines for a in analyses)
        total_classes = sum(len(a.classes) for a in analyses)
        total_functions = sum(a.total_functions for a in analyses)

        return {
            "total_files": len(analyses),
            "total_lines": total_lines,
            "total_code_lines": total_code,
            "total_comment_lines": total_comments,
            "total_classes": total_classes,
            "total_functions": total_functions,
            "average_lines_per_file": total_lines / len(analyses) if analyses else 0,
            "documentation_ratio": (total_comments / total_code * 100) if total_code else 0,
            "external_dependencies": list(self.get_external_dependencies(analyses)),
            "files_with_errors": [a.path for a in analyses if a.parse_error]
        }
