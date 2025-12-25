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

    def extract_code_hierarchy(self, directory: str, include_subdirs: bool = True,
                                exclude_patterns: Optional[List[str]] = None,
                                exclude_dirs: Optional[List[str]] = None,
                                include_content: bool = True,
                                max_file_size_kb: int = 500) -> str:
        """
        Extract all Python code into a hierarchical text representation.

        Args:
            directory: Root directory to scan
            include_subdirs: Include subdirectories
            exclude_patterns: File patterns to exclude (e.g., ['test_*.py', '*_test.py'])
            exclude_dirs: Directories to exclude
            include_content: Include actual file content
            max_file_size_kb: Maximum file size to include content (in KB)

        Returns:
            Formatted string with hierarchical code representation
        """
        exclude_dirs = exclude_dirs or ['__pycache__', '.git', 'venv', '.venv',
                                         'node_modules', '.idea', '.vscode', 'dist', 'build']
        exclude_patterns = exclude_patterns or []

        output_lines = []
        root_path = Path(directory)

        # Header
        output_lines.append("=" * 80)
        output_lines.append(f"  CODE EXTRACTION - {root_path.name}")
        output_lines.append(f"  Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        output_lines.append("=" * 80)
        output_lines.append("")

        # Collect files
        files = []
        if include_subdirs:
            for root, dirs, filenames in os.walk(directory):
                dirs[:] = [d for d in dirs if d not in exclude_dirs]
                for filename in filenames:
                    if filename.endswith('.py'):
                        if not self._matches_exclude_pattern(filename, exclude_patterns):
                            files.append(os.path.join(root, filename))
        else:
            for filename in os.listdir(directory):
                if filename.endswith('.py'):
                    if not self._matches_exclude_pattern(filename, exclude_patterns):
                        files.append(os.path.join(directory, filename))

        files.sort()

        # Generate table of contents
        output_lines.append("TABLE OF CONTENTS")
        output_lines.append("-" * 40)
        for i, file_path in enumerate(files, 1):
            rel_path = os.path.relpath(file_path, directory)
            output_lines.append(f"  {i:3}. {rel_path}")
        output_lines.append("")
        output_lines.append(f"Total: {len(files)} Python files")
        output_lines.append("")

        # Process each file
        for file_path in files:
            rel_path = os.path.relpath(file_path, directory)
            file_size = os.path.getsize(file_path)

            output_lines.append("")
            output_lines.append("=" * 80)
            output_lines.append(f"FILE: {rel_path}")
            output_lines.append(f"Size: {file_size:,} bytes | Path: {file_path}")
            output_lines.append("=" * 80)

            # Analyze file structure
            analysis = self.analyze_file(file_path)

            # Show structure summary
            output_lines.append("")
            output_lines.append("STRUCTURE:")
            output_lines.append(f"  Lines: {analysis.line_count} (Code: {analysis.code_lines}, Comments: {analysis.comment_lines}, Blank: {analysis.blank_lines})")

            if analysis.imports or analysis.from_imports:
                output_lines.append(f"  Imports: {', '.join(analysis.imports + analysis.from_imports)}")

            if analysis.classes:
                output_lines.append(f"  Classes ({len(analysis.classes)}):")
                for cls in analysis.classes:
                    bases = f"({', '.join(cls.bases)})" if cls.bases else ""
                    output_lines.append(f"    - {cls.name}{bases} [line {cls.lineno}]")
                    for method in cls.methods:
                        args = ", ".join(method.args)
                        async_prefix = "async " if method.is_async else ""
                        output_lines.append(f"        {async_prefix}def {method.name}({args}) [line {method.lineno}]")

            if analysis.functions:
                output_lines.append(f"  Functions ({len(analysis.functions)}):")
                for func in analysis.functions:
                    args = ", ".join(func.args)
                    async_prefix = "async " if func.is_async else ""
                    ret = f" -> {func.return_type}" if func.return_type else ""
                    output_lines.append(f"    - {async_prefix}def {func.name}({args}){ret} [line {func.lineno}]")

            # Include content if requested and file not too large
            if include_content and file_size <= max_file_size_kb * 1024:
                output_lines.append("")
                output_lines.append("-" * 40)
                output_lines.append("CODE:")
                output_lines.append("-" * 40)

                try:
                    with open(file_path, 'r', encoding=analysis.encoding, errors='replace') as f:
                        content = f.read()

                    # Add line numbers
                    lines = content.split('\n')
                    max_line_num = len(str(len(lines)))
                    for i, line in enumerate(lines, 1):
                        output_lines.append(f"{i:>{max_line_num}} | {line}")
                except Exception as e:
                    output_lines.append(f"  [Error reading file: {e}]")

            elif include_content:
                output_lines.append("")
                output_lines.append(f"  [File too large: {file_size / 1024:.1f} KB > {max_file_size_kb} KB limit]")

        # Footer with summary
        output_lines.append("")
        output_lines.append("=" * 80)
        output_lines.append("SUMMARY")
        output_lines.append("=" * 80)

        total_lines = sum(self.analyze_file(f).line_count for f in files[:10])  # Sample for performance
        output_lines.append(f"Total files: {len(files)}")
        output_lines.append(f"Excluded directories: {', '.join(exclude_dirs)}")
        if exclude_patterns:
            output_lines.append(f"Excluded patterns: {', '.join(exclude_patterns)}")
        output_lines.append("")
        output_lines.append("=" * 80)
        output_lines.append("  END OF EXTRACTION")
        output_lines.append("=" * 80)

        return '\n'.join(output_lines)

    def _matches_exclude_pattern(self, filename: str, patterns: List[str]) -> bool:
        """Check if filename matches any exclude pattern."""
        import fnmatch
        for pattern in patterns:
            if fnmatch.fnmatch(filename, pattern):
                return True
        return False

    def save_code_extraction(self, directory: str, output_path: str, **kwargs) -> bool:
        """
        Extract code and save to a file.

        Args:
            directory: Root directory to scan
            output_path: Path to save the extraction
            **kwargs: Additional arguments for extract_code_hierarchy

        Returns:
            True if successful, False otherwise
        """
        try:
            content = self.extract_code_hierarchy(directory, **kwargs)
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(content)
            return True
        except Exception as e:
            print(f"Error saving extraction: {e}")
            return False
