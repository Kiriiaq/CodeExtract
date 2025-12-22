"""
VBA Optimizer Module - Optimize and clean VBA code.
Supports comment removal, auto-indentation, minification, and more.
"""

import re
from dataclasses import dataclass, field
from typing import Dict, List, Tuple, Optional
from enum import Enum, auto


class OptimizationType(Enum):
    """Types of optimization that can be applied."""
    REMOVE_COMMENTS = auto()
    AUTO_INDENT = auto()
    REMOVE_EMPTY_LINES = auto()
    RENAME_UNUSED_VARS = auto()
    MINIFY = auto()


@dataclass
class OptimizationOptions:
    """Options for VBA optimization."""
    remove_comments: bool = False
    auto_indent: bool = False
    remove_empty_lines: bool = False
    rename_unused_vars: bool = False
    minify: bool = False
    indent_size: int = 4
    backup_original: bool = True


@dataclass
class OptimizationResult:
    """Result of an optimization operation."""
    success: bool
    original_code: str
    optimized_code: str
    modifications: List[str] = field(default_factory=list)
    original_lines: int = 0
    optimized_lines: int = 0
    original_size: int = 0
    optimized_size: int = 0
    error_message: str = ""

    @property
    def size_reduction(self) -> float:
        """Calculate percentage of size reduction."""
        if self.original_size == 0:
            return 0.0
        return (self.original_size - self.optimized_size) / self.original_size * 100

    @property
    def line_reduction(self) -> float:
        """Calculate percentage of line reduction."""
        if self.original_lines == 0:
            return 0.0
        return (self.original_lines - self.optimized_lines) / self.original_lines * 100


class VBAOptimizer:
    """
    Optimize VBA code with various transformations.
    Supports comment removal, auto-indentation, minification, and more.
    """

    # VBA keywords that increase indentation
    INDENT_INCREASE = [
        r'^(Private |Public |Friend )?(Sub|Function|Property\s+(?:Get|Let|Set))\b',
        r'^If\b.*Then\s*$',
        r'^Select\s+Case\b',
        r'^For\b',
        r'^Do\b',
        r'^While\b',
        r'^With\b',
        r'^Type\b',
        r'^Enum\b',
    ]

    # VBA keywords that decrease indentation
    INDENT_DECREASE = [
        r'^End\s+(Sub|Function|Property|If|Select|With|Type|Enum)\b',
        r'^Next\b',
        r'^Loop\b',
        r'^Wend\b',
    ]

    # VBA keywords for same-line indent adjustment
    INDENT_SPECIAL = [
        r'^ElseIf\b',
        r'^Else\b',
        r'^Case\b',
    ]

    def __init__(self):
        self.options = OptimizationOptions()

    def optimize(self, code: str, options: Optional[OptimizationOptions] = None) -> OptimizationResult:
        """
        Optimize VBA code according to specified options.

        Args:
            code: The VBA code to optimize
            options: Optimization options (uses defaults if not provided)

        Returns:
            OptimizationResult with the optimized code
        """
        options = options or self.options
        result = OptimizationResult(
            success=True,
            original_code=code,
            optimized_code=code,
            original_lines=len(code.splitlines()),
            original_size=len(code)
        )

        try:
            optimized = code

            # Apply optimizations in order
            if options.remove_comments:
                optimized, count = self._remove_comments(optimized)
                if count > 0:
                    result.modifications.append(f"Removed {count} comments")

            if options.auto_indent:
                optimized = self._auto_indent(optimized, options.indent_size)
                result.modifications.append("Applied auto-indentation")

            if options.rename_unused_vars:
                optimized, count = self._rename_unused_variables(optimized)
                if count > 0:
                    result.modifications.append(f"Renamed {count} unused variables")

            if options.remove_empty_lines:
                optimized, count = self._remove_empty_lines(optimized)
                if count > 0:
                    result.modifications.append(f"Removed {count} empty lines")

            if options.minify:
                optimized = self._minify(optimized)
                result.modifications.append("Minified code")

            result.optimized_code = optimized
            result.optimized_lines = len(optimized.splitlines())
            result.optimized_size = len(optimized)

        except Exception as e:
            result.success = False
            result.error_message = str(e)

        return result

    def _remove_comments(self, code: str) -> Tuple[str, int]:
        """Remove VBA comments while preserving strings."""
        lines = code.split('\n')
        processed_lines = []
        removed_count = 0

        for line in lines:
            # Track if we're inside a string
            in_string = False
            comment_pos = -1

            for i, char in enumerate(line):
                if char == '"' and (i == 0 or line[i-1] != '\\'):
                    in_string = not in_string
                elif char == "'" and not in_string:
                    comment_pos = i
                    break

            if comment_pos >= 0:
                # Remove comment part
                new_line = line[:comment_pos].rstrip()
                if new_line or comment_pos == 0:
                    processed_lines.append(new_line)
                removed_count += 1
            else:
                processed_lines.append(line)

        return '\n'.join(processed_lines), removed_count

    def _auto_indent(self, code: str, indent_size: int = 4) -> str:
        """Apply automatic indentation to VBA code."""
        lines = code.split('\n')
        processed_lines = []
        indent_level = 0
        indent_str = ' ' * indent_size

        for line in lines:
            stripped = line.strip()

            if not stripped:
                processed_lines.append('')
                continue

            # Check for indent decrease (before adding line)
            for pattern in self.INDENT_DECREASE:
                if re.match(pattern, stripped, re.IGNORECASE):
                    indent_level = max(0, indent_level - 1)
                    break

            # Check for special keywords (temporary decrease)
            temp_decrease = False
            for pattern in self.INDENT_SPECIAL:
                if re.match(pattern, stripped, re.IGNORECASE):
                    temp_decrease = True
                    break

            # Add line with appropriate indentation
            if temp_decrease:
                processed_lines.append(indent_str * max(0, indent_level - 1) + stripped)
            else:
                processed_lines.append(indent_str * indent_level + stripped)

            # Check for indent increase (after adding line)
            for pattern in self.INDENT_INCREASE:
                if re.match(pattern, stripped, re.IGNORECASE):
                    indent_level += 1
                    break

        return '\n'.join(processed_lines)

    def _rename_unused_variables(self, code: str) -> Tuple[str, int]:
        """Rename unused variables with 'unused_' prefix."""
        lines = code.split('\n')

        # Find all variable declarations
        var_pattern = re.compile(r'^\s*Dim\s+(\w+)', re.IGNORECASE)
        declared_vars = set()

        for line in lines:
            match = var_pattern.match(line)
            if match:
                declared_vars.add(match.group(1))

        # Count usage of each variable
        var_usage = {var: 0 for var in declared_vars}

        for line in lines:
            if not var_pattern.match(line):
                for var in declared_vars:
                    var_usage[var] += len(re.findall(r'\b' + re.escape(var) + r'\b', line))

        # Find unused variables
        unused_vars = [var for var, count in var_usage.items() if count == 0]

        # Rename unused variables
        processed_code = code
        renamed_count = 0

        for var in unused_vars:
            new_name = f"unused_{var}"
            pattern = r'\bDim\s+' + re.escape(var) + r'\b'
            if re.search(pattern, processed_code, re.IGNORECASE):
                processed_code = re.sub(
                    pattern,
                    f'Dim {new_name}',
                    processed_code,
                    flags=re.IGNORECASE
                )
                renamed_count += 1

        return processed_code, renamed_count

    def _remove_empty_lines(self, code: str) -> Tuple[str, int]:
        """Remove excessive empty lines (keep max one consecutive)."""
        lines = code.split('\n')
        processed_lines = []
        prev_empty = False
        removed_count = 0

        for line in lines:
            is_empty = not line.strip()

            if is_empty:
                if not prev_empty:
                    processed_lines.append(line)
                else:
                    removed_count += 1
                prev_empty = True
            else:
                processed_lines.append(line)
                prev_empty = False

        return '\n'.join(processed_lines), removed_count

    def _minify(self, code: str) -> str:
        """Minify VBA code by removing unnecessary whitespace."""
        lines = code.split('\n')
        processed_lines = []

        for line in lines:
            stripped = line.strip()
            if not stripped:
                continue

            # Remove multiple spaces (but preserve strings)
            minified = self._minify_line(stripped)
            processed_lines.append(minified)

        return '\n'.join(processed_lines)

    def _minify_line(self, line: str) -> str:
        """Minify a single line while preserving strings."""
        result = []
        in_string = False
        prev_char = ''

        for char in line:
            if char == '"' and prev_char != '\\':
                in_string = not in_string
                result.append(char)
            elif in_string:
                result.append(char)
            elif char == ' ':
                if result and result[-1] != ' ':
                    result.append(char)
            else:
                result.append(char)
            prev_char = char

        return ''.join(result)

    def get_example(self, optimization_type: OptimizationType) -> Tuple[str, str]:
        """Get before/after example for an optimization type."""
        examples = {
            OptimizationType.REMOVE_COMMENTS: (
                "' This is a comment\nDim x As Integer ' inline comment\nx = 5",
                "Dim x As Integer\nx = 5"
            ),
            OptimizationType.AUTO_INDENT: (
                "Sub Test()\nDim x As Integer\nIf x > 0 Then\nMsgBox \"Positive\"\nEnd If\nEnd Sub",
                "Sub Test()\n    Dim x As Integer\n    If x > 0 Then\n        MsgBox \"Positive\"\n    End If\nEnd Sub"
            ),
            OptimizationType.REMOVE_EMPTY_LINES: (
                "Line 1\n\n\n\nLine 2\n\nLine 3",
                "Line 1\n\nLine 2\n\nLine 3"
            ),
            OptimizationType.RENAME_UNUSED_VARS: (
                "Dim unusedVar As String\nDim usedVar As Integer\nusedVar = 10",
                "Dim unused_unusedVar As String\nDim usedVar As Integer\nusedVar = 10"
            ),
            OptimizationType.MINIFY: (
                "Sub Test()\n    Dim x As Integer\n    x = 5\nEnd Sub",
                "Sub Test()\nDim x As Integer\nx = 5\nEnd Sub"
            )
        }
        return examples.get(optimization_type, ("", ""))

    def analyze_code(self, code: str) -> Dict[str, int]:
        """Analyze VBA code and return statistics."""
        lines = code.split('\n')

        # Count different elements
        total_lines = len(lines)
        empty_lines = sum(1 for line in lines if not line.strip())
        comment_lines = sum(1 for line in lines if line.strip().startswith("'"))

        # Count procedures
        proc_pattern = re.compile(
            r'^\s*(Public|Private|Friend)?\s*(Sub|Function|Property)',
            re.IGNORECASE | re.MULTILINE
        )
        procedures = len(proc_pattern.findall(code))

        # Count variables
        var_pattern = re.compile(r'^\s*Dim\s+', re.IGNORECASE | re.MULTILINE)
        variables = len(var_pattern.findall(code))

        return {
            "total_lines": total_lines,
            "code_lines": total_lines - empty_lines - comment_lines,
            "empty_lines": empty_lines,
            "comment_lines": comment_lines,
            "procedures": procedures,
            "variables": variables,
            "characters": len(code)
        }
