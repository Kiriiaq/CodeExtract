"""
VBA Analyzer Module - Advanced VBA code analysis with regex patterns.
Extracts procedures, variables, constants with scope information.
Provides matplotlib/seaborn visualizations and pandas DataFrame export.
"""

import re
import os
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any
from enum import Enum

# Optional imports with graceful fallback
PANDAS_AVAILABLE = False
MATPLOTLIB_AVAILABLE = False

try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    pass

try:
    import matplotlib.pyplot as plt
    import matplotlib
    matplotlib.use('Agg')  # Use non-interactive backend by default
    MATPLOTLIB_AVAILABLE = True
    try:
        import seaborn as sns
        sns.set_palette("husl")
    except ImportError:
        pass
except ImportError:
    pass


class VBAElementType(Enum):
    """Types of VBA elements."""
    PROCEDURE = "Procedure"
    VARIABLE = "Variable"
    CONSTANT = "Constant"
    TYPE_DEFINITION = "Type"
    ENUM_DEFINITION = "Enum"
    API_DECLARATION = "API"


class VBAScope(Enum):
    """VBA scope modifiers."""
    PUBLIC = "Public"
    PRIVATE = "Private"
    FRIEND = "Friend"
    GLOBAL = "Global"
    DIM = "Dim"
    STATIC = "Static"


@dataclass
class VBAProcedure:
    """Represents a VBA procedure (Sub, Function, Property)."""
    name: str
    procedure_type: str  # Sub, Function, Property Get/Let/Set
    scope: str
    parameters: str
    return_type: Optional[str]
    module_name: str
    line_number: int
    signature: str

    def to_dict(self) -> Dict[str, Any]:
        return {
            'name': self.name,
            'type': self.procedure_type,
            'scope': self.scope,
            'parameters': self.parameters,
            'return_type': self.return_type or '',
            'module': self.module_name,
            'line': self.line_number,
            'signature': self.signature
        }


@dataclass
class VBAVariable:
    """Represents a VBA variable or constant."""
    name: str
    var_type: str
    declaration: str  # Dim, Private, Public, Global, Static, Const
    scope: str
    value: Optional[str]  # For constants
    module_name: str
    procedure_name: Optional[str]  # None if module-level
    line_number: int
    source: str

    def to_dict(self) -> Dict[str, Any]:
        return {
            'name': self.name,
            'var_type': self.var_type,
            'declaration': self.declaration,
            'scope': self.scope,
            'value': self.value or '',
            'module': self.module_name,
            'procedure': self.procedure_name or '',
            'line': self.line_number,
            'source': self.source
        }


@dataclass
class VBAAnalysisResult:
    """Result of VBA code analysis."""
    success: bool
    source_file: str
    module_name: str
    procedures: List[VBAProcedure] = field(default_factory=list)
    variables: List[VBAVariable] = field(default_factory=list)
    error_message: str = ""
    analysis_time: float = 0.0

    @property
    def total_procedures(self) -> int:
        return len(self.procedures)

    @property
    def total_variables(self) -> int:
        return len(self.variables)

    @property
    def total_constants(self) -> int:
        return len([v for v in self.variables if v.declaration == 'Const'])

    def to_dict(self) -> Dict[str, Any]:
        return {
            'source_file': self.source_file,
            'module_name': self.module_name,
            'procedures': [p.to_dict() for p in self.procedures],
            'variables': [v.to_dict() for v in self.variables],
            'stats': {
                'total_procedures': self.total_procedures,
                'total_variables': self.total_variables,
                'total_constants': self.total_constants
            }
        }


class VBAAnalyzer:
    """
    Advanced VBA code analyzer with regex-based extraction.
    Extracts procedures, variables, constants with full scope information.
    """

    def __init__(self):
        # Regex patterns for VBA code analysis
        self.patterns = {
            # Procedures: Sub, Function, Property Get/Let/Set
            'procedure': re.compile(
                r'^\s*(Public|Private|Friend)?\s*'
                r'(Sub|Function|Property\s+(?:Get|Let|Set))\s+'
                r'(\w+)\s*'
                r'\(([^)]*)\)'
                r'(?:\s+As\s+(\w+))?',
                re.IGNORECASE | re.MULTILINE
            ),
            # Variables with type: Dim/Private/Public/Global/Static variable As Type
            'variable': re.compile(
                r'^\s*(Dim|Private|Public|Global|Static)\s+'
                r'(\w+(?:\s*,\s*\w+)*)\s+'
                r'As\s+(\w+(?:\([^)]*\))?)',
                re.IGNORECASE | re.MULTILINE
            ),
            # Constants with value: Const NAME As Type = Value
            'const_value': re.compile(
                r'^\s*(Public|Private)?\s*Const\s+'
                r'(\w+)\s+'
                r'As\s+(\w+)\s*=\s*(.+)$',
                re.IGNORECASE | re.MULTILINE
            ),
            # Simple constants without type: Const NAME = Value
            'const_simple': re.compile(
                r'^\s*(Public|Private)?\s*Const\s+'
                r'(\w+)\s*=\s*(.+)$',
                re.IGNORECASE | re.MULTILINE
            ),
            # Type definitions
            'type_def': re.compile(
                r'^\s*(Public|Private)?\s*Type\s+(\w+)',
                re.IGNORECASE | re.MULTILINE
            ),
            # Enum definitions
            'enum_def': re.compile(
                r'^\s*(Public|Private)?\s*Enum\s+(\w+)',
                re.IGNORECASE | re.MULTILINE
            ),
            # API declarations
            'api_declare': re.compile(
                r'^\s*(Public|Private)?\s*Declare\s+'
                r'(PtrSafe\s+)?(Sub|Function)\s+'
                r'(\w+)\s+Lib\s+"([^"]+)"',
                re.IGNORECASE | re.MULTILINE
            ),
            # Module-level Option statements
            'option': re.compile(
                r'^\s*Option\s+(Explicit|Base|Compare|Private)',
                re.IGNORECASE | re.MULTILINE
            )
        }

        # Module types mapping
        self.module_types = {
            1: "Module Standard",
            2: "Module de Classe",
            3: "UserForm",
            100: "Document"
        }

    def analyze_code(self, code: str, module_name: str = "Module1",
                     source_file: str = "") -> VBAAnalysisResult:
        """
        Analyze VBA code and extract all elements.

        Args:
            code: VBA source code
            module_name: Name of the module
            source_file: Source file path

        Returns:
            VBAAnalysisResult with extracted procedures and variables
        """
        start_time = datetime.now()

        try:
            procedures = self._extract_procedures(code, module_name)
            variables = self._extract_variables(code, module_name, procedures)

            analysis_time = (datetime.now() - start_time).total_seconds()

            return VBAAnalysisResult(
                success=True,
                source_file=source_file,
                module_name=module_name,
                procedures=procedures,
                variables=variables,
                analysis_time=analysis_time
            )

        except Exception as e:
            return VBAAnalysisResult(
                success=False,
                source_file=source_file,
                module_name=module_name,
                error_message=str(e)
            )

    def _extract_procedures(self, code: str, module_name: str) -> List[VBAProcedure]:
        """Extract all procedures from VBA code."""
        procedures = []
        lines = code.split('\n')

        for i, line in enumerate(lines, 1):
            match = self.patterns['procedure'].match(line)
            if match:
                scope = match.group(1) or "Public"
                proc_type = match.group(2)
                proc_name = match.group(3)
                params = match.group(4) or ""
                return_type = match.group(5)

                procedures.append(VBAProcedure(
                    name=proc_name,
                    procedure_type=proc_type,
                    scope=scope,
                    parameters=params.strip(),
                    return_type=return_type,
                    module_name=module_name,
                    line_number=i,
                    signature=line.strip()
                ))

        return procedures

    def _extract_variables(self, code: str, module_name: str,
                           procedures: List[VBAProcedure]) -> List[VBAVariable]:
        """Extract all variables and constants from VBA code."""
        variables = []
        lines = code.split('\n')
        current_procedure = None

        for i, line in enumerate(lines, 1):
            # Track current procedure
            for proc in procedures:
                if proc.line_number == i:
                    current_procedure = proc.name
                    break

            # Check for procedure end
            if re.match(r'^\s*End\s+(Sub|Function|Property)', line, re.IGNORECASE):
                current_procedure = None
                continue

            # Extract constants with type and value
            const_match = self.patterns['const_value'].match(line)
            if const_match:
                scope = const_match.group(1) or ("Private" if current_procedure else "Public")
                const_name = const_match.group(2)
                const_type = const_match.group(3)
                const_value = const_match.group(4).strip()

                variables.append(VBAVariable(
                    name=const_name,
                    var_type=const_type,
                    declaration="Const",
                    scope=scope,
                    value=const_value,
                    module_name=module_name,
                    procedure_name=current_procedure,
                    line_number=i,
                    source=line.strip()
                ))
                continue

            # Extract simple constants
            simple_const = self.patterns['const_simple'].match(line)
            if simple_const and not const_match:
                scope = simple_const.group(1) or ("Private" if current_procedure else "Public")
                const_name = simple_const.group(2)
                const_value = simple_const.group(3).strip()

                variables.append(VBAVariable(
                    name=const_name,
                    var_type="Variant",
                    declaration="Const",
                    scope=scope,
                    value=const_value,
                    module_name=module_name,
                    procedure_name=current_procedure,
                    line_number=i,
                    source=line.strip()
                ))
                continue

            # Extract variables
            var_match = self.patterns['variable'].match(line)
            if var_match:
                declaration = var_match.group(1)
                var_names_str = var_match.group(2)
                var_type = var_match.group(3)

                # Handle multiple variables on same line
                var_names = [n.strip() for n in var_names_str.split(',')]

                # Determine scope
                if declaration in ('Private', 'Public', 'Global'):
                    scope = declaration
                elif current_procedure:
                    scope = "Local"
                else:
                    scope = "Module"

                for var_name in var_names:
                    variables.append(VBAVariable(
                        name=var_name,
                        var_type=var_type,
                        declaration=declaration,
                        scope=scope,
                        value=None,
                        module_name=module_name,
                        procedure_name=current_procedure,
                        line_number=i,
                        source=line.strip()
                    ))

        return variables

    def to_dataframe(self, results: List[VBAAnalysisResult]) -> Optional['pd.DataFrame']:
        """
        Convert analysis results to a pandas DataFrame.

        Args:
            results: List of VBAAnalysisResult objects

        Returns:
            DataFrame with all extracted elements, or None if pandas unavailable
        """
        if not PANDAS_AVAILABLE:
            return None

        data = []

        for result in results:
            # Add procedures
            for proc in result.procedures:
                data.append({
                    'Classeur': result.source_file,
                    'Module': result.module_name,
                    'Type_Module': '',
                    'Procedure': proc.name,
                    'Type_Procedure': proc.procedure_type,
                    'Scope_Procedure': proc.scope,
                    'Declaration': 'Procedure',
                    'Nom_Variable': '',
                    'Type_Variable': proc.return_type or '',
                    'Valeur': '',
                    'Ligne': proc.line_number,
                    'Code_Source': proc.signature
                })

            # Add variables
            for var in result.variables:
                data.append({
                    'Classeur': result.source_file,
                    'Module': result.module_name,
                    'Type_Module': '',
                    'Procedure': var.procedure_name or '',
                    'Type_Procedure': '',
                    'Scope_Procedure': '',
                    'Declaration': var.declaration,
                    'Nom_Variable': var.name,
                    'Type_Variable': var.var_type,
                    'Valeur': var.value or '',
                    'Ligne': var.line_number,
                    'Code_Source': var.source
                })

        return pd.DataFrame(data)

    def export_to_excel(self, results: List[VBAAnalysisResult], output_path: str) -> bool:
        """
        Export analysis results to Excel file.

        Args:
            results: List of VBAAnalysisResult objects
            output_path: Output file path

        Returns:
            True if successful
        """
        if not PANDAS_AVAILABLE:
            return False

        try:
            df = self.to_dataframe(results)
            if df is None or df.empty:
                return False

            df.to_excel(output_path, index=False, sheet_name="Analyse_VBA")
            return True

        except Exception:
            return False

    def generate_statistics(self, results: List[VBAAnalysisResult]) -> Dict[str, Any]:
        """Generate comprehensive statistics from analysis results."""
        stats = {
            'total_modules': len(results),
            'total_procedures': 0,
            'total_variables': 0,
            'total_constants': 0,
            'procedures_by_type': {},
            'procedures_by_scope': {},
            'variables_by_type': {},
            'variables_by_declaration': {},
            'procedures_per_module': {},
            'variables_per_module': {}
        }

        for result in results:
            stats['total_procedures'] += result.total_procedures
            stats['total_variables'] += result.total_variables
            stats['total_constants'] += result.total_constants
            stats['procedures_per_module'][result.module_name] = result.total_procedures
            stats['variables_per_module'][result.module_name] = result.total_variables

            for proc in result.procedures:
                proc_type = proc.procedure_type
                stats['procedures_by_type'][proc_type] = stats['procedures_by_type'].get(proc_type, 0) + 1
                stats['procedures_by_scope'][proc.scope] = stats['procedures_by_scope'].get(proc.scope, 0) + 1

            for var in result.variables:
                stats['variables_by_type'][var.var_type] = stats['variables_by_type'].get(var.var_type, 0) + 1
                stats['variables_by_declaration'][var.declaration] = stats['variables_by_declaration'].get(var.declaration, 0) + 1

        return stats

    def plot_procedures_by_module(self, results: List[VBAAnalysisResult],
                                   output_path: Optional[str] = None,
                                   figsize: Tuple[int, int] = (12, 6)) -> Optional[Any]:
        """
        Create bar chart of procedures per module.

        Args:
            results: Analysis results
            output_path: Optional path to save the figure
            figsize: Figure size (width, height)

        Returns:
            matplotlib figure or None if not available
        """
        if not MATPLOTLIB_AVAILABLE:
            return None

        stats = self.generate_statistics(results)
        proc_per_module = stats['procedures_per_module']

        if not proc_per_module:
            return None

        fig, ax = plt.subplots(figsize=figsize)

        modules = list(proc_per_module.keys())
        counts = list(proc_per_module.values())

        bars = ax.bar(modules, counts, color='skyblue', edgecolor='navy')
        ax.set_title('Nombre de procédures par module', fontsize=14, fontweight='bold')
        ax.set_xlabel('Module')
        ax.set_ylabel('Nombre de procédures')
        plt.xticks(rotation=45, ha='right')

        # Add value labels on bars
        for bar, count in zip(bars, counts):
            ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.1,
                   str(count), ha='center', va='bottom', fontsize=9)

        plt.tight_layout()

        if output_path:
            fig.savefig(output_path, dpi=150, bbox_inches='tight')

        return fig

    def plot_variables_by_type(self, results: List[VBAAnalysisResult],
                                output_path: Optional[str] = None,
                                figsize: Tuple[int, int] = (15, 6)) -> Optional[Any]:
        """
        Create charts for variables by type (bar + pie).

        Args:
            results: Analysis results
            output_path: Optional path to save the figure
            figsize: Figure size

        Returns:
            matplotlib figure or None
        """
        if not MATPLOTLIB_AVAILABLE:
            return None

        stats = self.generate_statistics(results)
        var_by_type = stats['variables_by_type']
        var_by_decl = stats['variables_by_declaration']

        if not var_by_type and not var_by_decl:
            return None

        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=figsize)

        # Bar chart - Variables by type
        if var_by_type:
            types = list(var_by_type.keys())
            counts = list(var_by_type.values())
            ax1.bar(types, counts, color='lightcoral', edgecolor='darkred')
            ax1.set_title('Distribution des types de variables', fontweight='bold')
            ax1.set_xlabel('Type de variable')
            ax1.set_ylabel('Nombre')
            ax1.tick_params(axis='x', rotation=45)

        # Pie chart - Declarations
        if var_by_decl:
            labels = list(var_by_decl.keys())
            sizes = list(var_by_decl.values())
            colors = plt.cm.Pastel1(range(len(labels)))
            ax2.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=90, colors=colors)
            ax2.set_title('Répartition des déclarations', fontweight='bold')

        plt.tight_layout()

        if output_path:
            fig.savefig(output_path, dpi=150, bbox_inches='tight')

        return fig

    def plot_declarations_distribution(self, results: List[VBAAnalysisResult],
                                        output_path: Optional[str] = None,
                                        figsize: Tuple[int, int] = (12, 6)) -> Optional[Any]:
        """
        Create stacked area chart of declarations per module.

        Args:
            results: Analysis results
            output_path: Optional path to save
            figsize: Figure size

        Returns:
            matplotlib figure or None
        """
        if not MATPLOTLIB_AVAILABLE or not PANDAS_AVAILABLE:
            return None

        df = self.to_dataframe(results)
        if df is None or df.empty:
            return None

        # Create pivot table
        pivot_data = df.pivot_table(
            index='Module',
            columns='Declaration',
            values='Nom_Variable',
            aggfunc='count',
            fill_value=0
        )

        if pivot_data.empty:
            return None

        fig, ax = plt.subplots(figsize=figsize)

        pivot_data.plot(kind='area', ax=ax, alpha=0.7)
        ax.set_title('Distribution des déclarations par module', fontweight='bold')
        ax.set_xlabel('Module')
        ax.set_ylabel('Nombre de déclarations')
        ax.legend(title='Type de déclaration', bbox_to_anchor=(1.05, 1), loc='upper left')

        plt.tight_layout()

        if output_path:
            fig.savefig(output_path, dpi=150, bbox_inches='tight')

        return fig

    def generate_report(self, results: List[VBAAnalysisResult]) -> str:
        """Generate a text report of the analysis."""
        stats = self.generate_statistics(results)

        report = []
        report.append("=" * 80)
        report.append(" RAPPORT D'ANALYSE VBA")
        report.append(f" Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        report.append("=" * 80)
        report.append("")

        # Summary
        report.append("RÉSUMÉ:")
        report.append("-" * 40)
        report.append(f"  Modules analysés: {stats['total_modules']}")
        report.append(f"  Procédures: {stats['total_procedures']}")
        report.append(f"  Variables: {stats['total_variables']}")
        report.append(f"  Constantes: {stats['total_constants']}")
        report.append("")

        # Procedures by type
        if stats['procedures_by_type']:
            report.append("PROCÉDURES PAR TYPE:")
            report.append("-" * 40)
            for proc_type, count in sorted(stats['procedures_by_type'].items()):
                report.append(f"  {proc_type}: {count}")
            report.append("")

        # Procedures by scope
        if stats['procedures_by_scope']:
            report.append("PROCÉDURES PAR PORTÉE:")
            report.append("-" * 40)
            for scope, count in sorted(stats['procedures_by_scope'].items()):
                report.append(f"  {scope}: {count}")
            report.append("")

        # Variables by type (top 10)
        if stats['variables_by_type']:
            report.append("TYPES DE VARIABLES (TOP 10):")
            report.append("-" * 40)
            sorted_types = sorted(stats['variables_by_type'].items(), key=lambda x: x[1], reverse=True)[:10]
            for var_type, count in sorted_types:
                report.append(f"  {var_type}: {count}")
            report.append("")

        # Per module details
        report.append("DÉTAIL PAR MODULE:")
        report.append("-" * 40)
        for result in results:
            report.append(f"\n  Module: {result.module_name}")
            report.append(f"    Procédures: {result.total_procedures}")
            report.append(f"    Variables: {result.total_variables}")
            report.append(f"    Constantes: {result.total_constants}")

            if result.procedures:
                report.append("    Procédures:")
                for proc in result.procedures[:10]:  # Limit to 10
                    report.append(f"      - {proc.scope} {proc.procedure_type} {proc.name}")

        report.append("")
        report.append("=" * 80)
        report.append(" FIN DU RAPPORT")
        report.append("=" * 80)

        return "\n".join(report)


def get_hex_preview(file_path: str, max_bytes: int = 512) -> str:
    """
    Generate hexadecimal preview of a binary file.

    Args:
        file_path: Path to the file
        max_bytes: Maximum bytes to read

    Returns:
        Formatted hex dump string
    """
    try:
        with open(file_path, 'rb') as f:
            data = f.read(max_bytes)

        if not data:
            return "Empty file"

        lines = []
        lines.append(f"File: {os.path.basename(file_path)}")
        lines.append(f"Size: {os.path.getsize(file_path)} bytes")
        lines.append(f"Preview: first {len(data)} bytes")
        lines.append("=" * 75)
        lines.append("")
        lines.append("Offset    00 01 02 03 04 05 06 07  08 09 0A 0B 0C 0D 0E 0F  ASCII")
        lines.append("-" * 75)

        for i in range(0, len(data), 16):
            chunk = data[i:i+16]

            # Offset
            offset = f"{i:08X}"

            # Hex part
            hex_left = ' '.join(f'{b:02X}' for b in chunk[:8])
            hex_right = ' '.join(f'{b:02X}' for b in chunk[8:16])
            hex_part = f"{hex_left:<23}  {hex_right:<23}"

            # ASCII part
            ascii_part = ''.join(
                chr(b) if 32 <= b < 127 else '.'
                for b in chunk
            )

            lines.append(f"{offset}  {hex_part}  {ascii_part}")

        if len(data) == max_bytes:
            lines.append("")
            lines.append(f"... (truncated at {max_bytes} bytes)")

        return "\n".join(lines)

    except Exception as e:
        return f"Error reading file: {e}"


def is_binary_file(file_path: str, sample_size: int = 8192) -> bool:
    """
    Check if a file is binary by examining its content.

    Args:
        file_path: Path to the file
        sample_size: Number of bytes to check

    Returns:
        True if file appears to be binary
    """
    try:
        with open(file_path, 'rb') as f:
            chunk = f.read(sample_size)

        if not chunk:
            return False

        # Check for null bytes (common in binary files)
        if b'\x00' in chunk:
            return True

        # Check ratio of non-text characters
        text_chars = bytearray({7, 8, 9, 10, 12, 13, 27} | set(range(0x20, 0x100)) - {0x7f})
        non_text = sum(1 for b in chunk if b not in text_chars)

        return non_text / len(chunk) > 0.30

    except Exception:
        return False
