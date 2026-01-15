"""
VBA Extractor Module - Extract VBA code from Office files.
Supports Excel, Word, and PowerPoint files with multiple extraction methods.
"""

import os
import sys
import tempfile
import zipfile
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from enum import Enum

# Optional imports with graceful fallback
OLETOOLS_AVAILABLE = False
WIN32COM_AVAILABLE = False

try:
    from oletools.olevba import VBA_Parser
    OLETOOLS_AVAILABLE = True
except ImportError:
    pass

try:
    import win32com.client
    WIN32COM_AVAILABLE = True
except ImportError:
    pass


class ExtractionMethod(Enum):
    """Available extraction methods."""
    AUTO = "auto"
    WIN32COM = "win32com"
    OLETOOLS = "oletools"


@dataclass
class VBAModule:
    """Represents an extracted VBA module."""
    name: str
    module_type: str
    code: str
    source_file: str
    stream_path: str = ""
    attributes: Dict[str, str] = field(default_factory=dict)

    @property
    def extension(self) -> str:
        """Get the appropriate file extension."""
        type_map = {
            "Module standard": "bas",
            "Module": "bas",
            "Classe": "cls",
            "Module de Classe": "cls",
            "UserForm": "frm",
            "Document": "cls",
        }
        return type_map.get(self.module_type, "txt")

    @property
    def line_count(self) -> int:
        """Get the number of lines in the module."""
        return len(self.code.splitlines())


@dataclass
class ExtractionResult:
    """Result of a VBA extraction operation."""
    success: bool
    source_file: str
    modules: List[VBAModule] = field(default_factory=list)
    method_used: str = ""
    error_message: str = ""
    extraction_time: float = 0.0

    @property
    def total_lines(self) -> int:
        return sum(m.line_count for m in self.modules)

    @property
    def total_modules(self) -> int:
        return len(self.modules)


class VBAExtractor:
    """
    Extract VBA code from Microsoft Office files.
    Supports multiple extraction methods with automatic fallback.
    """

    SUPPORTED_EXTENSIONS = {
        '.xlsm', '.xlsb', '.xls', '.xla', '.xlam',  # Excel
        '.docm', '.doc', '.dotm',                    # Word
        '.pptm', '.ppt', '.potm', '.ppsm',          # PowerPoint
    }

    MODULE_TYPES = {
        1: "Module standard",
        2: "Module de Classe",
        3: "UserForm",
        100: "Document"
    }

    def __init__(self, preferred_method: ExtractionMethod = ExtractionMethod.AUTO):
        self.preferred_method = preferred_method
        self._check_dependencies()

    def _check_dependencies(self) -> Dict[str, bool]:
        """Check available extraction methods."""
        return {
            "oletools": OLETOOLS_AVAILABLE,
            "win32com": WIN32COM_AVAILABLE
        }

    def can_extract(self) -> bool:
        """Check if any extraction method is available."""
        return OLETOOLS_AVAILABLE or WIN32COM_AVAILABLE

    def get_available_methods(self) -> List[str]:
        """Get list of available extraction methods."""
        methods = []
        if WIN32COM_AVAILABLE:
            methods.append("win32com")
        if OLETOOLS_AVAILABLE:
            methods.append("oletools")
        return methods

    def is_supported_file(self, file_path: str) -> bool:
        """Check if file type is supported."""
        ext = Path(file_path).suffix.lower()
        return ext in self.SUPPORTED_EXTENSIONS

    def extract(self, file_path: str, output_dir: Optional[str] = None,
                create_individual_files: bool = True,
                create_concatenated_file: bool = True) -> ExtractionResult:
        """
        Extract VBA code from an Office file.

        Args:
            file_path: Path to the Office file
            output_dir: Directory to save extracted files
            create_individual_files: Create separate files for each module
            create_concatenated_file: Create a single file with all code

        Returns:
            ExtractionResult with extracted modules
        """
        start_time = datetime.now()

        if not os.path.exists(file_path):
            return ExtractionResult(
                success=False,
                source_file=file_path,
                error_message=f"File not found: {file_path}"
            )

        if not self.is_supported_file(file_path):
            return ExtractionResult(
                success=False,
                source_file=file_path,
                error_message=f"Unsupported file type: {Path(file_path).suffix}"
            )

        # Determine extraction method
        method = self._select_method()
        if not method:
            return ExtractionResult(
                success=False,
                source_file=file_path,
                error_message="No extraction method available. Install oletools or pywin32."
            )

        # Extract VBA code
        try:
            if method == "win32com":
                modules = self._extract_with_win32com(file_path)
            else:
                modules = self._extract_with_oletools(file_path)

            if not modules:
                return ExtractionResult(
                    success=False,
                    source_file=file_path,
                    method_used=method,
                    error_message="No VBA code found in file"
                )

            # Save files if output directory specified
            if output_dir:
                self._save_modules(
                    modules, output_dir,
                    create_individual_files,
                    create_concatenated_file,
                    file_path
                )

            extraction_time = (datetime.now() - start_time).total_seconds()

            return ExtractionResult(
                success=True,
                source_file=file_path,
                modules=modules,
                method_used=method,
                extraction_time=extraction_time
            )

        except Exception as e:
            return ExtractionResult(
                success=False,
                source_file=file_path,
                method_used=method,
                error_message=str(e)
            )

    def _select_method(self) -> Optional[str]:
        """Select the best available extraction method."""
        if self.preferred_method == ExtractionMethod.WIN32COM:
            return "win32com" if WIN32COM_AVAILABLE else None
        elif self.preferred_method == ExtractionMethod.OLETOOLS:
            return "oletools" if OLETOOLS_AVAILABLE else None
        else:  # AUTO
            if sys.platform == "win32" and WIN32COM_AVAILABLE:
                return "win32com"
            elif OLETOOLS_AVAILABLE:
                return "oletools"
            elif WIN32COM_AVAILABLE:
                return "win32com"
        return None

    def _extract_with_win32com(self, file_path: str) -> List[VBAModule]:
        """Extract VBA using Win32COM automation. Supports Excel, Word, and PowerPoint."""
        ext = Path(file_path).suffix.lower()

        if ext in {'.xlsm', '.xlsb', '.xls', '.xla', '.xlam'}:
            return self._extract_excel_win32com(file_path)
        elif ext in {'.docm', '.doc', '.dotm'}:
            return self._extract_word_win32com(file_path)
        elif ext in {'.pptm', '.ppt', '.potm', '.ppsm'}:
            return self._extract_powerpoint_win32com(file_path)
        else:
            return []

    def _extract_excel_win32com(self, file_path: str) -> List[VBAModule]:
        """Extract VBA from Excel files using Win32COM."""
        modules = []
        excel = None

        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            workbook = excel.Workbooks.Open(os.path.abspath(file_path))

            try:
                vb_project = workbook.VBProject

                for component in vb_project.VBComponents:
                    if component.CodeModule.CountOfLines > 0:
                        code = component.CodeModule.Lines(1, component.CodeModule.CountOfLines)
                        module_type = self.MODULE_TYPES.get(component.Type, "Unknown")

                        modules.append(VBAModule(
                            name=component.Name,
                            module_type=module_type,
                            code=code,
                            source_file=file_path
                        ))

            finally:
                workbook.Close(SaveChanges=False)

        finally:
            if excel:
                excel.Quit()

        return modules

    def _extract_word_win32com(self, file_path: str) -> List[VBAModule]:
        """Extract VBA from Word files using Win32COM."""
        modules = []
        word = None

        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0  # wdAlertsNone

            doc = word.Documents.Open(os.path.abspath(file_path), ReadOnly=True)

            try:
                vb_project = doc.VBProject

                for component in vb_project.VBComponents:
                    if component.CodeModule.CountOfLines > 0:
                        code = component.CodeModule.Lines(1, component.CodeModule.CountOfLines)
                        module_type = self.MODULE_TYPES.get(component.Type, "Unknown")

                        modules.append(VBAModule(
                            name=component.Name,
                            module_type=module_type,
                            code=code,
                            source_file=file_path
                        ))

            finally:
                doc.Close(SaveChanges=False)

        except Exception as e:
            # Word may not have VBProject access enabled
            raise Exception(f"Word VBA extraction failed: {e}. Enable 'Trust access to VBA project' in Word options.")

        finally:
            if word:
                word.Quit()

        return modules

    def _extract_powerpoint_win32com(self, file_path: str) -> List[VBAModule]:
        """Extract VBA from PowerPoint files using Win32COM."""
        modules = []
        ppt = None

        try:
            ppt = win32com.client.Dispatch("PowerPoint.Application")
            # PowerPoint doesn't support Visible=False in some versions
            ppt.DisplayAlerts = 0  # ppAlertsNone

            presentation = ppt.Presentations.Open(
                os.path.abspath(file_path),
                ReadOnly=True,
                WithWindow=False
            )

            try:
                vb_project = presentation.VBProject

                for component in vb_project.VBComponents:
                    if component.CodeModule.CountOfLines > 0:
                        code = component.CodeModule.Lines(1, component.CodeModule.CountOfLines)
                        module_type = self.MODULE_TYPES.get(component.Type, "Unknown")

                        modules.append(VBAModule(
                            name=component.Name,
                            module_type=module_type,
                            code=code,
                            source_file=file_path
                        ))

            finally:
                presentation.Close()

        except Exception as e:
            raise Exception(f"PowerPoint VBA extraction failed: {e}. Enable 'Trust access to VBA project' in PowerPoint options.")

        finally:
            if ppt:
                ppt.Quit()

        return modules

    def _extract_with_oletools(self, file_path: str) -> List[VBAModule]:
        """Extract VBA using oletools library."""
        modules = []

        vba_parser = VBA_Parser(file_path)
        try:
            if vba_parser.detect_vba_macros():
                for (filename, stream_path, vba_filename, vba_code) in vba_parser.extract_macros():
                    # Determine module type from code
                    if "Attribute VB_PredeclaredId" in vba_code:
                        module_type = "Classe/UserForm"
                    elif vba_code.strip().startswith("Attribute VB_"):
                        module_type = "Module"
                    else:
                        module_type = "Code"

                    # Clean filename
                    safe_name = vba_filename.replace('/', '_').replace('\\', '_')
                    if not safe_name:
                        safe_name = f"module_{len(modules) + 1}"

                    modules.append(VBAModule(
                        name=safe_name,
                        module_type=module_type,
                        code=vba_code,
                        source_file=file_path,
                        stream_path=stream_path
                    ))
        finally:
            vba_parser.close()

        return modules

    def _save_modules(self, modules: List[VBAModule], output_dir: str,
                      create_individual: bool, create_concatenated: bool,
                      source_file: str) -> None:
        """Save extracted modules to files."""
        os.makedirs(output_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Individual files
        if create_individual:
            for module in modules:
                filename = f"{module.name}.{module.extension}"
                filepath = os.path.join(output_dir, filename)

                with open(filepath, 'w', encoding='utf-8') as f:
                    f.write(f"' Module: {module.name}\n")
                    f.write(f"' Type: {module.module_type}\n")
                    f.write(f"' Source: {os.path.basename(source_file)}\n")
                    f.write(f"' Extracted: {timestamp}\n")
                    f.write("' " + "=" * 60 + "\n\n")
                    f.write(module.code)

        # Concatenated file
        if create_concatenated:
            base_name = Path(source_file).stem
            concat_file = os.path.join(output_dir, f"{base_name}_all_vba.txt")

            with open(concat_file, 'w', encoding='utf-8') as f:
                f.write("=" * 80 + "\n")
                f.write(f" VBA CODE EXTRACTED FROM: {os.path.basename(source_file)}\n")
                f.write(f" Extraction Date: {timestamp}\n")
                f.write(f" Total Modules: {len(modules)}\n")
                f.write("=" * 80 + "\n\n")

                # Table of contents
                f.write("TABLE OF CONTENTS:\n")
                f.write("-" * 40 + "\n")
                for i, module in enumerate(modules, 1):
                    f.write(f"{i:3d}. {module.name} ({module.module_type})\n")
                f.write("\n" + "=" * 80 + "\n\n")

                # Module contents
                for i, module in enumerate(modules, 1):
                    f.write("\n" + "#" * 80 + "\n")
                    f.write(f"# MODULE {i}: {module.name}\n")
                    f.write(f"# Type: {module.module_type}\n")
                    f.write(f"# Lines: {module.line_count}\n")
                    f.write("#" * 80 + "\n\n")
                    f.write(module.code)
                    f.write("\n\n" + "-" * 80 + "\n")

                f.write("\n" + "=" * 80 + "\n")
                f.write(" END OF FILE\n")
                f.write("=" * 80 + "\n")
