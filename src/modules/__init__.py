"""
CodeExtractPro - Modules Package
Contains all processing modules for code extraction and analysis.
"""

from .vba_extractor import VBAExtractor
from .python_analyzer import PythonAnalyzer
from .folder_scanner import FolderScanner
from .vba_optimizer import VBAOptimizer
from .report_generator import ReportGenerator
from .vba_analyzer import VBAAnalyzer, get_hex_preview, is_binary_file

__all__ = [
    'VBAExtractor',
    'PythonAnalyzer',
    'FolderScanner',
    'VBAOptimizer',
    'ReportGenerator',
    'VBAAnalyzer',
    'get_hex_preview',
    'is_binary_file'
]
