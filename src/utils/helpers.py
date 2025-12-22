"""
Helper Functions - Common utilities used across the application.
"""

import os
import re
import hashlib
from pathlib import Path
from typing import Optional, List, Tuple
from datetime import datetime


def safe_path(path: str) -> Path:
    """Convert string to Path and ensure it's safe."""
    return Path(path).resolve()


def detect_encoding(file_path: str, sample_size: int = 8192) -> str:
    """
    Detect the encoding of a file.
    Returns the most likely encoding.
    """
    encodings = ['utf-8', 'utf-8-sig', 'latin-1', 'cp1252', 'iso-8859-1']

    # Try to read with BOM detection
    with open(file_path, 'rb') as f:
        raw = f.read(sample_size)

    # Check for BOM
    if raw.startswith(b'\xef\xbb\xbf'):
        return 'utf-8-sig'
    if raw.startswith(b'\xff\xfe'):
        return 'utf-16-le'
    if raw.startswith(b'\xfe\xff'):
        return 'utf-16-be'

    # Try each encoding
    for encoding in encodings:
        try:
            raw.decode(encoding)
            return encoding
        except (UnicodeDecodeError, LookupError):
            continue

    return 'latin-1'  # Fallback


def format_size(size_bytes: int) -> str:
    """Format byte size to human-readable string."""
    for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
        if size_bytes < 1024:
            return f"{size_bytes:.2f} {unit}" if unit != 'B' else f"{size_bytes} {unit}"
        size_bytes /= 1024
    return f"{size_bytes:.2f} PB"


def sanitize_filename(filename: str, replacement: str = "_") -> str:
    """
    Sanitize a filename by removing invalid characters.
    """
    # Remove invalid characters
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        filename = filename.replace(char, replacement)

    # Remove control characters
    filename = ''.join(c for c in filename if ord(c) >= 32)

    # Trim and limit length
    filename = filename.strip().strip('.')
    if len(filename) > 255:
        name, ext = os.path.splitext(filename)
        filename = name[:255 - len(ext)] + ext

    return filename or "unnamed"


def generate_timestamp() -> str:
    """Generate a timestamp string for file naming."""
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def calculate_file_hash(file_path: str, algorithm: str = 'md5') -> str:
    """Calculate hash of a file."""
    hash_func = hashlib.new(algorithm)
    with open(file_path, 'rb') as f:
        for chunk in iter(lambda: f.read(8192), b''):
            hash_func.update(chunk)
    return hash_func.hexdigest()


def read_file_safe(file_path: str, max_size: int = 10 * 1024 * 1024) -> Tuple[Optional[str], str]:
    """
    Safely read a file with encoding detection.
    Returns (content, encoding) or (None, error_message).
    """
    try:
        size = os.path.getsize(file_path)
        if size > max_size:
            return None, f"File too large: {format_size(size)}"

        encoding = detect_encoding(file_path)
        with open(file_path, 'r', encoding=encoding, errors='replace') as f:
            return f.read(), encoding

    except Exception as e:
        return None, str(e)


def find_files(
    directory: str,
    patterns: List[str],
    exclude_dirs: Optional[List[str]] = None,
    exclude_patterns: Optional[List[str]] = None,
    max_depth: Optional[int] = None
) -> List[str]:
    """
    Find files matching patterns in a directory.

    Args:
        directory: Root directory to search
        patterns: List of glob patterns (e.g., ['*.py', '*.txt'])
        exclude_dirs: Directory names to exclude
        exclude_patterns: File patterns to exclude
        max_depth: Maximum recursion depth

    Returns:
        List of matching file paths
    """
    from fnmatch import fnmatch

    exclude_dirs = exclude_dirs or ['__pycache__', '.git', 'node_modules', '.venv', 'venv']
    exclude_patterns = exclude_patterns or []
    files = []

    def should_include(filename: str) -> bool:
        for pattern in patterns:
            if fnmatch(filename, pattern):
                for exc_pattern in exclude_patterns:
                    if fnmatch(filename, exc_pattern):
                        return False
                return True
        return False

    def walk_dir(path: str, depth: int = 0):
        if max_depth is not None and depth > max_depth:
            return

        try:
            entries = os.listdir(path)
        except PermissionError:
            return

        for entry in entries:
            full_path = os.path.join(path, entry)

            if os.path.isdir(full_path):
                if entry not in exclude_dirs:
                    walk_dir(full_path, depth + 1)
            elif os.path.isfile(full_path):
                if should_include(entry):
                    files.append(full_path)

    walk_dir(directory)
    return sorted(files)


def create_directory_tree(directory: str, exclude_dirs: Optional[List[str]] = None) -> str:
    """
    Create a text representation of a directory tree.
    """
    exclude_dirs = exclude_dirs or ['__pycache__', '.git', 'node_modules']
    lines = []

    def add_entry(path: str, prefix: str = "", is_last: bool = True):
        name = os.path.basename(path)
        connector = "└── " if is_last else "├── "
        lines.append(f"{prefix}{connector}{name}")

        if os.path.isdir(path):
            try:
                entries = sorted(os.listdir(path))
                entries = [e for e in entries if e not in exclude_dirs]
                dirs = [e for e in entries if os.path.isdir(os.path.join(path, e))]
                files = [e for e in entries if os.path.isfile(os.path.join(path, e))]
                all_entries = dirs + files

                new_prefix = prefix + ("    " if is_last else "│   ")
                for i, entry in enumerate(all_entries):
                    entry_path = os.path.join(path, entry)
                    add_entry(entry_path, new_prefix, i == len(all_entries) - 1)
            except PermissionError:
                pass

    lines.append(os.path.basename(directory) or directory)
    try:
        entries = sorted(os.listdir(directory))
        entries = [e for e in entries if e not in exclude_dirs]
        for i, entry in enumerate(entries):
            entry_path = os.path.join(directory, entry)
            add_entry(entry_path, "", i == len(entries) - 1)
    except PermissionError:
        pass

    return "\n".join(lines)


def merge_dicts(*dicts) -> dict:
    """Deep merge multiple dictionaries."""
    result = {}
    for d in dicts:
        for key, value in d.items():
            if key in result and isinstance(result[key], dict) and isinstance(value, dict):
                result[key] = merge_dicts(result[key], value)
            else:
                result[key] = value
    return result


def truncate_string(s: str, max_length: int = 100, suffix: str = "...") -> str:
    """Truncate a string to a maximum length."""
    if len(s) <= max_length:
        return s
    return s[:max_length - len(suffix)] + suffix
