"""
Export Manager - Universal export system supporting multiple formats.
Supports CSV, JSON, HTML, TXT, and PDF export with templates.
"""

import csv
import json
import os
import zipfile
from abc import ABC, abstractmethod
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Union
import html as html_module


@dataclass
class ExportResult:
    """Result of an export operation."""
    success: bool
    file_path: str
    format: str
    size: int = 0
    message: str = ""
    duration: float = 0.0


class BaseExporter(ABC):
    """Abstract base class for exporters."""

    @abstractmethod
    def export(self, data: Any, file_path: str, **options) -> ExportResult:
        """Export data to file."""
        pass

    @property
    @abstractmethod
    def format_name(self) -> str:
        """Return the format name."""
        pass

    @property
    @abstractmethod
    def file_extension(self) -> str:
        """Return the file extension."""
        pass


class JSONExporter(BaseExporter):
    """Export data to JSON format."""

    @property
    def format_name(self) -> str:
        return "JSON"

    @property
    def file_extension(self) -> str:
        return ".json"

    def export(self, data: Any, file_path: str, **options) -> ExportResult:
        start = datetime.now()
        try:
            indent = options.get('indent', 2)
            ensure_ascii = options.get('ensure_ascii', False)

            # Convert to serializable format
            serializable = self._make_serializable(data)

            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(serializable, f, indent=indent, ensure_ascii=ensure_ascii)

            size = os.path.getsize(file_path)
            duration = (datetime.now() - start).total_seconds()

            return ExportResult(True, file_path, "JSON", size, "Export successful", duration)

        except Exception as e:
            return ExportResult(False, file_path, "JSON", message=str(e))

    def _make_serializable(self, obj: Any) -> Any:
        """Convert object to JSON-serializable format."""
        if isinstance(obj, datetime):
            return obj.isoformat()
        elif isinstance(obj, Path):
            return str(obj)
        elif isinstance(obj, set):
            return list(obj)
        elif isinstance(obj, dict):
            return {k: self._make_serializable(v) for k, v in obj.items()}
        elif isinstance(obj, (list, tuple)):
            return [self._make_serializable(item) for item in obj]
        elif hasattr(obj, '__dict__'):
            return self._make_serializable(obj.__dict__)
        elif hasattr(obj, 'to_dict'):
            return self._make_serializable(obj.to_dict())
        return obj


class CSVExporter(BaseExporter):
    """Export data to CSV format."""

    @property
    def format_name(self) -> str:
        return "CSV"

    @property
    def file_extension(self) -> str:
        return ".csv"

    def export(self, data: Any, file_path: str, **options) -> ExportResult:
        start = datetime.now()
        try:
            delimiter = options.get('delimiter', ',')
            encoding = options.get('encoding', 'utf-8-sig')  # BOM for Excel

            # Convert data to list of dicts
            rows = self._to_rows(data)

            if not rows:
                return ExportResult(False, file_path, "CSV", message="No data to export")

            # Get all unique keys for headers
            headers = []
            for row in rows:
                for key in row.keys():
                    if key not in headers:
                        headers.append(key)

            with open(file_path, 'w', encoding=encoding, newline='') as f:
                writer = csv.DictWriter(f, fieldnames=headers, delimiter=delimiter)
                writer.writeheader()
                writer.writerows(rows)

            size = os.path.getsize(file_path)
            duration = (datetime.now() - start).total_seconds()

            return ExportResult(True, file_path, "CSV", size, f"Exported {len(rows)} rows", duration)

        except Exception as e:
            return ExportResult(False, file_path, "CSV", message=str(e))

    def _to_rows(self, data: Any) -> List[Dict]:
        """Convert data to list of dictionaries."""
        if isinstance(data, list):
            rows = []
            for item in data:
                if isinstance(item, dict):
                    rows.append(item)
                elif hasattr(item, '__dict__'):
                    rows.append(self._flatten_dict(item.__dict__))
                elif hasattr(item, 'to_dict'):
                    rows.append(self._flatten_dict(item.to_dict()))
            return rows
        elif isinstance(data, dict):
            # Check for nested list data with common keys
            if 'rows' in data and isinstance(data['rows'], list):
                return self._to_rows(data['rows'])
            elif 'data' in data and isinstance(data['data'], list):
                return self._to_rows(data['data'])
            elif 'files' in data and isinstance(data['files'], list):
                return self._to_rows(data['files'])
            # For simple dicts (stats, summary), return as single row
            return [self._flatten_dict(data)]
        return []

    def _flatten_dict(self, d: Dict, parent_key: str = '', sep: str = '_') -> Dict:
        """Flatten nested dictionaries."""
        items = []
        for k, v in d.items():
            new_key = f"{parent_key}{sep}{k}" if parent_key else k
            if isinstance(v, dict):
                items.extend(self._flatten_dict(v, new_key, sep).items())
            elif isinstance(v, (list, set)):
                items.append((new_key, ', '.join(str(x) for x in v)))
            elif isinstance(v, datetime):
                items.append((new_key, v.isoformat()))
            else:
                items.append((new_key, v))
        return dict(items)


class TXTExporter(BaseExporter):
    """Export data to plain text format."""

    @property
    def format_name(self) -> str:
        return "TXT"

    @property
    def file_extension(self) -> str:
        return ".txt"

    def export(self, data: Any, file_path: str, **options) -> ExportResult:
        start = datetime.now()
        try:
            title = options.get('title', 'Export Report')
            encoding = options.get('encoding', 'utf-8')

            content = self._format_text(data, title)

            with open(file_path, 'w', encoding=encoding) as f:
                f.write(content)

            size = os.path.getsize(file_path)
            duration = (datetime.now() - start).total_seconds()

            return ExportResult(True, file_path, "TXT", size, "Export successful", duration)

        except Exception as e:
            return ExportResult(False, file_path, "TXT", message=str(e))

    def _format_text(self, data: Any, title: str) -> str:
        """Format data as text."""
        lines = [
            "=" * 80,
            f" {title}",
            "=" * 80,
            f" Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            "=" * 80,
            ""
        ]

        if isinstance(data, dict):
            # Statistics section
            if 'statistics' in data or 'summary' in data:
                stats = data.get('statistics', data.get('summary', {}))
                lines.append("STATISTICS")
                lines.append("-" * 40)
                for key, value in stats.items():
                    if isinstance(value, (int, float, str)):
                        label = key.replace('_', ' ').title()
                        lines.append(f"  {label}: {value}")
                lines.append("")

            # Files section
            if 'files' in data:
                lines.append("FILES")
                lines.append("-" * 40)
                for f in data['files'][:100]:
                    if isinstance(f, dict):
                        name = f.get('name', f.get('path', 'Unknown'))
                        lines.append(f"  - {name}")
                    else:
                        lines.append(f"  - {f}")
                lines.append("")

            # Modules section (for VBA)
            if 'modules' in data:
                lines.append("MODULES")
                lines.append("-" * 40)
                for m in data['modules']:
                    if isinstance(m, dict):
                        name = m.get('name', 'Unknown')
                        mtype = m.get('type', m.get('module_type', ''))
                        lines.append(f"  - {name} ({mtype})")
                    else:
                        lines.append(f"  - {m}")
                lines.append("")

        elif isinstance(data, list):
            lines.append("DATA")
            lines.append("-" * 40)
            for item in data[:100]:
                lines.append(f"  - {item}")
            lines.append("")

        elif isinstance(data, str):
            lines.append(data)

        lines.extend([
            "",
            "=" * 80,
            " End of Report",
            "=" * 80
        ])

        return '\n'.join(lines)


class HTMLExporter(BaseExporter):
    """Export data to HTML format with modern styling."""

    @property
    def format_name(self) -> str:
        return "HTML"

    @property
    def file_extension(self) -> str:
        return ".html"

    def export(self, data: Any, file_path: str, **options) -> ExportResult:
        start = datetime.now()
        try:
            title = options.get('title', 'Export Report')
            theme = options.get('theme', 'dark')

            content = self._generate_html(data, title, theme)

            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(content)

            size = os.path.getsize(file_path)
            duration = (datetime.now() - start).total_seconds()

            return ExportResult(True, file_path, "HTML", size, "Export successful", duration)

        except Exception as e:
            return ExportResult(False, file_path, "HTML", message=str(e))

    def _generate_html(self, data: Any, title: str, theme: str) -> str:
        """Generate HTML document."""
        # CSS styles
        if theme == 'dark':
            bg_color = '#0f172a'
            bg_secondary = '#1e293b'
            text_color = '#f1f5f9'
            text_secondary = '#94a3b8'
            border_color = '#334155'
            primary_color = '#3b82f6'
        else:
            bg_color = '#ffffff'
            bg_secondary = '#f8fafc'
            text_color = '#1e293b'
            text_secondary = '#64748b'
            border_color = '#e2e8f0'
            primary_color = '#2563eb'

        css = f'''
        :root {{
            --bg: {bg_color};
            --bg-secondary: {bg_secondary};
            --text: {text_color};
            --text-secondary: {text_secondary};
            --border: {border_color};
            --primary: {primary_color};
            --success: #10b981;
            --warning: #f59e0b;
            --danger: #ef4444;
        }}
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            background: var(--bg);
            color: var(--text);
            line-height: 1.6;
            padding: 2rem;
        }}
        .container {{ max-width: 1200px; margin: 0 auto; }}
        h1 {{
            font-size: 2rem;
            margin-bottom: 0.5rem;
            color: var(--primary);
        }}
        .meta {{ color: var(--text-secondary); font-size: 0.875rem; margin-bottom: 2rem; }}
        .section {{
            background: var(--bg-secondary);
            border: 1px solid var(--border);
            border-radius: 8px;
            padding: 1.5rem;
            margin-bottom: 1.5rem;
        }}
        .section h2 {{
            font-size: 1.25rem;
            margin-bottom: 1rem;
            padding-bottom: 0.5rem;
            border-bottom: 1px solid var(--border);
        }}
        .stats-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
            gap: 1rem;
        }}
        .stat-card {{
            background: var(--bg);
            border: 1px solid var(--border);
            border-radius: 6px;
            padding: 1rem;
            text-align: center;
        }}
        .stat-value {{
            font-size: 1.75rem;
            font-weight: 700;
            color: var(--primary);
        }}
        .stat-label {{
            font-size: 0.75rem;
            color: var(--text-secondary);
            text-transform: uppercase;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
        }}
        th, td {{
            padding: 0.75rem;
            text-align: left;
            border-bottom: 1px solid var(--border);
        }}
        th {{
            font-weight: 600;
            color: var(--text-secondary);
            font-size: 0.75rem;
            text-transform: uppercase;
        }}
        tr:hover {{ background: var(--bg); }}
        .badge {{
            display: inline-block;
            padding: 0.25rem 0.5rem;
            border-radius: 4px;
            font-size: 0.75rem;
            font-weight: 600;
        }}
        .badge-success {{ background: var(--success); color: white; }}
        .badge-warning {{ background: var(--warning); color: black; }}
        .footer {{
            text-align: center;
            margin-top: 2rem;
            padding-top: 1rem;
            border-top: 1px solid var(--border);
            color: var(--text-secondary);
            font-size: 0.875rem;
        }}
        '''

        # Build HTML content
        stats_html = ""
        files_html = ""
        modules_html = ""

        if isinstance(data, dict):
            # Statistics
            stats = data.get('statistics', data.get('summary', {}))
            if stats:
                stat_cards = []
                for key, value in stats.items():
                    if isinstance(value, (int, float)):
                        label = key.replace('_', ' ').title()
                        stat_cards.append(f'''
                        <div class="stat-card">
                            <div class="stat-value">{value:,}</div>
                            <div class="stat-label">{html_module.escape(label)}</div>
                        </div>
                        ''')
                if stat_cards:
                    stats_html = f'''
                    <div class="section">
                        <h2>Statistics</h2>
                        <div class="stats-grid">
                            {''.join(stat_cards)}
                        </div>
                    </div>
                    '''

            # Files table
            files = data.get('files', [])
            if files:
                rows = []
                for f in files[:100]:
                    if isinstance(f, dict):
                        name = html_module.escape(str(f.get('name', f.get('path', 'Unknown'))))
                        size = f.get('size', 0)
                        lines = f.get('line_count', f.get('lines', '-'))
                        rows.append(f'<tr><td>{name}</td><td>{lines}</td><td>{self._format_size(size)}</td></tr>')
                if rows:
                    files_html = f'''
                    <div class="section">
                        <h2>Files ({len(files)} total)</h2>
                        <table>
                            <thead><tr><th>Name</th><th>Lines</th><th>Size</th></tr></thead>
                            <tbody>{''.join(rows)}</tbody>
                        </table>
                    </div>
                    '''

            # Modules table (for VBA)
            modules = data.get('modules', [])
            if modules:
                rows = []
                for m in modules:
                    if isinstance(m, dict):
                        name = html_module.escape(str(m.get('name', 'Unknown')))
                        mtype = html_module.escape(str(m.get('type', m.get('module_type', '-'))))
                        lines = m.get('line_count', m.get('lines', '-'))
                        rows.append(f'<tr><td>{name}</td><td>{mtype}</td><td>{lines}</td></tr>')
                if rows:
                    modules_html = f'''
                    <div class="section">
                        <h2>VBA Modules ({len(modules)} total)</h2>
                        <table>
                            <thead><tr><th>Name</th><th>Type</th><th>Lines</th></tr></thead>
                            <tbody>{''.join(rows)}</tbody>
                        </table>
                    </div>
                    '''

        return f'''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{html_module.escape(title)}</title>
    <style>{css}</style>
</head>
<body>
    <div class="container">
        <h1>{html_module.escape(title)}</h1>
        <p class="meta">Generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} by CodeExtractPro</p>

        {stats_html}
        {modules_html}
        {files_html}

        <div class="footer">
            <p>Generated by CodeExtractPro v1.0</p>
        </div>
    </div>
</body>
</html>'''

    def _format_size(self, size: int) -> str:
        """Format size to human-readable."""
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size < 1024:
                return f"{size:.1f} {unit}" if unit != 'B' else f"{size} {unit}"
            size /= 1024
        return f"{size:.1f} TB"


class ExportManager:
    """
    Universal export manager supporting multiple formats.
    """

    def __init__(self):
        self.exporters: Dict[str, BaseExporter] = {
            'json': JSONExporter(),
            'csv': CSVExporter(),
            'txt': TXTExporter(),
            'html': HTMLExporter(),
        }

    def get_available_formats(self) -> List[str]:
        """Get list of available export formats."""
        return list(self.exporters.keys())

    def export(self, data: Any, file_path: str, format: Optional[str] = None, **options) -> ExportResult:
        """
        Export data to specified format.

        Args:
            data: Data to export
            file_path: Output file path
            format: Export format (auto-detected from extension if not provided)
            **options: Format-specific options

        Returns:
            ExportResult with status and details
        """
        # Auto-detect format from extension
        if format is None:
            ext = Path(file_path).suffix.lower()
            format = ext.lstrip('.')

        if format not in self.exporters:
            return ExportResult(
                False, file_path, format,
                message=f"Unsupported format: {format}. Available: {', '.join(self.exporters.keys())}"
            )

        # Ensure directory exists
        Path(file_path).parent.mkdir(parents=True, exist_ok=True)

        return self.exporters[format].export(data, file_path, **options)

    def export_multiple(self, data: Any, base_path: str, formats: List[str], **options) -> Dict[str, ExportResult]:
        """Export data to multiple formats."""
        results = {}
        base = Path(base_path)

        for fmt in formats:
            if fmt in self.exporters:
                file_path = str(base.with_suffix(self.exporters[fmt].file_extension))
                results[fmt] = self.export(data, file_path, fmt, **options)

        return results

    def create_archive(self, files: List[str], archive_path: str) -> ExportResult:
        """Create a ZIP archive of exported files."""
        start = datetime.now()
        try:
            with zipfile.ZipFile(archive_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                for file_path in files:
                    if os.path.exists(file_path):
                        zf.write(file_path, os.path.basename(file_path))

            size = os.path.getsize(archive_path)
            duration = (datetime.now() - start).total_seconds()

            return ExportResult(True, archive_path, "ZIP", size, f"Archived {len(files)} files", duration)

        except Exception as e:
            return ExportResult(False, archive_path, "ZIP", message=str(e))


# Global export manager instance
_export_manager: Optional[ExportManager] = None


def get_export_manager() -> ExportManager:
    """Get the global export manager instance."""
    global _export_manager
    if _export_manager is None:
        _export_manager = ExportManager()
    return _export_manager
