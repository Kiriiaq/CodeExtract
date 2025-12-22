"""
Report Generator Module - Generate comprehensive reports in multiple formats.
Supports HTML, Markdown, JSON, and plain text output.
"""

import json
import os
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional
import html


class ReportGenerator:
    """
    Generate professional reports in multiple formats.
    Supports HTML, Markdown, JSON, and text output.
    """

    def __init__(self, title: str = "Code Analysis Report"):
        self.title = title
        self.generated_at = datetime.now()

    def generate_html(self, data: Dict[str, Any], output_path: str,
                      template: str = "modern") -> None:
        """Generate an HTML report with modern styling."""
        html_content = self._build_html(data, template)

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_content)

    def generate_markdown(self, data: Dict[str, Any], output_path: str) -> None:
        """Generate a Markdown report."""
        md_content = self._build_markdown(data)

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(md_content)

    def generate_json(self, data: Dict[str, Any], output_path: str,
                      indent: int = 2) -> None:
        """Generate a JSON report."""
        # Convert datetime objects to strings
        serializable = self._make_serializable(data)

        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(serializable, f, indent=indent, ensure_ascii=False)

    def generate_text(self, data: Dict[str, Any], output_path: str) -> None:
        """Generate a plain text report."""
        text_content = self._build_text(data)

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(text_content)

    def _make_serializable(self, obj: Any) -> Any:
        """Convert an object to be JSON serializable."""
        if isinstance(obj, datetime):
            return obj.isoformat()
        elif isinstance(obj, Path):
            return str(obj)
        elif isinstance(obj, set):
            return list(obj)
        elif isinstance(obj, dict):
            return {k: self._make_serializable(v) for k, v in obj.items()}
        elif isinstance(obj, list):
            return [self._make_serializable(item) for item in obj]
        elif hasattr(obj, '__dict__'):
            return self._make_serializable(obj.__dict__)
        return obj

    def _build_html(self, data: Dict[str, Any], template: str) -> str:
        """Build HTML content."""
        stats = data.get('statistics', {})
        files = data.get('files', [])

        return f'''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{html.escape(self.title)}</title>
    <style>
        :root {{
            --primary: #3b82f6;
            --primary-dark: #2563eb;
            --bg: #0f172a;
            --bg-secondary: #1e293b;
            --text: #f1f5f9;
            --text-secondary: #94a3b8;
            --border: #334155;
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
        }}
        .container {{ max-width: 1400px; margin: 0 auto; padding: 2rem; }}
        .header {{ text-align: center; margin-bottom: 3rem; }}
        h1 {{
            font-size: 2.5rem;
            background: linear-gradient(135deg, var(--primary) 0%, #06b6d4 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            margin-bottom: 0.5rem;
        }}
        .date {{ color: var(--text-secondary); font-size: 0.875rem; }}
        .stats-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 1.5rem;
            margin: 2rem 0;
        }}
        .stat-card {{
            background: var(--bg-secondary);
            border: 1px solid var(--border);
            border-radius: 12px;
            padding: 1.5rem;
            transition: transform 0.2s;
        }}
        .stat-card:hover {{ transform: translateY(-2px); }}
        .stat-value {{
            font-size: 2rem;
            font-weight: 700;
            color: var(--primary);
        }}
        .stat-label {{
            font-size: 0.875rem;
            color: var(--text-secondary);
        }}
        .section {{
            background: var(--bg-secondary);
            border: 1px solid var(--border);
            border-radius: 12px;
            padding: 1.5rem;
            margin: 2rem 0;
        }}
        .section h2 {{
            color: var(--primary);
            margin-bottom: 1rem;
            padding-bottom: 0.5rem;
            border-bottom: 1px solid var(--border);
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            margin: 1rem 0;
        }}
        th, td {{
            padding: 0.75rem;
            text-align: left;
            border-bottom: 1px solid var(--border);
        }}
        th {{ color: var(--text-secondary); font-weight: 600; }}
        tr:hover {{ background: var(--bg); }}
        .badge {{
            display: inline-block;
            padding: 0.25rem 0.75rem;
            border-radius: 20px;
            font-size: 0.75rem;
            font-weight: 600;
        }}
        .badge-success {{ background: var(--success); color: white; }}
        .badge-warning {{ background: var(--warning); color: black; }}
        .badge-danger {{ background: var(--danger); color: white; }}
        .footer {{
            text-align: center;
            margin-top: 3rem;
            padding-top: 2rem;
            border-top: 1px solid var(--border);
            color: var(--text-secondary);
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>{html.escape(self.title)}</h1>
            <p class="date">Generated: {self.generated_at.strftime('%Y-%m-%d %H:%M:%S')}</p>
        </div>

        <div class="stats-grid">
            {self._build_stat_cards(stats)}
        </div>

        {self._build_files_section(files)}

        <div class="footer">
            <p>Generated by CodeExtractPro</p>
        </div>
    </div>
</body>
</html>'''

    def _build_stat_cards(self, stats: Dict[str, Any]) -> str:
        """Build HTML stat cards."""
        cards = []
        stat_labels = {
            'total_files': 'Total Files',
            'total_lines': 'Total Lines',
            'total_code_lines': 'Code Lines',
            'total_classes': 'Classes',
            'total_functions': 'Functions',
            'total_modules': 'VBA Modules'
        }

        for key, label in stat_labels.items():
            value = stats.get(key, 0)
            if value:
                cards.append(f'''
                <div class="stat-card">
                    <div class="stat-value">{value:,}</div>
                    <div class="stat-label">{label}</div>
                </div>''')

        return ''.join(cards)

    def _build_files_section(self, files: List[Dict]) -> str:
        """Build HTML files section."""
        if not files:
            return ''

        rows = []
        for f in files[:100]:  # Limit to first 100 files
            name = html.escape(f.get('name', 'Unknown'))
            size = f.get('size', 0)
            lines = f.get('line_count', 0)
            rows.append(f'''
            <tr>
                <td>{name}</td>
                <td>{lines:,}</td>
                <td>{self._format_size(size)}</td>
            </tr>''')

        return f'''
        <div class="section">
            <h2>Files Analyzed</h2>
            <table>
                <thead>
                    <tr>
                        <th>File</th>
                        <th>Lines</th>
                        <th>Size</th>
                    </tr>
                </thead>
                <tbody>
                    {''.join(rows)}
                </tbody>
            </table>
        </div>'''

    def _build_markdown(self, data: Dict[str, Any]) -> str:
        """Build Markdown content."""
        stats = data.get('statistics', {})
        files = data.get('files', [])

        lines = [
            f"# {self.title}",
            "",
            f"**Generated:** {self.generated_at.strftime('%Y-%m-%d %H:%M:%S')}",
            "",
            "## Summary Statistics",
            "",
            "| Metric | Value |",
            "|--------|-------|"
        ]

        for key, value in stats.items():
            if isinstance(value, (int, float)):
                lines.append(f"| {key.replace('_', ' ').title()} | {value:,} |")

        if files:
            lines.extend([
                "",
                "## Files Analyzed",
                "",
                "| File | Lines | Size |",
                "|------|-------|------|"
            ])

            for f in files[:50]:
                name = f.get('name', 'Unknown')
                line_count = f.get('line_count', 0)
                size = self._format_size(f.get('size', 0))
                lines.append(f"| {name} | {line_count:,} | {size} |")

        lines.extend([
            "",
            "---",
            "*Generated by CodeExtractPro*"
        ])

        return '\n'.join(lines)

    def _build_text(self, data: Dict[str, Any]) -> str:
        """Build plain text content."""
        stats = data.get('statistics', {})
        files = data.get('files', [])

        lines = [
            "=" * 80,
            f" {self.title}",
            "=" * 80,
            f"Generated: {self.generated_at.strftime('%Y-%m-%d %H:%M:%S')}",
            "",
            "-" * 40,
            "SUMMARY STATISTICS",
            "-" * 40,
        ]

        for key, value in stats.items():
            if isinstance(value, (int, float)):
                label = key.replace('_', ' ').title()
                lines.append(f"  {label}: {value:,}")

        if files:
            lines.extend([
                "",
                "-" * 40,
                "FILES ANALYZED",
                "-" * 40,
            ])

            for f in files[:50]:
                name = f.get('name', 'Unknown')
                line_count = f.get('line_count', 0)
                lines.append(f"  - {name} ({line_count:,} lines)")

        lines.extend([
            "",
            "=" * 80,
            "Generated by CodeExtractPro",
            "=" * 80
        ])

        return '\n'.join(lines)

    def _format_size(self, size: int) -> str:
        """Format size to human-readable string."""
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size < 1024:
                return f"{size:.1f} {unit}" if unit != 'B' else f"{size} {unit}"
            size /= 1024
        return f"{size:.1f} TB"
