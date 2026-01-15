# CodeExtractPro - VBA Code Extractor & Python Analyzer

**Professional VBA extraction tool for Microsoft Office files with Python code analysis capabilities**

[![Python](https://img.shields.io/badge/Python-3.9+-blue.svg)](https://python.org)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)](https://github.com/Kiriiaq/CodeExtract/releases)
[![Release](https://img.shields.io/badge/Release-v1.0.0-orange.svg)](https://github.com/Kiriiaq/CodeExtract/releases/latest)

---

## What is CodeExtractPro?

CodeExtractPro solves a common problem for developers and analysts: **extracting VBA macros from Microsoft Office files** (Excel, Word, PowerPoint) quickly and reliably. It also provides tools for analyzing Python projects and scanning folder structures.

**Key benefits:**
- Extract VBA code from `.xlsm`, `.xlsb`, `.docm`, `.pptm` and other macro-enabled Office formats
- Analyze Python project structure and code metrics
- Scan directories with detailed file information
- Optimize and clean VBA code
- Export results to JSON, CSV, or HTML

---

## Quick Start (Recommended)

### Download Ready-to-Use Executables

Download from [**Releases**](https://github.com/Kiriiaq/CodeExtract/releases/latest):

| File | Description | Size |
|------|-------------|------|
| `CodeExtractPro.exe` | **Production version** - No console window, clean UI | ~40 MB |
| `CodeExtractPro_Debug.exe` | **Debug version** - Console output for troubleshooting | ~40 MB |

**Just download and double-click to run.** No installation or Python required.

### Which Version Should I Use?

- **CodeExtractPro.exe** (Production): For everyday use. Clean interface without console window.
- **CodeExtractPro_Debug.exe** (Debug): Use when experiencing issues. Shows detailed logs in a console window for troubleshooting.

---

## Features

### VBA Extractor
Extract VBA code from Microsoft Office files (Excel, Word, PowerPoint).

| Option | Description |
|--------|-------------|
| Individual files | Export each module as separate `.bas`, `.cls`, `.frm` files |
| Concatenated file | Combine all modules into a single file |
| Extraction method | Auto, Win32COM (Windows), OleTools (cross-platform) |
| Include metadata | Add file information headers |
| Preserve formatting | Keep original code formatting |
| Extract UserForms | Include form definitions |
| Line numbers | Add line numbers to output |
| Encoding | UTF-8, UTF-8-BOM, Latin-1, CP1252, ASCII |

**Supported formats:** `.xlsm`, `.xlsb`, `.xls`, `.xla`, `.xlam`, `.docm`, `.doc`, `.dotm`, `.pptm`, `.ppt`, `.potm`, `.ppsm`

### Python Analyzer
Analyze Python code structure and quality metrics.

| Option | Description |
|--------|-------------|
| Subdirectories | Recursively analyze subfolders |
| Follow symlinks | Follow symbolic links |
| Analyze imports | Track import statements |
| Analyze complexity | Calculate cyclomatic complexity |
| Count docstrings | Measure documentation coverage |
| Detect duplicates | Find duplicate code blocks |
| Exclude tests | Skip test files (`test_*.py`, `*_test.py`) |
| Exclude __init__ | Skip `__init__.py` files |

**Metrics:** Files, lines, code lines, classes, functions, documentation ratio

### Folder Scanner
Scan and analyze directory structures.

| Option | Description |
|--------|-------------|
| Include content | Extract file contents |
| Include binaries | Include binary files |
| Hidden files | Show hidden files/folders |
| Max KB | Maximum file size limit |
| Depth | Scan depth (1-10 or unlimited) |
| Exclude .git/.svn | Skip version control folders |
| Exclude __pycache__ | Skip Python cache |
| Exclude node_modules | Skip npm packages |
| Show sizes | Display file sizes |
| Show dates | Display modification dates |
| Output style | Tree, flat list, or JSON |

### VBA Optimizer
Optimize and clean VBA code.

| Option | Description |
|--------|-------------|
| Remove comments | Strip all VBA comments |
| Remove empty lines | Remove blank lines |
| Remove Debug.* | Strip debug statements |
| Auto-indentation | Automatic code indentation |
| Indent size | Spaces per indent level (1-8) |
| Normalize case | Standardize keyword casing |
| Minify | Compress code to minimum size |
| Rename variables | Shorten variable names |
| Obfuscate (beta) | Code obfuscation |
| Create backup | Backup before changes |
| Validate syntax | Check VBA syntax validity |
| Preview only | Show changes without saving |

---

## Installation from Source (Developers)

### Prerequisites
- Python 3.9+
- pip

### Setup

```bash
# Clone the repository
git clone https://github.com/Kiriiaq/CodeExtract.git
cd CodeExtract

# Install dependencies
pip install -r requirements.txt

# Run the application
python main.py
```

### Dependencies

| Package | Version | Required | Description |
|---------|---------|----------|-------------|
| customtkinter | >= 5.2.0 | Yes | Modern UI framework |
| oletools | >= 0.60 | Optional | VBA extraction (cross-platform) |
| pywin32 | >= 306 | Optional | Excel automation (Windows only) |

---

## Building Executables

```bash
# Install build dependencies
pip install pyinstaller customtkinter oletools pywin32

# Build both versions
pyinstaller --clean --noconfirm build_release.spec
pyinstaller --clean --noconfirm build_debug.spec

# Or use the build script (Windows)
build.bat
```

**Output:**
- `dist/CodeExtractPro.exe` - Production build (no console)
- `dist/CodeExtractPro_Debug.exe` - Debug build (with console)

---

## Usage

### Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `F1` | Open help dialog |
| `Ctrl+1` | Switch to VBA Extractor |
| `Ctrl+2` | Switch to Python Analyzer |
| `Ctrl+3` | Switch to Folder Scanner |
| `Ctrl+4` | Switch to VBA Optimizer |
| `Ctrl+Q` | Quit application |

### Export Formats

All tools support exporting results to:
- **JSON** - Structured data format
- **CSV** - Spreadsheet compatible
- **HTML** - Interactive web reports

### Settings

- **Theme**: Dark, Light, System
- **Color scheme**: Blue, Green, Dark-blue
- **Auto-open exports**: Open files after export
- **Configuration backup**: Export/import settings

---

## Project Structure

```
CodeExtractPro/
├── main.py                 # Application entry point
├── requirements.txt        # Python dependencies
├── build_release.spec      # PyInstaller config (release)
├── build_debug.spec        # PyInstaller config (debug)
├── build.bat               # Windows build script
├── LICENSE                 # MIT License
├── README.md               # This file
├── assets/                 # Icons and resources
├── src/
│   ├── core/
│   │   ├── config_manager.py    # Configuration management
│   │   ├── export_manager.py    # Multi-format export
│   │   ├── logging_system.py    # Logging infrastructure
│   │   └── workflow.py          # Workflow management
│   ├── modules/
│   │   ├── vba_extractor.py     # VBA extraction module
│   │   ├── python_analyzer.py   # Python analysis module
│   │   ├── folder_scanner.py    # Directory scanning module
│   │   ├── vba_optimizer.py     # VBA optimization module
│   │   └── report_generator.py  # Report generation
│   ├── ui/
│   │   └── main_window.py       # GUI implementation
│   └── utils/
│       ├── widgets.py           # Custom UI components
│       └── helpers.py           # Utility functions
└── tests/
    └── test_modules.py          # Unit tests
```

---

## Configuration

Configuration is stored in:
- **Windows**: `%USERPROFILE%\.codeextractpro\config.json`

Settings are automatically saved and include:
- Window size and position
- Theme preferences
- Default options for each tool
- Export preferences

---

## License

MIT License - See [LICENSE](LICENSE) for details.

---

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

---

## Support

- **Issues**: [GitHub Issues](https://github.com/Kiriiaq/CodeExtract/issues)
- **Discussions**: [GitHub Discussions](https://github.com/Kiriiaq/CodeExtract/discussions)

---

## Keywords

VBA extractor, Excel macro extractor, Office VBA tool, Python code analyzer, folder scanner, VBA optimizer, extract macros from Excel, xlsm vba extraction, docm macro extraction, code analysis tool

---

*Built with Python and CustomTkinter*
