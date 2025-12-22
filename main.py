#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CodeExtractPro v2.0 - Professional Code Extraction & Analysis Suite
Main entry point for the application.

Usage:
    python main.py          # Launch GUI
    python main.py --cli    # Launch CLI mode (coming soon)
"""

import sys
import os
import io

# ========== PYINSTALLER FIX ==========
# Redirect stdout/stderr for PyInstaller compatibility
if sys.stdout is None:
    sys.stdout = io.StringIO()
if sys.stderr is None:
    sys.stderr = io.StringIO()

# Disable colorclass for oletools compatibility
os.environ['COLORCLASS_DISABLE'] = '1'
os.environ['PYTHONIOENCODING'] = 'utf-8'

# Suppress warnings
import warnings
warnings.filterwarnings("ignore", category=UserWarning)

# ========== PATH SETUP ==========
# Add src directory to path
src_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'src')
if src_path not in sys.path:
    sys.path.insert(0, src_path)

# ========== HI-DPI AWARENESS (Windows) ==========
if sys.platform == "win32":
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass


def check_dependencies():
    """Check and report on available dependencies."""
    deps = {
        'customtkinter': False,
        'oletools': False,
        'win32com': False,
    }

    try:
        import customtkinter
        deps['customtkinter'] = True
    except ImportError:
        pass

    try:
        from oletools.olevba import VBA_Parser
        deps['oletools'] = True
    except ImportError:
        pass

    try:
        import win32com.client
        deps['win32com'] = True
    except ImportError:
        pass

    return deps


def install_missing_deps():
    """Attempt to install missing critical dependencies."""
    import subprocess

    deps = check_dependencies()

    if not deps['customtkinter']:
        print("Installing customtkinter...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "customtkinter"])
            print("customtkinter installed successfully")
        except Exception as e:
            print(f"Failed to install customtkinter: {e}")
            return False

    return True


def main():
    """Main entry point."""
    print("=" * 60)
    print(" CodeExtractPro v2.0 - Professional Code Extraction Suite")
    print("=" * 60)
    print()

    # Check dependencies
    deps = check_dependencies()
    print("Checking dependencies:")
    print(f"  - customtkinter: {'OK' if deps['customtkinter'] else 'MISSING (required)'}")
    print(f"  - oletools: {'OK' if deps['oletools'] else 'Optional (VBA extraction)'}")
    print(f"  - win32com: {'OK' if deps['win32com'] else 'Optional (Windows Excel automation)'}")
    print()

    # Install missing critical dependencies
    if not deps['customtkinter']:
        print("Critical dependency missing. Attempting to install...")
        if not install_missing_deps():
            print("\nPlease install dependencies manually:")
            print("  pip install customtkinter oletools pywin32")
            input("\nPress Enter to exit...")
            sys.exit(1)

    # Import and launch application
    try:
        from src.ui.main_window import MainWindow

        print("Starting application...")
        print()

        app = MainWindow()
        app.run()

    except ImportError as e:
        print(f"\nImport error: {e}")
        print("\nPlease ensure all dependencies are installed:")
        print("  pip install -r requirements.txt")
        input("\nPress Enter to exit...")
        sys.exit(1)

    except Exception as e:
        print(f"\nError starting application: {e}")
        import traceback
        traceback.print_exc()
        input("\nPress Enter to exit...")
        sys.exit(1)


def cli_main():
    """CLI mode entry point (future implementation)."""
    print("CLI mode is not yet implemented.")
    print("Please use the GUI: python main.py")


if __name__ == "__main__":
    # Check for CLI mode
    if "--cli" in sys.argv:
        cli_main()
    else:
        main()
