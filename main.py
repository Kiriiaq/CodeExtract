#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CodeExtractPro v1.0 - Professional Code Extraction & Analysis Suite
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
# Get base directory (handles both normal and PyInstaller execution)
if getattr(sys, 'frozen', False):
    # Running as PyInstaller bundle
    base_path = sys._MEIPASS
else:
    # Running as script
    base_path = os.path.dirname(os.path.abspath(__file__))

# Add src directory to path
src_path = os.path.join(base_path, 'src')
if src_path not in sys.path:
    sys.path.insert(0, src_path)

# Also add base_path for package imports
if base_path not in sys.path:
    sys.path.insert(0, base_path)

# ========== HI-DPI AWARENESS (Windows) ==========
if sys.platform == "win32":
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

# ========== TASKBAR ICON FIX (Windows) ==========
if sys.platform == "win32":
    try:
        import ctypes
        # Set AppUserModelID for proper taskbar icon display
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("CodeExtractPro.v1.0")
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
    print(" CodeExtractPro v1.0 - Professional Code Extraction Suite")
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
            try:
                input("\nPress Enter to exit...")
            except (RuntimeError, EOFError):
                # No stdin available (PyInstaller windowed mode)
                pass
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
        try:
            input("\nPress Enter to exit...")
        except (RuntimeError, EOFError):
            # No stdin available (PyInstaller windowed mode)
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Import Error", f"Failed to import: {e}\n\nPlease ensure all dependencies are installed:\npip install -r requirements.txt")
            root.destroy()
        sys.exit(1)

    except Exception as e:
        print(f"\nError starting application: {e}")
        import traceback
        traceback.print_exc()
        try:
            input("\nPress Enter to exit...")
        except (RuntimeError, EOFError):
            # No stdin available (PyInstaller windowed mode)
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Error", f"Error starting application:\n{e}")
            root.destroy()
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
