"""
Main Window - Modern GUI for CodeExtractPro v1.0.
Each tool is completely independent with its own interface, configuration, and export capabilities.
Features: dark/light themes, keyboard shortcuts, integrated help, tooltips.
"""

import os
import sys
import threading
import webbrowser
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import Dict, List, Optional, Any, Callable

import customtkinter as ctk

# Add src to path for PyInstaller compatibility
_src_path = str(Path(__file__).parent.parent)
if _src_path not in sys.path:
    sys.path.insert(0, _src_path)

# Also add the parent of src (project root) for 'src.x' style imports
_root_path = str(Path(__file__).parent.parent.parent)
if _root_path not in sys.path:
    sys.path.insert(0, _root_path)

try:
    # Try relative imports first (for running as package)
    from ..core.config_manager import get_config
    from ..core.export_manager import get_export_manager
    from ..core.logging_system import LogEntry, get_logger
    from ..modules.vba_extractor import VBAExtractor, ExtractionMethod
    from ..modules.python_analyzer import PythonAnalyzer
    from ..modules.folder_scanner import FolderScanner
    from ..modules.vba_optimizer import VBAOptimizer, OptimizationOptions
    from ..modules.vba_analyzer import VBAAnalyzer, get_hex_preview, is_binary_file
    from ..utils.widgets import ToolTip
except ImportError:
    # Fall back to absolute imports (for PyInstaller or direct execution)
    from core.config_manager import get_config
    from core.export_manager import get_export_manager
    from core.logging_system import LogEntry, get_logger
    from modules.vba_extractor import VBAExtractor, ExtractionMethod
    from modules.python_analyzer import PythonAnalyzer
    from modules.folder_scanner import FolderScanner
    from modules.vba_optimizer import VBAOptimizer, OptimizationOptions
    from modules.vba_analyzer import VBAAnalyzer, get_hex_preview, is_binary_file
    from utils.widgets import ToolTip


class HelpSystem:
    """Integrated help system."""
    HELP_TEXTS = {
        "vba_extractor": {"title": "VBA Extractor", "description": "Extracts VBA code from Office files.", "usage": ["1. Select Office file", "2. Choose options", "3. Click Extract"], "tips": ["Win32COM works best on Windows"]},
        "python_analyzer": {"title": "Python Analyzer", "description": "Analyzes Python code structure.", "usage": ["1. Select directory", "2. Configure options", "3. Click Analyze"], "tips": ["Enable subdirs for full projects"]},
        "folder_scanner": {"title": "Folder Scanner", "description": "Scans directory structures.", "usage": ["1. Select directory", "2. Configure filters", "3. Click Scan"], "tips": ["Exclude .git for faster scans"]},
        "vba_optimizer": {"title": "VBA Optimizer", "description": "Optimizes VBA code.", "usage": ["1. Load code", "2. Select options", "3. Click Optimize"], "tips": ["Backup before minifying"]}
    }

    @classmethod
    def get_help(cls, tool_id): return cls.HELP_TEXTS.get(tool_id, {})

    @classmethod
    def show_help_dialog(cls, parent, tool_id):
        data = cls.get_help(tool_id)
        if not data: return
        dlg = ctk.CTkToplevel(parent)
        dlg.title(f"Help - {data.get('title', '')}")
        dlg.geometry("450x350")
        dlg.transient(parent)
        dlg.grab_set()
        scroll = ctk.CTkScrollableFrame(dlg)
        scroll.pack(fill="both", expand=True, padx=15, pady=15)
        ctk.CTkLabel(scroll, text=data.get('title', ''), font=ctk.CTkFont(size=18, weight="bold")).pack(anchor="w")
        ctk.CTkLabel(scroll, text=data.get('description', ''), font=ctk.CTkFont(size=11), wraplength=400).pack(anchor="w", pady=10)
        ctk.CTkLabel(scroll, text="Usage:", font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w")
        for s in data.get('usage', []): ctk.CTkLabel(scroll, text=s, font=ctk.CTkFont(size=10)).pack(anchor="w", padx=10)
        ctk.CTkLabel(scroll, text="Tips:", font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", pady=(10,0))
        for t in data.get('tips', []): ctk.CTkLabel(scroll, text=f"* {t}", font=ctk.CTkFont(size=10)).pack(anchor="w", padx=10)
        ctk.CTkButton(dlg, text="Close", command=dlg.destroy, width=80).pack(pady=10)


class BaseToolFrame(ctk.CTkFrame):
    """Base class for tool frames."""
    def __init__(self, parent, tool_id, tool_name, **kw):
        super().__init__(parent, **kw)
        self.tool_id, self.tool_name = tool_id, tool_name
        self.config, self.export_manager, self.logger = get_config(), get_export_manager(), get_logger()
        self.is_running, self._last_result = False, None
        self._create_header()
        self._create_content()
        self._create_footer()

    def _create_header(self):
        hdr = ctk.CTkFrame(self, fg_color="transparent")
        hdr.pack(fill="x", padx=15, pady=(15, 10))
        ctk.CTkLabel(hdr, text=self.tool_name, font=ctk.CTkFont(size=20, weight="bold")).pack(side="left")
        btn = ctk.CTkButton(hdr, text="?", width=28, height=28, command=lambda: HelpSystem.show_help_dialog(self.winfo_toplevel(), self.tool_id))
        btn.pack(side="right")
        ToolTip(btn, "Help")

    def _create_content(self):
        self.content_frame = ctk.CTkFrame(self)
        self.content_frame.pack(fill="both", expand=True, padx=15, pady=5)

    def _create_footer(self):
        ftr = ctk.CTkFrame(self, fg_color="transparent")
        ftr.pack(fill="x", padx=10, pady=(3, 8))
        self.progress_bar = ctk.CTkProgressBar(ftr, height=4)
        self.progress_bar.pack(fill="x", pady=(0, 4))
        self.progress_bar.set(0)
        self.status_label = ctk.CTkLabel(ftr, text="Ready", font=ctk.CTkFont(size=9), text_color=("gray50", "gray60"))
        self.status_label.pack(side="left")
        exp = ctk.CTkFrame(ftr, fg_color="transparent")
        exp.pack(side="right")
        for f, t in [("json", "JSON"), ("csv", "CSV"), ("html", "HTML")]:
            b = ctk.CTkButton(exp, text=t, width=45, height=22, font=ctk.CTkFont(size=9), command=lambda x=f: self._export_result(x))
            b.pack(side="left", padx=1)

    def _export_result(self, fmt):
        if not self._last_result:
            messagebox.showwarning("Warning", "No results to export")
            return
        path = filedialog.asksaveasfilename(title=f"Export {fmt.upper()}", defaultextension=f".{fmt}",
                                             filetypes=[(f"{fmt.upper()} files", f"*.{fmt}"), ("All files", "*.*")])
        if path:
            r = self.export_manager.export(self._last_result, path, fmt)
            if r.success:
                messagebox.showinfo("Success", f"Exported to {path}")
                if self.config.config.export.open_after_export: webbrowser.open(path)
            else:
                messagebox.showerror("Error", f"Export failed: {r.message}")

    def set_progress(self, val, status=None):
        self.progress_bar.set(val)
        if status: self.status_label.configure(text=status)

    def run_async(self, func, callback=None):
        if self.is_running: return
        self.is_running = True
        self.set_progress(0, "Processing...")
        def run():
            try:
                res = func()
                self.after(0, lambda: self._done(res, callback))
            except Exception as e:
                self.after(0, lambda: self._err(str(e)))
        threading.Thread(target=run, daemon=True).start()

    def _done(self, res, cb):
        self.is_running = False
        self.set_progress(1.0, "Complete")
        if cb: cb(res)

    def _err(self, e):
        self.is_running = False
        self.set_progress(0, f"Error: {e}")
        messagebox.showerror("Error", e)


class VBAExtractorFrame(BaseToolFrame):
    def __init__(self, parent, **kw):
        self.extractor = VBAExtractor()
        super().__init__(parent, "vba_extractor", "VBA Extractor", **kw)

    def _create_content(self):
        self.content_frame = ctk.CTkFrame(self)
        self.content_frame.pack(fill="both", expand=True, padx=10, pady=3)

        # Panneau gauche avec scrollbar pour les options
        left = ctk.CTkFrame(self.content_frame, width=320)
        left.pack(side="left", fill="both", padx=(0, 4))
        left.pack_propagate(False)

        # Source file section
        ctk.CTkLabel(left, text="📁 Source File", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", padx=8, pady=(8, 4))
        ef = ctk.CTkFrame(left, fg_color="transparent")
        ef.pack(fill="x", padx=8)
        self.file_var = ctk.StringVar()
        ctk.CTkEntry(ef, textvariable=self.file_var, placeholder_text="Select Office file...", height=28).pack(side="left", fill="x", expand=True, padx=(0, 4))
        ctk.CTkButton(ef, text="...", width=35, height=28, command=self._browse).pack(side="right")

        # Frame scrollable pour les options
        opts_scroll = ctk.CTkScrollableFrame(left, label_text="⚙️ Options", label_font=ctk.CTkFont(size=11, weight="bold"),
                                              fg_color=("gray88", "gray20"), corner_radius=8)
        opts_scroll.pack(fill="both", expand=True, padx=8, pady=8)

        # === Section: Output Format ===
        out_frame = ctk.CTkFrame(opts_scroll, fg_color=("gray95", "gray25"), corner_radius=6)
        out_frame.pack(fill="x", pady=4)
        ctk.CTkLabel(out_frame, text="Format de sortie", font=ctk.CTkFont(size=10, weight="bold"),
                     text_color=("#059669", "#10b981")).pack(anchor="w", padx=8, pady=(6, 2))

        cfg = self.config.config.vba_extractor
        self.indiv_var = ctk.BooleanVar(value=cfg.create_individual_files)
        self.concat_var = ctk.BooleanVar(value=cfg.create_concatenated_file)
        ctk.CTkCheckBox(out_frame, text="Fichiers individuels (.bas, .cls)", variable=self.indiv_var,
                        font=ctk.CTkFont(size=10), height=22, checkbox_width=18, checkbox_height=18).pack(anchor="w", padx=12, pady=1)
        ctk.CTkCheckBox(out_frame, text="Fichier concaténé unique", variable=self.concat_var,
                        font=ctk.CTkFont(size=10), height=22, checkbox_width=18, checkbox_height=18).pack(anchor="w", padx=12, pady=(1, 6))

        # === Section: Extraction Method ===
        meth_frame = ctk.CTkFrame(opts_scroll, fg_color=("gray95", "gray25"), corner_radius=6)
        meth_frame.pack(fill="x", pady=4)
        ctk.CTkLabel(meth_frame, text="Méthode d'extraction", font=ctk.CTkFont(size=10, weight="bold"),
                     text_color=("#059669", "#10b981")).pack(anchor="w", padx=8, pady=(6, 2))

        self.method_var = ctk.StringVar(value=cfg.extraction_method)
        ctk.CTkOptionMenu(meth_frame, values=["auto"] + self.extractor.get_available_methods(),
                          variable=self.method_var, width=140, height=26,
                          font=ctk.CTkFont(size=10)).pack(anchor="w", padx=12, pady=(2, 6))

        # === Section: Advanced Options (NEW) ===
        adv_frame = ctk.CTkFrame(opts_scroll, fg_color=("gray95", "gray25"), corner_radius=6)
        adv_frame.pack(fill="x", pady=4)
        ctk.CTkLabel(adv_frame, text="🔧 Options avancées", font=ctk.CTkFont(size=10, weight="bold"),
                     text_color=("#f59e0b", "#fbbf24")).pack(anchor="w", padx=8, pady=(6, 2))

        self.include_metadata_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(adv_frame, text="Inclure métadonnées", variable=self.include_metadata_var,
                        font=ctk.CTkFont(size=10), height=22, checkbox_width=18, checkbox_height=18).pack(anchor="w", padx=12, pady=1)

        self.preserve_formatting_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(adv_frame, text="Préserver formatage", variable=self.preserve_formatting_var,
                        font=ctk.CTkFont(size=10), height=22, checkbox_width=18, checkbox_height=18).pack(anchor="w", padx=12, pady=1)

        self.extract_forms_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(adv_frame, text="Extraire UserForms", variable=self.extract_forms_var,
                        font=ctk.CTkFont(size=10), height=22, checkbox_width=18, checkbox_height=18).pack(anchor="w", padx=12, pady=1)

        self.add_line_numbers_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(adv_frame, text="Numéros de ligne", variable=self.add_line_numbers_var,
                        font=ctk.CTkFont(size=10), height=22, checkbox_width=18, checkbox_height=18).pack(anchor="w", padx=12, pady=(1, 6))

        # === Section: Encoding (NEW) ===
        enc_frame = ctk.CTkFrame(opts_scroll, fg_color=("gray95", "gray25"), corner_radius=6)
        enc_frame.pack(fill="x", pady=4)
        ctk.CTkLabel(enc_frame, text="Encodage", font=ctk.CTkFont(size=10, weight="bold"),
                     text_color=("#059669", "#10b981")).pack(anchor="w", padx=8, pady=(6, 2))

        self.encoding_var = ctk.StringVar(value="utf-8")
        ctk.CTkOptionMenu(enc_frame, values=["utf-8", "utf-8-bom", "latin-1", "cp1252", "ascii"],
                          variable=self.encoding_var, width=120, height=26,
                          font=ctk.CTkFont(size=10)).pack(anchor="w", padx=12, pady=(2, 6))

        # Bouton Extract
        ctk.CTkButton(left, text="▶ Extract VBA", command=self._extract, height=36,
                      font=ctk.CTkFont(size=12, weight="bold"),
                      fg_color=("#10b981", "#059669"), hover_color=("#059669", "#047857")).pack(padx=8, pady=8, fill="x")

        # Panneau droit - Résultats
        right = ctk.CTkFrame(self.content_frame)
        right.pack(side="right", fill="both", expand=True, padx=(4, 0))
        ctk.CTkLabel(right, text="📄 Results", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", padx=8, pady=6)
        self.results_text = ctk.CTkTextbox(right, font=ctk.CTkFont(family="Consolas", size=9))
        self.results_text.pack(fill="both", expand=True, padx=8, pady=(0, 8))

    def _browse(self):
        p = filedialog.askopenfilename(title="Select Office File", filetypes=[("Office files", "*.xlsm;*.xlsb;*.xls;*.docm;*.pptm"), ("All", "*.*")])
        if p: self.file_var.set(p)

    def _extract(self):
        fp = self.file_var.get()
        if not fp:
            messagebox.showerror("Error", "Select a file first")
            return
        out = filedialog.askdirectory(title="Output Directory")
        if not out: return
        def do():
            m = self.method_var.get()
            if m != "auto": self.extractor.preferred_method = ExtractionMethod(m)
            return self.extractor.extract(fp, out, self.indiv_var.get(), self.concat_var.get())
        def done(r):
            self._last_result = {"success": r.success, "modules": r.total_modules, "lines": r.total_lines}
            self.results_text.delete("1.0", "end")
            if r.success:
                self.results_text.insert("end", f"Extracted {r.total_modules} modules ({r.total_lines} lines)\n\n")
                for m in r.modules:
                    self.results_text.insert("end", f"* {m.name}.{m.extension} - {m.line_count} lines\n")
                self.set_progress(1.0, f"{r.total_modules} modules extracted")
            else:
                self.results_text.insert("end", f"Error: {r.error_message}")
                self.set_progress(0, "Failed")
        self.run_async(do, done)


class PythonAnalyzerFrame(BaseToolFrame):
    def __init__(self, parent, **kw):
        self.analyzer = PythonAnalyzer()
        super().__init__(parent, "python_analyzer", "Python Analyzer", **kw)

    def _create_content(self):
        self.content_frame = ctk.CTkFrame(self)
        self.content_frame.pack(fill="both", expand=True, padx=10, pady=3)

        # Top section - Directory and options
        top = ctk.CTkFrame(self.content_frame)
        top.pack(fill="x", padx=4, pady=4)

        # Directory selection
        dir_frame = ctk.CTkFrame(top, fg_color="transparent")
        dir_frame.pack(fill="x", pady=4)
        ctk.CTkLabel(dir_frame, text="📁 Source Directory", font=ctk.CTkFont(size=11, weight="bold")).pack(side="left", padx=4)
        self.dir_var = ctk.StringVar()
        ctk.CTkEntry(dir_frame, textvariable=self.dir_var, placeholder_text="Select Python directory...", height=28).pack(side="left", fill="x", expand=True, padx=4)
        ctk.CTkButton(dir_frame, text="...", width=35, height=28, command=self._browse).pack(side="left", padx=(0, 4))

        # Action buttons
        btn_frame = ctk.CTkFrame(dir_frame, fg_color="transparent")
        btn_frame.pack(side="right")
        ctk.CTkButton(btn_frame, text="▶ Analyze", command=self._analyze, height=28, width=80,
                      font=ctk.CTkFont(size=10, weight="bold"),
                      fg_color=("#10b981", "#059669"), hover_color=("#059669", "#047857")).pack(side="left", padx=2)
        ctk.CTkButton(btn_frame, text="📄 Extract", command=self._extract_code, height=28, width=80,
                      font=ctk.CTkFont(size=10, weight="bold"),
                      fg_color=("#3b82f6", "#2563eb"), hover_color=("#2563eb", "#1d4ed8")).pack(side="left", padx=2)

        # Options scrollable frame
        opts_scroll = ctk.CTkScrollableFrame(top, height=120, label_text="⚙️ Options",
                                              label_font=ctk.CTkFont(size=10, weight="bold"),
                                              fg_color=("gray88", "gray20"), corner_radius=6)
        opts_scroll.pack(fill="x", padx=4, pady=4)

        # Options grid
        opts_grid = ctk.CTkFrame(opts_scroll, fg_color="transparent")
        opts_grid.pack(fill="x")

        cfg = self.config.config.python_analyzer

        # Column 1 - Basic options
        col1 = ctk.CTkFrame(opts_grid, fg_color=("gray95", "gray25"), corner_radius=6)
        col1.pack(side="left", fill="both", expand=True, padx=2, pady=2)
        ctk.CTkLabel(col1, text="Parcours", font=ctk.CTkFont(size=9, weight="bold"),
                     text_color=("#059669", "#10b981")).pack(anchor="w", padx=6, pady=(4, 2))
        self.subdirs_var = ctk.BooleanVar(value=cfg.include_subdirs)
        ctk.CTkCheckBox(col1, text="Sous-dossiers", variable=self.subdirs_var,
                        font=ctk.CTkFont(size=9), height=20, checkbox_width=16, checkbox_height=16).pack(anchor="w", padx=8, pady=1)
        self.follow_symlinks_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(col1, text="Suivre liens", variable=self.follow_symlinks_var,
                        font=ctk.CTkFont(size=9), height=20, checkbox_width=16, checkbox_height=16).pack(anchor="w", padx=8, pady=1)
        self.include_content_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(col1, text="Inclure code", variable=self.include_content_var,
                        font=ctk.CTkFont(size=9), height=20, checkbox_width=16, checkbox_height=16).pack(anchor="w", padx=8, pady=(1, 4))

        # Column 2 - Analysis depth
        col2 = ctk.CTkFrame(opts_grid, fg_color=("gray95", "gray25"), corner_radius=6)
        col2.pack(side="left", fill="both", expand=True, padx=2, pady=2)
        ctk.CTkLabel(col2, text="Analyse", font=ctk.CTkFont(size=9, weight="bold"),
                     text_color=("#059669", "#10b981")).pack(anchor="w", padx=6, pady=(4, 2))
        self.analyze_imports_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(col2, text="Imports", variable=self.analyze_imports_var,
                        font=ctk.CTkFont(size=9), height=20, checkbox_width=16, checkbox_height=16).pack(anchor="w", padx=8, pady=1)
        self.analyze_complexity_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(col2, text="Complexité", variable=self.analyze_complexity_var,
                        font=ctk.CTkFont(size=9), height=20, checkbox_width=16, checkbox_height=16).pack(anchor="w", padx=8, pady=1)
        # Max file size
        size_row = ctk.CTkFrame(col2, fg_color="transparent")
        size_row.pack(fill="x", padx=8, pady=(1, 4))
        ctk.CTkLabel(size_row, text="Max KB:", font=ctk.CTkFont(size=8)).pack(side="left")
        self.max_size_var = ctk.StringVar(value="500")
        ctk.CTkEntry(size_row, textvariable=self.max_size_var, width=45, height=18, font=ctk.CTkFont(size=8)).pack(side="left", padx=2)

        # Column 3 - Exclusions
        col3 = ctk.CTkFrame(opts_grid, fg_color=("gray95", "gray25"), corner_radius=6)
        col3.pack(side="left", fill="both", expand=True, padx=2, pady=2)
        ctk.CTkLabel(col3, text="Exclure fichiers", font=ctk.CTkFont(size=9, weight="bold"),
                     text_color=("#f59e0b", "#fbbf24")).pack(anchor="w", padx=6, pady=(4, 2))
        self.exclude_tests_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(col3, text="test_*.py", variable=self.exclude_tests_var,
                        font=ctk.CTkFont(size=9), height=20, checkbox_width=16, checkbox_height=16).pack(anchor="w", padx=8, pady=1)
        self.exclude_init_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(col3, text="__init__.py", variable=self.exclude_init_var,
                        font=ctk.CTkFont(size=9), height=20, checkbox_width=16, checkbox_height=16).pack(anchor="w", padx=8, pady=1)
        self.exclude_setup_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(col3, text="setup.py", variable=self.exclude_setup_var,
                        font=ctk.CTkFont(size=9), height=20, checkbox_width=16, checkbox_height=16).pack(anchor="w", padx=8, pady=(1, 4))

        # Column 5 - Regex filter with examples
        col5 = ctk.CTkFrame(opts_grid, fg_color=("gray95", "gray25"), corner_radius=6)
        col5.pack(side="left", fill="both", expand=True, padx=2, pady=2)
        ctk.CTkLabel(col5, text="Filtre Regex", font=ctk.CTkFont(size=9, weight="bold"),
                     text_color=("#8b5cf6", "#a78bfa")).pack(anchor="w", padx=6, pady=(4, 2))
        self.regex_var = ctk.StringVar()
        ctk.CTkEntry(col5, textvariable=self.regex_var, placeholder_text="Ex: ^main.*",
                     height=22, font=ctk.CTkFont(size=9)).pack(fill="x", padx=8, pady=2)
        # Regex examples buttons
        regex_btn_frame = ctk.CTkFrame(col5, fg_color="transparent")
        regex_btn_frame.pack(fill="x", padx=8, pady=(0, 4))
        regex_examples = [("main", "^main"), ("config", "config"), ("!test", "^(?!test)")]
        for label, pattern in regex_examples:
            ctk.CTkButton(regex_btn_frame, text=label, width=40, height=18,
                          font=ctk.CTkFont(size=8),
                          command=lambda p=pattern: self.regex_var.set(p)).pack(side="left", padx=1)

        # Column 4 - Exclude directories
        col4 = ctk.CTkFrame(opts_grid, fg_color=("gray95", "gray25"), corner_radius=6)
        col4.pack(side="left", fill="both", expand=True, padx=2, pady=2)
        ctk.CTkLabel(col4, text="Exclure dossiers", font=ctk.CTkFont(size=9, weight="bold"),
                     text_color=("#ef4444", "#f87171")).pack(anchor="w", padx=6, pady=(4, 2))
        self.exclude_venv_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(col4, text="venv / .venv", variable=self.exclude_venv_var,
                        font=ctk.CTkFont(size=9), height=20, checkbox_width=16, checkbox_height=16).pack(anchor="w", padx=8, pady=1)
        self.exclude_pycache_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(col4, text="__pycache__", variable=self.exclude_pycache_var,
                        font=ctk.CTkFont(size=9), height=20, checkbox_width=16, checkbox_height=16).pack(anchor="w", padx=8, pady=1)
        self.exclude_git_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(col4, text=".git / .idea", variable=self.exclude_git_var,
                        font=ctk.CTkFont(size=9), height=20, checkbox_width=16, checkbox_height=16).pack(anchor="w", padx=8, pady=(1, 4))

        # Results section
        res = ctk.CTkFrame(self.content_frame)
        res.pack(fill="both", expand=True, padx=4, pady=(0, 4))

        # Statistics panel (left)
        sp = ctk.CTkFrame(res, width=220)
        sp.pack(side="left", fill="y", padx=(0, 4))
        sp.pack_propagate(False)
        ctk.CTkLabel(sp, text="📊 Statistics", font=ctk.CTkFont(size=11, weight="bold")).pack(anchor="w", padx=6, pady=6)
        self.stats_text = ctk.CTkTextbox(sp, font=ctk.CTkFont(size=9))
        self.stats_text.pack(fill="both", expand=True, padx=6, pady=(0, 6))

        # Files panel (right)
        fp = ctk.CTkFrame(res)
        fp.pack(side="right", fill="both", expand=True, padx=(4, 0))
        ctk.CTkLabel(fp, text="📁 Files", font=ctk.CTkFont(size=11, weight="bold")).pack(anchor="w", padx=6, pady=6)
        self.files_text = ctk.CTkTextbox(fp, font=ctk.CTkFont(family="Consolas", size=9))
        self.files_text.pack(fill="both", expand=True, padx=6, pady=(0, 6))

    def _browse(self):
        p = filedialog.askdirectory(title="Select Python Directory")
        if p: self.dir_var.set(p)

    def _get_exclude_dirs(self):
        """Build list of directories to exclude based on options."""
        exclude = []
        if self.exclude_venv_var.get():
            exclude.extend(['venv', '.venv', 'env', '.env'])
        if self.exclude_pycache_var.get():
            exclude.append('__pycache__')
        if self.exclude_git_var.get():
            exclude.extend(['.git', '.idea', '.vscode', 'node_modules'])
        return exclude

    def _get_exclude_patterns(self):
        """Build list of file patterns to exclude based on options."""
        patterns = []
        if self.exclude_tests_var.get():
            patterns.extend(['test_*.py', '*_test.py', 'tests.py'])
        if self.exclude_init_var.get():
            patterns.append('__init__.py')
        if self.exclude_setup_var.get():
            patterns.extend(['setup.py', 'conftest.py'])
        return patterns

    def _analyze(self):
        d = self.dir_var.get()
        if not d:
            messagebox.showerror("Error", "Select a directory first")
            return
        exclude_dirs = self._get_exclude_dirs()
        def do():
            a = self.analyzer.analyze_directory(d, include_subdirs=self.subdirs_var.get(), exclude_dirs=exclude_dirs)
            s = self.analyzer.generate_summary(a)
            return a, s
        def done(r):
            a, s = r
            self._last_result = {"summary": s, "files": len(a)}
            self.stats_text.delete("1.0", "end")
            self.stats_text.insert("1.0", f"Files: {s['total_files']}\nLines: {s['total_lines']:,}\nCode: {s['total_code_lines']:,}\nClasses: {s['total_classes']}\nFunctions: {s['total_functions']}\nDoc ratio: {s['documentation_ratio']:.1f}%")
            self.files_text.delete("1.0", "end")
            for x in a[:50]:
                self.files_text.insert("end", f"{x.name} - {x.line_count} lines\n")
            self.set_progress(1.0, f"Analyzed {len(a)} files")
        self.run_async(do, done)

    def _extract_code(self):
        """Extract all Python code to a text file."""
        d = self.dir_var.get()
        if not d:
            messagebox.showerror("Error", "Select a directory first")
            return

        # Ask for output file
        output_path = filedialog.asksaveasfilename(
            title="Save Code Extraction",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            initialfile=f"{Path(d).name}_code_extraction.txt"
        )
        if not output_path:
            return

        exclude_dirs = self._get_exclude_dirs()
        exclude_patterns = self._get_exclude_patterns()

        try:
            max_size = int(self.max_size_var.get())
        except ValueError:
            max_size = 500

        def do():
            return self.analyzer.extract_code_hierarchy(
                d,
                include_subdirs=self.subdirs_var.get(),
                exclude_patterns=exclude_patterns,
                exclude_dirs=exclude_dirs,
                include_content=self.include_content_var.get(),
                max_file_size_kb=max_size
            )

        def done(content):
            try:
                with open(output_path, 'w', encoding='utf-8') as f:
                    f.write(content)
                self._last_result = {"output_path": output_path, "size": len(content)}
                self.set_progress(1.0, f"Extracted to {Path(output_path).name}")
                messagebox.showinfo("Success", f"Code extracted to:\n{output_path}")
                # Open file if option enabled
                if self.config.config.export.open_after_export:
                    webbrowser.open(output_path)
            except Exception as e:
                self.set_progress(0, "Failed")
                messagebox.showerror("Error", f"Failed to save: {e}")

        self.run_async(do, done)


class FolderScannerFrame(BaseToolFrame):
    def __init__(self, parent, **kw):
        self.scanner = FolderScanner()
        super().__init__(parent, "folder_scanner", "Folder Scanner", **kw)

    def _create_content(self):
        self.content_frame = ctk.CTkFrame(self)
        self.content_frame.pack(fill="both", expand=True, padx=10, pady=3)

        # Top section - Directory selection
        top = ctk.CTkFrame(self.content_frame)
        top.pack(fill="x", padx=4, pady=4)

        dir_frame = ctk.CTkFrame(top, fg_color="transparent")
        dir_frame.pack(fill="x", pady=4)
        ctk.CTkLabel(dir_frame, text="📁 Directory", font=ctk.CTkFont(size=11, weight="bold")).pack(side="left", padx=4)
        self.dir_var = ctk.StringVar()
        ctk.CTkEntry(dir_frame, textvariable=self.dir_var, placeholder_text="Select directory...", height=28).pack(side="left", fill="x", expand=True, padx=4)
        ctk.CTkButton(dir_frame, text="...", width=35, height=28, command=self._browse).pack(side="left", padx=(0, 4))

        # Action buttons frame
        btn_frame = ctk.CTkFrame(dir_frame, fg_color="transparent")
        btn_frame.pack(side="right")
        ctk.CTkButton(btn_frame, text="▶ Scan", command=self._scan, height=28, width=70,
                      font=ctk.CTkFont(size=10, weight="bold"),
                      fg_color=("#10b981", "#059669"), hover_color=("#059669", "#047857")).pack(side="left", padx=2)
        ctk.CTkButton(btn_frame, text="📄 TXT", command=self._export_txt, height=28, width=60,
                      font=ctk.CTkFont(size=10),
                      fg_color=("#6b7280", "#4b5563"), hover_color=("#4b5563", "#374151")).pack(side="left", padx=2)
        ctk.CTkButton(btn_frame, text="📊 Excel", command=self._export_excel, height=28, width=70,
                      font=ctk.CTkFont(size=10, weight="bold"),
                      fg_color=("#22c55e", "#16a34a"), hover_color=("#16a34a", "#15803d")).pack(side="left", padx=2)
        ctk.CTkButton(btn_frame, text="🏗️ Archi", command=self._export_architecture, height=28, width=70,
                      font=ctk.CTkFont(size=10, weight="bold"),
                      fg_color=("#8b5cf6", "#7c3aed"), hover_color=("#7c3aed", "#6d28d9")).pack(side="left", padx=2)

        # Options scrollable
        opts_scroll = ctk.CTkScrollableFrame(top, height=120, label_text="⚙️ Options de scan",
                                              label_font=ctk.CTkFont(size=10, weight="bold"),
                                              fg_color=("gray88", "gray20"), corner_radius=6)
        opts_scroll.pack(fill="x", padx=4, pady=4)

        opts_grid = ctk.CTkFrame(opts_scroll, fg_color="transparent")
        opts_grid.pack(fill="x")

        cfg = self.config.config.folder_scanner

        # Column 1 - Content options
        col1 = ctk.CTkFrame(opts_grid, fg_color=("gray95", "gray25"), corner_radius=6)
        col1.pack(side="left", fill="both", expand=True, padx=2, pady=2)
        ctk.CTkLabel(col1, text="Contenu", font=ctk.CTkFont(size=9, weight="bold"),
                     text_color=("#059669", "#10b981")).pack(anchor="w", padx=6, pady=(4, 2))
        self.content_var = ctk.BooleanVar(value=cfg.include_content)
        ctk.CTkCheckBox(col1, text="Inclure contenu", variable=self.content_var,
                        font=ctk.CTkFont(size=9), height=20, checkbox_width=16, checkbox_height=16).pack(anchor="w", padx=8, pady=1)
        self.include_binary_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(col1, text="Inclure binaires", variable=self.include_binary_var,
                        font=ctk.CTkFont(size=9), height=20, checkbox_width=16, checkbox_height=16).pack(anchor="w", padx=8, pady=1)
        self.show_hidden_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(col1, text="Fichiers cachés", variable=self.show_hidden_var,
                        font=ctk.CTkFont(size=9), height=20, checkbox_width=16, checkbox_height=16).pack(anchor="w", padx=8, pady=(1, 4))

        # Column 2 - Size limits
        col2 = ctk.CTkFrame(opts_grid, fg_color=("gray95", "gray25"), corner_radius=6)
        col2.pack(side="left", fill="both", expand=True, padx=2, pady=2)
        ctk.CTkLabel(col2, text="Limites", font=ctk.CTkFont(size=9, weight="bold"),
                     text_color=("#f59e0b", "#fbbf24")).pack(anchor="w", padx=6, pady=(4, 2))
        max_row = ctk.CTkFrame(col2, fg_color="transparent")
        max_row.pack(fill="x", padx=8, pady=2)
        ctk.CTkLabel(max_row, text="Max KB:", font=ctk.CTkFont(size=9)).pack(side="left")
        self.max_entry = ctk.CTkEntry(max_row, width=50, height=22, font=ctk.CTkFont(size=9))
        self.max_entry.insert(0, str(cfg.max_file_size_kb))
        self.max_entry.pack(side="left", padx=4)
        depth_row = ctk.CTkFrame(col2, fg_color="transparent")
        depth_row.pack(fill="x", padx=8, pady=2)
        ctk.CTkLabel(depth_row, text="Profondeur:", font=ctk.CTkFont(size=9)).pack(side="left")
        self.depth_var = ctk.StringVar(value="∞")
        ctk.CTkOptionMenu(depth_row, values=["1", "2", "3", "5", "10", "∞"], variable=self.depth_var,
                          width=50, height=22, font=ctk.CTkFont(size=9)).pack(side="left", padx=4)

        # Column 3 - Exclusions
        col3 = ctk.CTkFrame(opts_grid, fg_color=("gray95", "gray25"), corner_radius=6)
        col3.pack(side="left", fill="both", expand=True, padx=2, pady=2)
        ctk.CTkLabel(col3, text="Exclusions", font=ctk.CTkFont(size=9, weight="bold"),
                     text_color=("#059669", "#10b981")).pack(anchor="w", padx=6, pady=(4, 2))
        self.exclude_git_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(col3, text=".git / .svn", variable=self.exclude_git_var,
                        font=ctk.CTkFont(size=9), height=20, checkbox_width=16, checkbox_height=16).pack(anchor="w", padx=8, pady=1)
        self.exclude_pycache_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(col3, text="__pycache__", variable=self.exclude_pycache_var,
                        font=ctk.CTkFont(size=9), height=20, checkbox_width=16, checkbox_height=16).pack(anchor="w", padx=8, pady=1)
        self.exclude_node_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(col3, text="node_modules", variable=self.exclude_node_var,
                        font=ctk.CTkFont(size=9), height=20, checkbox_width=16, checkbox_height=16).pack(anchor="w", padx=8, pady=(1, 4))

        # Column 4 - Output format
        col4 = ctk.CTkFrame(opts_grid, fg_color=("gray95", "gray25"), corner_radius=6)
        col4.pack(side="left", fill="both", expand=True, padx=2, pady=2)
        ctk.CTkLabel(col4, text="Affichage", font=ctk.CTkFont(size=9, weight="bold"),
                     text_color=("#059669", "#10b981")).pack(anchor="w", padx=6, pady=(4, 2))
        self.show_sizes_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(col4, text="Tailles fichiers", variable=self.show_sizes_var,
                        font=ctk.CTkFont(size=9), height=20, checkbox_width=16, checkbox_height=16).pack(anchor="w", padx=8, pady=1)
        self.show_dates_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(col4, text="Dates modif.", variable=self.show_dates_var,
                        font=ctk.CTkFont(size=9), height=20, checkbox_width=16, checkbox_height=16).pack(anchor="w", padx=8, pady=1)
        self.tree_style_var = ctk.StringVar(value="tree")
        ctk.CTkOptionMenu(col4, values=["tree", "flat", "json"], variable=self.tree_style_var,
                          width=60, height=20, font=ctk.CTkFont(size=9)).pack(anchor="w", padx=8, pady=(1, 4))

        # Statistics section
        stats_frame = ctk.CTkFrame(self.content_frame, height=55)
        stats_frame.pack(fill="x", padx=4, pady=4)
        stats_frame.pack_propagate(False)

        self.slabels = {}
        stats_data = [
            ("files", "📄 Files", "#10b981"),
            ("dirs", "📁 Dirs", "#3b82f6"),
            ("size", "💾 Size", "#f59e0b"),
            ("time", "⏱️ Time", "#8b5cf6")
        ]
        for k, label, color in stats_data:
            f = ctk.CTkFrame(stats_frame, fg_color=("gray95", "gray25"), corner_radius=6)
            f.pack(side="left", fill="both", expand=True, padx=4, pady=4)
            ctk.CTkLabel(f, text=label, font=ctk.CTkFont(size=9), text_color=(color, color)).pack(pady=(4, 0))
            self.slabels[k] = ctk.CTkLabel(f, text="--", font=ctk.CTkFont(size=14, weight="bold"))
            self.slabels[k].pack(pady=(0, 4))

        # Tree output
        tree_frame = ctk.CTkFrame(self.content_frame)
        tree_frame.pack(fill="both", expand=True, padx=4, pady=(0, 4))
        ctk.CTkLabel(tree_frame, text="🌳 Directory Tree", font=ctk.CTkFont(size=11, weight="bold")).pack(anchor="w", padx=6, pady=4)
        self.tree_text = ctk.CTkTextbox(tree_frame, font=ctk.CTkFont(family="Consolas", size=9))
        self.tree_text.pack(fill="both", expand=True, padx=6, pady=(0, 6))

    def _browse(self):
        p = filedialog.askdirectory(title="Select Directory")
        if p: self.dir_var.set(p)

    def _fmt_size(self, s):
        for u in ['B', 'KB', 'MB', 'GB']:
            if s < 1024: return f"{s:.1f} {u}" if u != 'B' else f"{s} {u}"
            s /= 1024
        return f"{s:.1f} TB"

    def _scan(self):
        d = self.dir_var.get()
        if not d:
            messagebox.showerror("Error", "Select a directory first")
            return
        try:
            mx = int(self.max_entry.get()) * 1024
        except:
            mx = 1024 * 1024
        def do():
            self.scanner.configure(max_file_size=mx, include_content=self.content_var.get())
            return self.scanner.scan(d)
        def done(r):
            self._scan_result = r  # Store for export
            self._last_result = {"files": r.total_files, "dirs": r.total_directories, "size": r.total_size}
            self.slabels["files"].configure(text=str(r.total_files))
            self.slabels["dirs"].configure(text=str(r.total_directories))
            self.slabels["size"].configure(text=self._fmt_size(r.total_size))
            self.slabels["time"].configure(text=f"{r.scan_time:.1f}s")
            self.tree_text.delete("1.0", "end")
            self.tree_text.insert("1.0", self.scanner.generate_tree(r))
            self.set_progress(1.0, f"Scanned {r.total_files} files")
        self.run_async(do, done)

    def _export_txt(self):
        """Export scan result to a TXT file."""
        if not hasattr(self, '_scan_result') or not self._scan_result:
            messagebox.showwarning("Warning", "Please scan a directory first")
            return

        d = self.dir_var.get()
        output_path = filedialog.asksaveasfilename(
            title="Export to Text File",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            initialfile=f"{Path(d).name}_scan.txt"
        )
        if not output_path:
            return

        def do():
            self.scanner.export_to_file(self._scan_result, output_path, include_content=self.content_var.get())
            return output_path

        def done(path):
            self.set_progress(1.0, f"Exported to {Path(path).name}")
            messagebox.showinfo("Success", f"Scan exported to:\n{path}")
            if self.config.config.export.open_after_export:
                webbrowser.open(path)

        self.run_async(do, done)

    def _export_excel(self):
        """Export scan result to an Excel file."""
        if not hasattr(self, '_scan_result') or not self._scan_result:
            messagebox.showwarning("Warning", "Please scan a directory first")
            return

        d = self.dir_var.get()
        output_path = filedialog.asksaveasfilename(
            title="Export to Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")],
            initialfile=f"{Path(d).name}_scan.xlsx"
        )
        if not output_path:
            return

        def do():
            success = self.scanner.export_to_excel(self._scan_result, output_path)
            return success, output_path

        def done(result):
            success, path = result
            if success:
                self.set_progress(1.0, f"Exported to {Path(path).name}")
                messagebox.showinfo("Success", f"Scan exported to:\n{path}")
                if self.config.config.export.open_after_export:
                    webbrowser.open(path)
            else:
                self.set_progress(0, "Export failed")
                messagebox.showerror("Error", "Failed to export scan results")

        self.run_async(do, done)

    def _export_architecture(self):
        """Export full architecture with table of contents and file contents."""
        if not hasattr(self, '_scan_result') or not self._scan_result:
            messagebox.showwarning("Warning", "Please scan a directory first")
            return

        d = self.dir_var.get()
        output_path = filedialog.asksaveasfilename(
            title="Export Full Architecture",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            initialfile=f"{Path(d).name}_architecture.txt"
        )
        if not output_path:
            return

        # Ask for extension filter
        ext_filter = None
        if messagebox.askyesno("Filter Extensions", "Do you want to filter by specific extensions?\n(e.g., only .py, .js files)"):
            from tkinter import simpledialog
            ext_input = simpledialog.askstring("Extensions", "Enter extensions separated by commas\n(e.g.: .py, .js, .html)", parent=self)
            if ext_input:
                ext_filter = [e.strip() for e in ext_input.split(',') if e.strip()]

        def do():
            success = self.scanner.export_full_architecture(
                self._scan_result,
                output_path,
                extensions_filter=ext_filter,
                include_line_numbers=True
            )
            return success, output_path

        def done(result):
            success, path = result
            if success:
                self.set_progress(1.0, f"Architecture exported to {Path(path).name}")
                messagebox.showinfo("Success", f"Full architecture exported to:\n{path}")
                if self.config.config.export.open_after_export:
                    webbrowser.open(path)
            else:
                self.set_progress(0, "Export failed")
                messagebox.showerror("Error", "Failed to export architecture")

        self.run_async(do, done)


class VBAOptimizerFrame(BaseToolFrame):
    def __init__(self, parent, **kw):
        self.optimizer = VBAOptimizer()
        super().__init__(parent, "vba_optimizer", "VBA Optimizer", **kw)

    def _create_content(self):
        self.content_frame = ctk.CTkFrame(self)
        self.content_frame.pack(fill="both", expand=True, padx=10, pady=3)

        # Options scrollable frame at top
        opts_scroll = ctk.CTkScrollableFrame(self.content_frame, height=110, label_text="⚙️ Options d'optimisation",
                                              label_font=ctk.CTkFont(size=10, weight="bold"),
                                              fg_color=("gray88", "gray20"), corner_radius=6)
        opts_scroll.pack(fill="x", padx=4, pady=4)

        opts_grid = ctk.CTkFrame(opts_scroll, fg_color="transparent")
        opts_grid.pack(fill="x")

        cfg = self.config.config.vba_optimizer
        self.opts = {}

        # Column 1 - Basic cleanup
        col1 = ctk.CTkFrame(opts_grid, fg_color=("gray95", "gray25"), corner_radius=6)
        col1.pack(side="left", fill="both", expand=True, padx=2, pady=2)
        ctk.CTkLabel(col1, text="Nettoyage", font=ctk.CTkFont(size=9, weight="bold"),
                     text_color=("#059669", "#10b981")).pack(anchor="w", padx=6, pady=(4, 2))
        self.opts["remove_comments"] = ctk.BooleanVar(value=cfg.remove_comments)
        ctk.CTkCheckBox(col1, text="Supprimer commentaires", variable=self.opts["remove_comments"],
                        font=ctk.CTkFont(size=9), height=20, checkbox_width=16, checkbox_height=16).pack(anchor="w", padx=8, pady=1)
        self.opts["remove_empty_lines"] = ctk.BooleanVar(value=cfg.remove_empty_lines)
        ctk.CTkCheckBox(col1, text="Supprimer lignes vides", variable=self.opts["remove_empty_lines"],
                        font=ctk.CTkFont(size=9), height=20, checkbox_width=16, checkbox_height=16).pack(anchor="w", padx=8, pady=1)
        self.remove_debug_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(col1, text="Supprimer Debug.*", variable=self.remove_debug_var,
                        font=ctk.CTkFont(size=9), height=20, checkbox_width=16, checkbox_height=16).pack(anchor="w", padx=8, pady=(1, 4))

        # Column 2 - Formatting
        col2 = ctk.CTkFrame(opts_grid, fg_color=("gray95", "gray25"), corner_radius=6)
        col2.pack(side="left", fill="both", expand=True, padx=2, pady=2)
        ctk.CTkLabel(col2, text="Formatage", font=ctk.CTkFont(size=9, weight="bold"),
                     text_color=("#3b82f6", "#60a5fa")).pack(anchor="w", padx=6, pady=(4, 2))
        self.opts["auto_indent"] = ctk.BooleanVar(value=cfg.auto_indent)
        ctk.CTkCheckBox(col2, text="Auto-indentation", variable=self.opts["auto_indent"],
                        font=ctk.CTkFont(size=9), height=20, checkbox_width=16, checkbox_height=16).pack(anchor="w", padx=8, pady=1)
        indent_row = ctk.CTkFrame(col2, fg_color="transparent")
        indent_row.pack(fill="x", padx=8, pady=2)
        ctk.CTkLabel(indent_row, text="Indent:", font=ctk.CTkFont(size=9)).pack(side="left")
        self.indent_entry = ctk.CTkEntry(indent_row, width=35, height=20, font=ctk.CTkFont(size=9))
        self.indent_entry.insert(0, str(cfg.indent_size))
        self.indent_entry.pack(side="left", padx=4)
        self.normalize_case_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(col2, text="Normaliser casse", variable=self.normalize_case_var,
                        font=ctk.CTkFont(size=9), height=20, checkbox_width=16, checkbox_height=16).pack(anchor="w", padx=8, pady=(1, 4))

        # Column 3 - Advanced
        col3 = ctk.CTkFrame(opts_grid, fg_color=("gray95", "gray25"), corner_radius=6)
        col3.pack(side="left", fill="both", expand=True, padx=2, pady=2)
        ctk.CTkLabel(col3, text="Avancé", font=ctk.CTkFont(size=9, weight="bold"),
                     text_color=("#f59e0b", "#fbbf24")).pack(anchor="w", padx=6, pady=(4, 2))
        self.opts["minify"] = ctk.BooleanVar(value=cfg.minify)
        ctk.CTkCheckBox(col3, text="Minifier", variable=self.opts["minify"],
                        font=ctk.CTkFont(size=9), height=20, checkbox_width=16, checkbox_height=16).pack(anchor="w", padx=8, pady=1)
        self.rename_vars_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(col3, text="Renommer variables", variable=self.rename_vars_var,
                        font=ctk.CTkFont(size=9), height=20, checkbox_width=16, checkbox_height=16).pack(anchor="w", padx=8, pady=1)
        self.obfuscate_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(col3, text="Obfusquer (beta)", variable=self.obfuscate_var,
                        font=ctk.CTkFont(size=9), height=20, checkbox_width=16, checkbox_height=16).pack(anchor="w", padx=8, pady=(1, 4))

        # Column 4 - Safety
        col4 = ctk.CTkFrame(opts_grid, fg_color=("gray95", "gray25"), corner_radius=6)
        col4.pack(side="left", fill="both", expand=True, padx=2, pady=2)
        ctk.CTkLabel(col4, text="Sécurité", font=ctk.CTkFont(size=9, weight="bold"),
                     text_color=("#ef4444", "#f87171")).pack(anchor="w", padx=6, pady=(4, 2))
        self.create_backup_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(col4, text="Créer backup", variable=self.create_backup_var,
                        font=ctk.CTkFont(size=9), height=20, checkbox_width=16, checkbox_height=16).pack(anchor="w", padx=8, pady=1)
        self.validate_syntax_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(col4, text="Valider syntaxe", variable=self.validate_syntax_var,
                        font=ctk.CTkFont(size=9), height=20, checkbox_width=16, checkbox_height=16).pack(anchor="w", padx=8, pady=1)
        self.preview_only_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(col4, text="Aperçu seul", variable=self.preview_only_var,
                        font=ctk.CTkFont(size=9), height=20, checkbox_width=16, checkbox_height=16).pack(anchor="w", padx=8, pady=(1, 4))

        # Action button
        action_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        action_frame.pack(fill="x", padx=4, pady=4)
        ctk.CTkButton(action_frame, text="▶ Optimize", command=self._optimize, height=32, width=120,
                      font=ctk.CTkFont(size=12, weight="bold"),
                      fg_color=("#10b981", "#059669"), hover_color=("#059669", "#047857")).pack(side="left", padx=4)
        self.stats_label = ctk.CTkLabel(action_frame, text="", font=ctk.CTkFont(size=10), text_color=("gray50", "gray60"))
        self.stats_label.pack(side="left", padx=10)

        # Code panels
        code = ctk.CTkFrame(self.content_frame)
        code.pack(fill="both", expand=True, padx=4, pady=(0, 4))

        # Input panel
        inp = ctk.CTkFrame(code)
        inp.pack(side="left", fill="both", expand=True, padx=(0, 2))
        ih = ctk.CTkFrame(inp, fg_color="transparent")
        ih.pack(fill="x", padx=6, pady=4)
        ctk.CTkLabel(ih, text="📥 Input (Avant)", font=ctk.CTkFont(size=11, weight="bold")).pack(side="left")
        ctk.CTkButton(ih, text="Load", width=50, height=24, font=ctk.CTkFont(size=9), command=self._load).pack(side="right", padx=2)
        ctk.CTkButton(ih, text="Paste", width=50, height=24, font=ctk.CTkFont(size=9), command=self._paste).pack(side="right", padx=2)
        ctk.CTkButton(ih, text="Clear", width=50, height=24, font=ctk.CTkFont(size=9), command=self._clear_input).pack(side="right", padx=2)
        ctk.CTkButton(ih, text="Example", width=60, height=24, font=ctk.CTkFont(size=9),
                      fg_color=("#8b5cf6", "#7c3aed"), hover_color=("#7c3aed", "#6d28d9"),
                      command=self._load_example).pack(side="right", padx=2)
        self.input_text = ctk.CTkTextbox(inp, font=ctk.CTkFont(family="Consolas", size=9))
        self.input_text.pack(fill="both", expand=True, padx=6, pady=(0, 6))

        # Output panel
        out = ctk.CTkFrame(code)
        out.pack(side="right", fill="both", expand=True, padx=(2, 0))
        oh = ctk.CTkFrame(out, fg_color="transparent")
        oh.pack(fill="x", padx=6, pady=4)
        ctk.CTkLabel(oh, text="📤 Output (Après)", font=ctk.CTkFont(size=11, weight="bold")).pack(side="left")
        ctk.CTkButton(oh, text="Save", width=50, height=24, font=ctk.CTkFont(size=9), command=self._save).pack(side="right", padx=2)
        ctk.CTkButton(oh, text="Copy", width=50, height=24, font=ctk.CTkFont(size=9), command=self._copy).pack(side="right", padx=2)
        self.output_text = ctk.CTkTextbox(out, font=ctk.CTkFont(family="Consolas", size=9))
        self.output_text.pack(fill="both", expand=True, padx=6, pady=(0, 6))

    def _load_example(self):
        """Load an example VBA code to demonstrate optimization."""
        example_code = '''' Module: ExampleModule
' This is a demonstration VBA module with various issues
' Author: Demo
' Date: 2024-01-01

Option Explicit

' Global variable declaration
Dim gCounter As Integer   ' Counter variable
Dim unusedGlobal As String   ' This variable is never used


Sub MainProcedure()
' This is the main procedure
Dim i As Integer
Dim j As Integer
Dim result As Double



' Initialize variables
i = 0
j = 10


' Loop through values
For i = 1 To j
    ' Increment counter
    gCounter = gCounter + 1
    Debug.Print "Counter: " & gCounter

    If i Mod 2 = 0 Then
        ' Even number
        result = i * 2
    Else
    ' Odd number
    result = i * 3
    End If

Next i


' Display result
MsgBox "Final result: " & result
End Sub


Function CalculateSum(ByVal a As Integer, ByVal b As Integer) As Integer
' Function to calculate sum of two numbers
' Parameters: a - first number, b - second number
' Returns: sum of a and b

    Dim sum As Integer    ' Local sum variable


    sum = a + b


    CalculateSum = sum
End Function
'''
        self.input_text.delete("1.0", "end")
        self.input_text.insert("1.0", example_code)
        self.stats_label.configure(text="Example loaded - Click 'Optimize' to see the result")

    def _paste(self):
        try:
            content = self.clipboard_get()
            self.input_text.delete("1.0", "end")
            self.input_text.insert("1.0", content)
        except: pass

    def _clear_input(self):
        self.input_text.delete("1.0", "end")

    def _copy(self):
        content = self.output_text.get("1.0", "end-1c")
        if content:
            self.clipboard_clear()
            self.clipboard_append(content)

    def _load(self):
        p = filedialog.askopenfilename(title="Load VBA", filetypes=[("VBA", "*.bas;*.cls;*.frm"), ("All", "*.*")])
        if p:
            with open(p, 'r', encoding='utf-8', errors='replace') as f:
                self.input_text.delete("1.0", "end")
                self.input_text.insert("1.0", f.read())

    def _save(self):
        c = self.output_text.get("1.0", "end-1c")
        if not c.strip():
            messagebox.showwarning("Warning", "No code to save")
            return
        p = filedialog.asksaveasfilename(title="Save VBA", defaultextension=".bas", filetypes=[("VBA Module", "*.bas"), ("Text", "*.txt")])
        if p:
            with open(p, 'w', encoding='utf-8') as f: f.write(c)
            messagebox.showinfo("Success", f"Saved to {p}")

    def _optimize(self):
        c = self.input_text.get("1.0", "end-1c")
        if not c.strip():
            messagebox.showwarning("Warning", "Enter code first")
            return
        try:
            ind = int(self.indent_entry.get())
        except:
            ind = 4
        o = OptimizationOptions(remove_comments=self.opts["remove_comments"].get(), auto_indent=self.opts["auto_indent"].get(), remove_empty_lines=self.opts["remove_empty_lines"].get(), minify=self.opts["minify"].get(), indent_size=ind)
        r = self.optimizer.optimize(c, o)
        self._last_result = {"orig_lines": r.original_lines, "opt_lines": r.optimized_lines, "changes": r.modifications}
        self.output_text.delete("1.0", "end")
        if r.success:
            self.output_text.insert("1.0", r.optimized_code)
            self.stats_label.configure(text=f"Lines: {r.original_lines} -> {r.optimized_lines} | Size: {r.original_size} -> {r.optimized_size}")
            self.set_progress(1.0, "Optimized")
        else:
            self.output_text.insert("1.0", f"Error: {r.error_message}")
            self.set_progress(0, "Failed")


class VBAAnalyzerFrame(BaseToolFrame):
    """Advanced VBA code analyzer with regex, graphs, and DataFrame export."""
    def __init__(self, parent, **kw):
        self.analyzer = VBAAnalyzer()
        self.analysis_results = []
        super().__init__(parent, "vba_analyzer", "VBA Analyzer", **kw)

    def _create_content(self):
        self.content_frame = ctk.CTkFrame(self)
        self.content_frame.pack(fill="both", expand=True, padx=10, pady=3)

        # Top section - Input options
        top = ctk.CTkFrame(self.content_frame)
        top.pack(fill="x", padx=4, pady=4)

        # Mode selection
        mode_frame = ctk.CTkFrame(top, fg_color="transparent")
        mode_frame.pack(fill="x", pady=4)
        ctk.CTkLabel(mode_frame, text="📊 VBA Analyzer", font=ctk.CTkFont(size=12, weight="bold")).pack(side="left", padx=4)

        self.mode_var = ctk.StringVar(value="code")
        ctk.CTkSegmentedButton(mode_frame, values=["Code VBA", "Fichier Hex"],
                               variable=self.mode_var, command=self._on_mode_change).pack(side="left", padx=10)

        # Action buttons
        btn_frame = ctk.CTkFrame(mode_frame, fg_color="transparent")
        btn_frame.pack(side="right")
        ctk.CTkButton(btn_frame, text="▶ Analyze", command=self._analyze, height=28, width=80,
                      font=ctk.CTkFont(size=10, weight="bold"),
                      fg_color=("#10b981", "#059669"), hover_color=("#059669", "#047857")).pack(side="left", padx=2)
        ctk.CTkButton(btn_frame, text="📈 Graphs", command=self._show_graphs, height=28, width=80,
                      font=ctk.CTkFont(size=10),
                      fg_color=("#8b5cf6", "#7c3aed"), hover_color=("#7c3aed", "#6d28d9")).pack(side="left", padx=2)
        ctk.CTkButton(btn_frame, text="📊 Excel", command=self._export_excel, height=28, width=70,
                      font=ctk.CTkFont(size=10),
                      fg_color=("#22c55e", "#16a34a"), hover_color=("#16a34a", "#15803d")).pack(side="left", padx=2)

        # Input frame (code or file)
        self.input_frame = ctk.CTkFrame(self.content_frame)
        self.input_frame.pack(fill="both", expand=True, padx=4, pady=4)

        self._create_code_input()

    def _create_code_input(self):
        """Create the code input interface."""
        for w in self.input_frame.winfo_children():
            w.destroy()

        # Left panel - Input
        left = ctk.CTkFrame(self.input_frame)
        left.pack(side="left", fill="both", expand=True, padx=(0, 2))

        ih = ctk.CTkFrame(left, fg_color="transparent")
        ih.pack(fill="x", padx=6, pady=4)
        ctk.CTkLabel(ih, text="📥 Code VBA", font=ctk.CTkFont(size=11, weight="bold")).pack(side="left")
        ctk.CTkButton(ih, text="Load", width=50, height=24, font=ctk.CTkFont(size=9), command=self._load_file).pack(side="right", padx=2)
        ctk.CTkButton(ih, text="Paste", width=50, height=24, font=ctk.CTkFont(size=9), command=self._paste).pack(side="right", padx=2)
        ctk.CTkButton(ih, text="Clear", width=50, height=24, font=ctk.CTkFont(size=9), command=self._clear).pack(side="right", padx=2)
        ctk.CTkButton(ih, text="Example", width=60, height=24, font=ctk.CTkFont(size=9),
                      fg_color=("#f59e0b", "#d97706"), command=self._load_example).pack(side="right", padx=2)

        self.input_text = ctk.CTkTextbox(left, font=ctk.CTkFont(family="Consolas", size=9))
        self.input_text.pack(fill="both", expand=True, padx=6, pady=(0, 6))

        # Right panel - Results
        right = ctk.CTkFrame(self.input_frame)
        right.pack(side="right", fill="both", expand=True, padx=(2, 0))

        rh = ctk.CTkFrame(right, fg_color="transparent")
        rh.pack(fill="x", padx=6, pady=4)
        ctk.CTkLabel(rh, text="📤 Analyse", font=ctk.CTkFont(size=11, weight="bold")).pack(side="left")
        ctk.CTkButton(rh, text="Copy", width=50, height=24, font=ctk.CTkFont(size=9), command=self._copy_results).pack(side="right", padx=2)
        ctk.CTkButton(rh, text="Save", width=50, height=24, font=ctk.CTkFont(size=9), command=self._save_results).pack(side="right", padx=2)

        self.result_text = ctk.CTkTextbox(right, font=ctk.CTkFont(family="Consolas", size=9))
        self.result_text.pack(fill="both", expand=True, padx=6, pady=(0, 6))

    def _create_hex_input(self):
        """Create the hex preview interface."""
        for w in self.input_frame.winfo_children():
            w.destroy()

        # File selection
        file_frame = ctk.CTkFrame(self.input_frame, fg_color="transparent")
        file_frame.pack(fill="x", padx=6, pady=6)
        ctk.CTkLabel(file_frame, text="📁 Binary File", font=ctk.CTkFont(size=11, weight="bold")).pack(side="left", padx=4)
        self.hex_file_var = ctk.StringVar()
        ctk.CTkEntry(file_frame, textvariable=self.hex_file_var, placeholder_text="Select binary file...", height=28).pack(side="left", fill="x", expand=True, padx=4)
        ctk.CTkButton(file_frame, text="...", width=35, height=28, command=self._browse_hex_file).pack(side="left", padx=(0, 4))
        ctk.CTkButton(file_frame, text="Preview", width=70, height=28,
                      fg_color=("#10b981", "#059669"), command=self._preview_hex).pack(side="left", padx=4)

        # Options
        opts_frame = ctk.CTkFrame(self.input_frame, fg_color="transparent")
        opts_frame.pack(fill="x", padx=6, pady=4)
        ctk.CTkLabel(opts_frame, text="Max bytes:", font=ctk.CTkFont(size=9)).pack(side="left", padx=4)
        self.hex_bytes_var = ctk.StringVar(value="512")
        ctk.CTkOptionMenu(opts_frame, values=["256", "512", "1024", "2048", "4096"],
                          variable=self.hex_bytes_var, width=80, height=24).pack(side="left", padx=4)

        # Hex output
        self.hex_text = ctk.CTkTextbox(self.input_frame, font=ctk.CTkFont(family="Consolas", size=9))
        self.hex_text.pack(fill="both", expand=True, padx=6, pady=(0, 6))

    def _on_mode_change(self, mode):
        if mode == "Code VBA":
            self._create_code_input()
        else:
            self._create_hex_input()

    def _load_file(self):
        p = filedialog.askopenfilename(title="Load VBA", filetypes=[("VBA", "*.bas;*.cls;*.frm;*.txt"), ("All", "*.*")])
        if p:
            try:
                with open(p, 'r', encoding='utf-8', errors='replace') as f:
                    self.input_text.delete("1.0", "end")
                    self.input_text.insert("1.0", f.read())
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load: {e}")

    def _paste(self):
        try:
            content = self.clipboard_get()
            self.input_text.delete("1.0", "end")
            self.input_text.insert("1.0", content)
        except:
            pass

    def _clear(self):
        self.input_text.delete("1.0", "end")
        self.result_text.delete("1.0", "end")
        self.analysis_results = []

    def _load_example(self):
        example = '''Option Explicit

' Module-level constants
Public Const APP_NAME As String = "MyApplication"
Private Const MAX_ITEMS As Integer = 100

' Module-level variables
Private mCounter As Long
Public gUserName As String
Dim mConfig As Object

' Main initialization procedure
Public Sub Initialize()
    Dim tempValue As Integer
    Dim i As Long
    Static callCount As Integer

    callCount = callCount + 1
    mCounter = 0

    For i = 1 To MAX_ITEMS
        mCounter = mCounter + 1
        Debug.Print "Item: " & i
    Next i
End Sub

' Function to calculate sum
Private Function CalculateSum(ByVal a As Double, ByVal b As Double) As Double
    Dim result As Double
    result = a + b
    CalculateSum = result
End Function

' Property Get example
Public Property Get Counter() As Long
    Counter = mCounter
End Property

' Property Let example
Public Property Let Counter(ByVal value As Long)
    mCounter = value
End Property

' Class initialization
Private Sub Class_Initialize()
    Dim msg As String
    msg = "Class initialized"
    Debug.Print msg
End Sub
'''
        self.input_text.delete("1.0", "end")
        self.input_text.insert("1.0", example)
        self.set_progress(0.1, "Example loaded")

    def _analyze(self):
        if self.mode_var.get() == "Fichier Hex":
            self._preview_hex()
            return

        code = self.input_text.get("1.0", "end-1c")
        if not code.strip():
            messagebox.showwarning("Warning", "Enter VBA code first")
            return

        def do():
            return self.analyzer.analyze_code(code, "Module1", "")

        def done(result):
            self.analysis_results = [result]
            self._last_result = result.to_dict()

            self.result_text.delete("1.0", "end")
            if result.success:
                # Display summary
                lines = []
                lines.append("=" * 50)
                lines.append(" ANALYSE VBA")
                lines.append("=" * 50)
                lines.append("")
                lines.append(f"Procedures: {result.total_procedures}")
                lines.append(f"Variables: {result.total_variables}")
                lines.append(f"Constants: {result.total_constants}")
                lines.append("")

                # Procedures
                if result.procedures:
                    lines.append("PROCEDURES:")
                    lines.append("-" * 40)
                    for proc in result.procedures:
                        ret = f" As {proc.return_type}" if proc.return_type else ""
                        lines.append(f"  {proc.scope} {proc.procedure_type} {proc.name}({proc.parameters}){ret}")
                    lines.append("")

                # Variables
                if result.variables:
                    lines.append("VARIABLES & CONSTANTS:")
                    lines.append("-" * 40)
                    for var in result.variables:
                        scope_info = f"[{var.procedure_name}]" if var.procedure_name else "[Module]"
                        value_info = f" = {var.value}" if var.value else ""
                        lines.append(f"  {var.declaration} {var.name} As {var.var_type}{value_info} {scope_info}")

                self.result_text.insert("1.0", "\n".join(lines))
                self.set_progress(1.0, f"Found {result.total_procedures} procs, {result.total_variables} vars")
            else:
                self.result_text.insert("1.0", f"Error: {result.error_message}")
                self.set_progress(0, "Analysis failed")

        self.run_async(do, done)

    def _show_graphs(self):
        if not self.analysis_results:
            messagebox.showwarning("Warning", "Analyze code first")
            return

        try:
            import matplotlib.pyplot as plt
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        except ImportError:
            messagebox.showerror("Error", "matplotlib not installed.\nInstall with: pip install matplotlib")
            return

        # Create graph window
        graph_win = ctk.CTkToplevel(self)
        graph_win.title("VBA Analysis Graphs")
        graph_win.geometry("1000x700")
        graph_win.transient(self.winfo_toplevel())

        # Notebook for tabs
        notebook = ctk.CTkTabview(graph_win)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)

        stats = self.analyzer.generate_statistics(self.analysis_results)

        # Tab 1: Procedures by type
        tab1 = notebook.add("Procedures")
        fig1, ax1 = plt.subplots(figsize=(8, 5))
        if stats['procedures_by_type']:
            types = list(stats['procedures_by_type'].keys())
            counts = list(stats['procedures_by_type'].values())
            bars = ax1.bar(types, counts, color=['#10b981', '#3b82f6', '#f59e0b', '#ef4444'][:len(types)])
            ax1.set_title('Procedures by Type', fontsize=14, fontweight='bold')
            ax1.set_ylabel('Count')
            for bar, count in zip(bars, counts):
                ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.1,
                        str(count), ha='center', va='bottom')
        else:
            ax1.text(0.5, 0.5, 'No procedures found', ha='center', va='center')
        canvas1 = FigureCanvasTkAgg(fig1, tab1)
        canvas1.get_tk_widget().pack(fill="both", expand=True)

        # Tab 2: Variables by type
        tab2 = notebook.add("Variables")
        fig2, (ax2a, ax2b) = plt.subplots(1, 2, figsize=(12, 5))
        if stats['variables_by_type']:
            types = list(stats['variables_by_type'].keys())[:10]
            counts = [stats['variables_by_type'][t] for t in types]
            ax2a.barh(types, counts, color='lightcoral')
            ax2a.set_title('Top Variable Types', fontweight='bold')
            ax2a.set_xlabel('Count')
        if stats['variables_by_declaration']:
            labels = list(stats['variables_by_declaration'].keys())
            sizes = list(stats['variables_by_declaration'].values())
            ax2b.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=90)
            ax2b.set_title('Declaration Types', fontweight='bold')
        plt.tight_layout()
        canvas2 = FigureCanvasTkAgg(fig2, tab2)
        canvas2.get_tk_widget().pack(fill="both", expand=True)

        # Tab 3: Scope distribution
        tab3 = notebook.add("Scope")
        fig3, ax3 = plt.subplots(figsize=(8, 5))
        if stats['procedures_by_scope']:
            scopes = list(stats['procedures_by_scope'].keys())
            counts = list(stats['procedures_by_scope'].values())
            colors = ['#22c55e' if s == 'Public' else '#ef4444' if s == 'Private' else '#3b82f6' for s in scopes]
            ax3.bar(scopes, counts, color=colors)
            ax3.set_title('Procedures by Scope', fontsize=14, fontweight='bold')
            ax3.set_ylabel('Count')
        canvas3 = FigureCanvasTkAgg(fig3, tab3)
        canvas3.get_tk_widget().pack(fill="both", expand=True)

    def _export_excel(self):
        if not self.analysis_results:
            messagebox.showwarning("Warning", "Analyze code first")
            return

        try:
            import pandas as pd
        except ImportError:
            messagebox.showerror("Error", "pandas not installed.\nInstall with: pip install pandas openpyxl")
            return

        output_path = filedialog.asksaveasfilename(
            title="Export to Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")]
        )
        if not output_path:
            return

        def do():
            return self.analyzer.export_to_excel(self.analysis_results, output_path)

        def done(success):
            if success:
                self.set_progress(1.0, f"Exported to {Path(output_path).name}")
                messagebox.showinfo("Success", f"Exported to:\n{output_path}")
                if self.config.config.export.open_after_export:
                    webbrowser.open(output_path)
            else:
                messagebox.showerror("Error", "Export failed")

        self.run_async(do, done)

    def _copy_results(self):
        content = self.result_text.get("1.0", "end-1c")
        if content:
            self.clipboard_clear()
            self.clipboard_append(content)

    def _save_results(self):
        content = self.result_text.get("1.0", "end-1c")
        if not content.strip():
            messagebox.showwarning("Warning", "No results to save")
            return
        p = filedialog.asksaveasfilename(title="Save Results", defaultextension=".txt",
                                          filetypes=[("Text", "*.txt"), ("All", "*.*")])
        if p:
            with open(p, 'w', encoding='utf-8') as f:
                f.write(content)
            messagebox.showinfo("Success", f"Saved to {p}")

    def _browse_hex_file(self):
        p = filedialog.askopenfilename(title="Select Binary File", filetypes=[("All files", "*.*")])
        if p:
            self.hex_file_var.set(p)

    def _preview_hex(self):
        if not hasattr(self, 'hex_file_var'):
            return
        file_path = self.hex_file_var.get()
        if not file_path:
            messagebox.showwarning("Warning", "Select a file first")
            return

        try:
            max_bytes = int(self.hex_bytes_var.get())
        except:
            max_bytes = 512

        hex_output = get_hex_preview(file_path, max_bytes)
        self.hex_text.delete("1.0", "end")
        self.hex_text.insert("1.0", hex_output)
        self.set_progress(1.0, f"Hex preview: {Path(file_path).name}")


class SettingsFrame(ctk.CTkFrame):
    def __init__(self, parent, main_win, **kw):
        super().__init__(parent, **kw)
        self.main_win, self.config = main_win, get_config()
        self._create()

    def _create(self):
        scroll = ctk.CTkScrollableFrame(self)
        scroll.pack(fill="both", expand=True, padx=15, pady=15)
        # Appearance
        af = ctk.CTkFrame(scroll)
        af.pack(fill="x", pady=10)
        ctk.CTkLabel(af, text="Appearance", font=ctk.CTkFont(size=15, weight="bold")).pack(anchor="w", padx=12, pady=12)
        for k, l, v in [("theme", "Theme", ["dark", "light", "system"]), ("color_scheme", "Colors", ["blue", "green", "dark-blue"])]:
            r = ctk.CTkFrame(af, fg_color="transparent")
            r.pack(fill="x", padx=12, pady=4)
            ctk.CTkLabel(r, text=l, width=150).pack(side="left")
            var = ctk.StringVar(value=self.config.get(f"ui.{k}", v[0]))
            ctk.CTkOptionMenu(r, values=v, variable=var, width=120, command=lambda x, kk=k: self._opt(kk, x)).pack(side="right")
        # Export
        ef = ctk.CTkFrame(scroll)
        ef.pack(fill="x", pady=10)
        ctk.CTkLabel(ef, text="Export", font=ctk.CTkFont(size=15, weight="bold")).pack(anchor="w", padx=12, pady=12)
        for k, l, v in [("default_format", "Format", ["html", "json", "csv"])]:
            r = ctk.CTkFrame(ef, fg_color="transparent")
            r.pack(fill="x", padx=12, pady=4)
            ctk.CTkLabel(r, text=l, width=150).pack(side="left")
            var = ctk.StringVar(value=self.config.get(f"export.{k}", v[0]))
            ctk.CTkOptionMenu(r, values=v, variable=var, width=120, command=lambda x, kk=k: self._exp_opt(kk, x)).pack(side="right")
        for k, l in [("open_after_export", "Open after export")]:
            r = ctk.CTkFrame(ef, fg_color="transparent")
            r.pack(fill="x", padx=12, pady=4)
            ctk.CTkLabel(r, text=l, width=150).pack(side="left")
            var = ctk.BooleanVar(value=self.config.get(f"export.{k}", True))
            ctk.CTkCheckBox(r, text="", variable=var, command=lambda kk=k, vv=var: self._exp_chk(kk, vv)).pack(side="right")
        # Actions
        bf = ctk.CTkFrame(scroll)
        bf.pack(fill="x", pady=10)
        ctk.CTkLabel(bf, text="Configuration", font=ctk.CTkFont(size=15, weight="bold")).pack(anchor="w", padx=12, pady=12)
        bb = ctk.CTkFrame(bf, fg_color="transparent")
        bb.pack(fill="x", padx=12, pady=(0, 12))
        ctk.CTkButton(bb, text="Reset Defaults", width=120, command=self._reset).pack(side="left", padx=4)
        ctk.CTkButton(bb, text="Export Config", width=120, command=self._export).pack(side="left", padx=4)
        ctk.CTkButton(bb, text="Import Config", width=120, command=self._import).pack(side="left", padx=4)
        # About
        ab = ctk.CTkFrame(scroll)
        ab.pack(fill="x", pady=10)
        ctk.CTkLabel(ab, text="About", font=ctk.CTkFont(size=15, weight="bold")).pack(anchor="w", padx=12, pady=12)
        ctk.CTkLabel(ab, text="CodeExtractPro v1.0.0\nProfessional Code Extraction Suite\n\nLicense: MIT", font=ctk.CTkFont(size=10), justify="left").pack(anchor="w", padx=12, pady=(0, 12))

    def _opt(self, k, v):
        self.config.set(f"ui.{k}", v)
        if k == "theme": ctk.set_appearance_mode(v)

    def _exp_opt(self, k, v):
        self.config.set(f"export.{k}", v)

    def _exp_chk(self, k, var):
        self.config.set(f"export.{k}", var.get())

    def _reset(self):
        if messagebox.askyesno("Confirm", "Reset all settings?"):
            self.config.reset_to_defaults()
            messagebox.showinfo("Done", "Settings reset. Restart app.")

    def _export(self):
        p = filedialog.asksaveasfilename(title="Export Config", defaultextension=".json", filetypes=[("JSON", "*.json")])
        if p and self.config.export_config(p):
            messagebox.showinfo("Success", f"Exported to {p}")

    def _import(self):
        p = filedialog.askopenfilename(title="Import Config", filetypes=[("JSON", "*.json")])
        if p and self.config.import_config(p):
            messagebox.showinfo("Success", "Imported. Restart app.")


class LogsFrame(ctk.CTkFrame):
    def __init__(self, parent, **kw):
        super().__init__(parent, **kw)
        self.logger = get_logger()
        self._create()
        self._setup_logging()

    def _create(self):
        ctrl = ctk.CTkFrame(self)
        ctrl.pack(fill="x", padx=15, pady=15)
        ctk.CTkLabel(ctrl, text="Logs", font=ctk.CTkFont(size=16, weight="bold")).pack(side="left")
        ctk.CTkButton(ctrl, text="Clear", width=70, command=self._clear).pack(side="right", padx=4)
        ctk.CTkButton(ctrl, text="Export", width=70, command=self._export).pack(side="right", padx=4)
        self.log_text = ctk.CTkTextbox(self, font=ctk.CTkFont(family="Consolas", size=9))
        self.log_text.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        self.log_text._textbox.tag_configure("INFO", foreground="white")
        self.log_text._textbox.tag_configure("SUCCESS", foreground="#10b981")
        self.log_text._textbox.tag_configure("WARNING", foreground="#f59e0b")
        self.log_text._textbox.tag_configure("ERROR", foreground="#ef4444")

    def _setup_logging(self):
        def cb(e):
            self.after(0, lambda: self._add(e))
        self.logger.add_callback(cb)

    def _add(self, e):
        ts = e.timestamp.strftime("%H:%M:%S")
        self.log_text.insert("end", f"[{ts}] [{e.level.name}] {e.message}\n", e.level.name)
        self.log_text.see("end")

    def _clear(self):
        self.log_text.delete("1.0", "end")

    def _export(self):
        p = filedialog.asksaveasfilename(title="Export Logs", defaultextension=".txt", filetypes=[("Text", "*.txt")])
        if p:
            with open(p, 'w', encoding='utf-8') as f:
                f.write(self.log_text.get("1.0", "end-1c"))
            messagebox.showinfo("Success", f"Exported to {p}")


class MainWindow(ctk.CTk):
    """Main application window with independent tools."""

    def __init__(self):
        super().__init__()
        self.config = get_config()
        self.logger = get_logger()
        ui = self.config.config.ui
        ctk.set_appearance_mode(ui.theme)
        ctk.set_default_color_theme(ui.color_scheme)
        self.title("CodeExtractPro v1.0 - Professional Code Extraction Suite")

        # Set application icon (taskbar + window)
        self._set_icon()

        # Taille réduite pour une IHM plus compacte
        self.geometry(f"{min(ui.window_width, 1200)}x{min(ui.window_height, 750)}")
        self.minsize(900, 600)
        if ui.window_x and ui.window_y:
            self.geometry(f"+{ui.window_x}+{ui.window_y}")
        self._create_ui()
        self._shortcuts()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.logger.info("CodeExtractPro v1.0 started")

    def _set_icon(self):
        """Set application icon for window and taskbar."""
        import sys
        try:
            # Find icon path - try multiple locations
            if getattr(sys, 'frozen', False):
                # Running as PyInstaller bundle
                base_path = sys._MEIPASS
            else:
                # Running as script
                base_path = Path(__file__).parent.parent.parent

            # Try ico folder first, then assets folder
            icon_paths = [
                Path(base_path) / "ico" / "icone.ico",
                Path(base_path) / "assets" / "icon.ico",
            ]

            icon_path = None
            for p in icon_paths:
                if p.exists():
                    icon_path = p
                    break

            if icon_path and icon_path.exists():
                # Set window icon
                self.iconbitmap(str(icon_path))

                # Windows taskbar icon fix - set app id
                if sys.platform == "win32":
                    try:
                        import ctypes
                        # Set AppUserModelID for proper taskbar grouping and icon
                        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("CodeExtractPro.v1.0")
                    except Exception:
                        pass
        except Exception as e:
            # Silently ignore icon errors
            pass

    def _create_ui(self):
        main = ctk.CTkFrame(self, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=6, pady=6)
        # Header compact
        hdr = ctk.CTkFrame(main, height=40, fg_color=("gray90", "gray17"))
        hdr.pack(fill="x", pady=(0, 4))
        hdr.pack_propagate(False)
        tf = ctk.CTkFrame(hdr, fg_color="transparent")
        tf.pack(side="left", padx=10, pady=4)
        ctk.CTkLabel(tf, text="CodeExtractPro v1.0", font=ctk.CTkFont(size=16, weight="bold")).pack(side="left")
        ctk.CTkLabel(tf, text=" - Professional Code Extraction Suite", font=ctk.CTkFont(size=10), text_color=("gray50", "gray60")).pack(side="left", padx=(4, 0))
        thf = ctk.CTkFrame(hdr, fg_color="transparent")
        thf.pack(side="right", padx=10)
        self.theme_var = ctk.StringVar(value=self.config.config.ui.theme)
        ctk.CTkOptionMenu(thf, values=["dark", "light", "system"], variable=self.theme_var, width=80, height=26, command=self._theme).pack()
        # Tabs
        self.tabs = ctk.CTkTabview(main)
        self.tabs.pack(fill="both", expand=True, pady=(0, 4))
        t1 = self.tabs.add("VBA Extractor")
        t2 = self.tabs.add("Python Analyzer")
        t3 = self.tabs.add("Folder Scanner")
        t4 = self.tabs.add("VBA Optimizer")
        t5 = self.tabs.add("VBA Analyzer")
        t6 = self.tabs.add("Settings")
        t7 = self.tabs.add("Logs")
        VBAExtractorFrame(t1).pack(fill="both", expand=True)
        PythonAnalyzerFrame(t2).pack(fill="both", expand=True)
        FolderScannerFrame(t3).pack(fill="both", expand=True)
        VBAOptimizerFrame(t4).pack(fill="both", expand=True)
        VBAAnalyzerFrame(t5).pack(fill="both", expand=True)
        SettingsFrame(t6, self).pack(fill="both", expand=True)
        LogsFrame(t7).pack(fill="both", expand=True)
        # Footer compact
        ftr = ctk.CTkFrame(main, height=22, fg_color=("gray90", "gray17"))
        ftr.pack(fill="x")
        ftr.pack_propagate(False)
        ctk.CTkLabel(ftr, text="Ready", font=ctk.CTkFont(size=9), text_color=("gray50", "gray60")).pack(side="left", padx=6, pady=2)
        ctk.CTkLabel(ftr, text="F1=Help | Ctrl+1-5=Tabs | Ctrl+Q=Quit", font=ctk.CTkFont(size=8), text_color=("gray60", "gray50")).pack(side="right", padx=6, pady=2)

    def _shortcuts(self):
        self.bind("<F1>", lambda e: self._help())
        self.bind("<Control-q>", lambda e: self._on_close())
        self.bind("<Control-1>", lambda e: self.tabs.set("VBA Extractor"))
        self.bind("<Control-2>", lambda e: self.tabs.set("Python Analyzer"))
        self.bind("<Control-3>", lambda e: self.tabs.set("Folder Scanner"))
        self.bind("<Control-4>", lambda e: self.tabs.set("VBA Optimizer"))
        self.bind("<Control-5>", lambda e: self.tabs.set("VBA Analyzer"))

    def _theme(self, t):
        ctk.set_appearance_mode(t)
        self.config.set("ui.theme", t)

    def _help(self):
        dlg = ctk.CTkToplevel(self)
        dlg.title("Help")
        dlg.geometry("500x400")
        dlg.transient(self)
        dlg.grab_set()
        scroll = ctk.CTkScrollableFrame(dlg)
        scroll.pack(fill="both", expand=True, padx=15, pady=15)
        ctk.CTkLabel(scroll, text="CodeExtractPro v1.0", font=ctk.CTkFont(size=20, weight="bold")).pack(anchor="w")
        txt = """
Tools:
* VBA Extractor - Extract VBA from Office files
* Python Analyzer - Analyze Python code
* Folder Scanner - Scan directories
* VBA Optimizer - Optimize VBA code
* VBA Analyzer - Advanced VBA analysis with graphs

Shortcuts:
* F1 - Help
* Ctrl+Q - Quit
* Ctrl+1-5 - Switch tabs

Export: JSON, CSV, HTML, Excel
Settings auto-saved.
        """
        ctk.CTkLabel(scroll, text=txt.strip(), font=ctk.CTkFont(size=10), justify="left").pack(anchor="w", pady=10)
        ctk.CTkButton(dlg, text="Close", command=dlg.destroy, width=80).pack(pady=10)

    def _on_close(self):
        self.config.set("ui.window_width", self.winfo_width(), auto_save=False)
        self.config.set("ui.window_height", self.winfo_height(), auto_save=False)
        self.config.set("ui.window_x", self.winfo_x(), auto_save=False)
        self.config.set("ui.window_y", self.winfo_y(), auto_save=False)
        self.config.save()
        if self.config.config.ui.confirm_on_exit:
            if not messagebox.askyesno("Confirm", "Quit?"):
                return
        self.logger.info("Closing")
        self.destroy()

    def run(self):
        self.mainloop()
