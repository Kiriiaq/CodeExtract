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
        ext = {".json": "json", ".csv": "csv", ".html": "html"}.get(f".{fmt}", ".txt")
        path = filedialog.asksaveasfilename(title=f"Export {fmt.upper()}", defaultextension=f".{fmt}")
        if path:
            r = self.export_manager.export(self._last_result, path, fmt)
            if r.success:
                messagebox.showinfo("Success", f"Exported to {path}")
                if self.config.config.export.open_after_export: webbrowser.open(path)
            else:
                messagebox.showerror("Error", f"Export failed: {r.error}")

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
        ctk.CTkButton(dir_frame, text="▶ Analyze", command=self._analyze, height=28, width=90,
                      font=ctk.CTkFont(size=11, weight="bold"),
                      fg_color=("#10b981", "#059669"), hover_color=("#059669", "#047857")).pack(side="right", padx=4)

        # Options scrollable frame
        opts_scroll = ctk.CTkScrollableFrame(top, height=100, label_text="⚙️ Options d'analyse",
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
                        font=ctk.CTkFont(size=9), height=20, checkbox_width=16, checkbox_height=16).pack(anchor="w", padx=8, pady=(1, 4))

        # Column 3 - Metrics
        col3 = ctk.CTkFrame(opts_grid, fg_color=("gray95", "gray25"), corner_radius=6)
        col3.pack(side="left", fill="both", expand=True, padx=2, pady=2)
        ctk.CTkLabel(col3, text="Métriques", font=ctk.CTkFont(size=9, weight="bold"),
                     text_color=("#f59e0b", "#fbbf24")).pack(anchor="w", padx=6, pady=(4, 2))
        self.count_docstrings_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(col3, text="Docstrings", variable=self.count_docstrings_var,
                        font=ctk.CTkFont(size=9), height=20, checkbox_width=16, checkbox_height=16).pack(anchor="w", padx=8, pady=1)
        self.detect_duplicates_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(col3, text="Doublons", variable=self.detect_duplicates_var,
                        font=ctk.CTkFont(size=9), height=20, checkbox_width=16, checkbox_height=16).pack(anchor="w", padx=8, pady=(1, 4))

        # Column 4 - Filters
        col4 = ctk.CTkFrame(opts_grid, fg_color=("gray95", "gray25"), corner_radius=6)
        col4.pack(side="left", fill="both", expand=True, padx=2, pady=2)
        ctk.CTkLabel(col4, text="Filtres", font=ctk.CTkFont(size=9, weight="bold"),
                     text_color=("#059669", "#10b981")).pack(anchor="w", padx=6, pady=(4, 2))
        self.exclude_tests_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(col4, text="Exclure tests", variable=self.exclude_tests_var,
                        font=ctk.CTkFont(size=9), height=20, checkbox_width=16, checkbox_height=16).pack(anchor="w", padx=8, pady=1)
        self.exclude_init_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(col4, text="Exclure __init__", variable=self.exclude_init_var,
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

    def _analyze(self):
        d = self.dir_var.get()
        if not d:
            messagebox.showerror("Error", "Select a directory first")
            return
        def do():
            a = self.analyzer.analyze_directory(d, include_subdirs=self.subdirs_var.get())
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
        ctk.CTkButton(dir_frame, text="▶ Scan", command=self._scan, height=28, width=80,
                      font=ctk.CTkFont(size=11, weight="bold"),
                      fg_color=("#10b981", "#059669"), hover_color=("#059669", "#047857")).pack(side="right", padx=4)

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
            self._last_result = {"files": r.total_files, "dirs": r.total_directories, "size": r.total_size}
            self.slabels["files"].configure(text=str(r.total_files))
            self.slabels["dirs"].configure(text=str(r.total_directories))
            self.slabels["size"].configure(text=self._fmt_size(r.total_size))
            self.slabels["time"].configure(text=f"{r.scan_time:.1f}s")
            self.tree_text.delete("1.0", "end")
            self.tree_text.insert("1.0", self.scanner.generate_tree(r))
            self.set_progress(1.0, f"Scanned {r.total_files} files")
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
        ctk.CTkLabel(ih, text="📥 Input", font=ctk.CTkFont(size=11, weight="bold")).pack(side="left")
        ctk.CTkButton(ih, text="Load", width=50, height=24, font=ctk.CTkFont(size=9), command=self._load).pack(side="right", padx=2)
        ctk.CTkButton(ih, text="Paste", width=50, height=24, font=ctk.CTkFont(size=9), command=self._paste).pack(side="right", padx=2)
        ctk.CTkButton(ih, text="Clear", width=50, height=24, font=ctk.CTkFont(size=9), command=self._clear_input).pack(side="right", padx=2)
        self.input_text = ctk.CTkTextbox(inp, font=ctk.CTkFont(family="Consolas", size=9))
        self.input_text.pack(fill="both", expand=True, padx=6, pady=(0, 6))

        # Output panel
        out = ctk.CTkFrame(code)
        out.pack(side="right", fill="both", expand=True, padx=(2, 0))
        oh = ctk.CTkFrame(out, fg_color="transparent")
        oh.pack(fill="x", padx=6, pady=4)
        ctk.CTkLabel(oh, text="📤 Output", font=ctk.CTkFont(size=11, weight="bold")).pack(side="left")
        ctk.CTkButton(oh, text="Save", width=50, height=24, font=ctk.CTkFont(size=9), command=self._save).pack(side="right", padx=2)
        ctk.CTkButton(oh, text="Copy", width=50, height=24, font=ctk.CTkFont(size=9), command=self._copy).pack(side="right", padx=2)
        self.output_text = ctk.CTkTextbox(out, font=ctk.CTkFont(family="Consolas", size=9))
        self.output_text.pack(fill="both", expand=True, padx=6, pady=(0, 6))

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
        # Taille réduite pour une IHM plus compacte
        self.geometry(f"{min(ui.window_width, 1200)}x{min(ui.window_height, 750)}")
        self.minsize(900, 600)
        if ui.window_x and ui.window_y:
            self.geometry(f"+{ui.window_x}+{ui.window_y}")
        self._create_ui()
        self._shortcuts()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.logger.info("CodeExtractPro v1.0 started")

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
        t5 = self.tabs.add("Settings")
        t6 = self.tabs.add("Logs")
        VBAExtractorFrame(t1).pack(fill="both", expand=True)
        PythonAnalyzerFrame(t2).pack(fill="both", expand=True)
        FolderScannerFrame(t3).pack(fill="both", expand=True)
        VBAOptimizerFrame(t4).pack(fill="both", expand=True)
        SettingsFrame(t5, self).pack(fill="both", expand=True)
        LogsFrame(t6).pack(fill="both", expand=True)
        # Footer compact
        ftr = ctk.CTkFrame(main, height=22, fg_color=("gray90", "gray17"))
        ftr.pack(fill="x")
        ftr.pack_propagate(False)
        ctk.CTkLabel(ftr, text="Ready", font=ctk.CTkFont(size=9), text_color=("gray50", "gray60")).pack(side="left", padx=6, pady=2)
        ctk.CTkLabel(ftr, text="F1=Help | Ctrl+1-4=Tabs | Ctrl+Q=Quit", font=ctk.CTkFont(size=8), text_color=("gray60", "gray50")).pack(side="right", padx=6, pady=2)

    def _shortcuts(self):
        self.bind("<F1>", lambda e: self._help())
        self.bind("<Control-q>", lambda e: self._on_close())
        self.bind("<Control-1>", lambda e: self.tabs.set("VBA Extractor"))
        self.bind("<Control-2>", lambda e: self.tabs.set("Python Analyzer"))
        self.bind("<Control-3>", lambda e: self.tabs.set("Folder Scanner"))
        self.bind("<Control-4>", lambda e: self.tabs.set("VBA Optimizer"))

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

Shortcuts:
* F1 - Help
* Ctrl+Q - Quit
* Ctrl+1-4 - Switch tabs

Export: JSON, CSV, HTML
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
