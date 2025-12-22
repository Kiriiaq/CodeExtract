"""
Custom Widgets - Reusable UI components for the application.
Includes tooltips, status icons, scrollable frames, and log viewers.
"""

import tkinter as tk
from tkinter import ttk
from typing import Optional, Callable, List
import customtkinter as ctk
from datetime import datetime


class ToolTip:
    """Modern tooltip for any widget."""

    def __init__(self, widget: tk.Widget, text: str, delay: int = 500):
        self.widget = widget
        self.text = text
        self.delay = delay
        self.tooltip: Optional[tk.Toplevel] = None
        self.schedule_id: Optional[str] = None

        widget.bind("<Enter>", self._schedule_show)
        widget.bind("<Leave>", self._hide)
        widget.bind("<ButtonPress>", self._hide)

    def _schedule_show(self, event: tk.Event) -> None:
        self._cancel_schedule()
        self.schedule_id = self.widget.after(self.delay, lambda: self._show(event))

    def _cancel_schedule(self) -> None:
        if self.schedule_id:
            self.widget.after_cancel(self.schedule_id)
            self.schedule_id = None

    def _show(self, event: tk.Event) -> None:
        if self.tooltip or not self.text:
            return

        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5

        self.tooltip = tk.Toplevel(self.widget)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")

        # Modern dark style
        frame = tk.Frame(
            self.tooltip,
            background="#2b2b2b",
            relief="solid",
            borderwidth=1
        )
        frame.pack()

        label = tk.Label(
            frame,
            text=self.text,
            justify="left",
            background="#2b2b2b",
            foreground="#ffffff",
            font=("Segoe UI", 9),
            padx=8,
            pady=4,
            wraplength=300
        )
        label.pack()

    def _hide(self, event: Optional[tk.Event] = None) -> None:
        self._cancel_schedule()
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None

    def update_text(self, text: str) -> None:
        """Update tooltip text."""
        self.text = text


class StatusIcon(ctk.CTkFrame):
    """Animated status icon with different states."""

    STATES = {
        'pending': {'color': '#6b7280', 'icon': '○'},
        'in_progress': {'color': '#3b82f6', 'icon': '◐'},
        'completed': {'color': '#10b981', 'icon': '●'},
        'error': {'color': '#ef4444', 'icon': '✗'},
        'skipped': {'color': '#f59e0b', 'icon': '◌'},
        'disabled': {'color': '#4b5563', 'icon': '−'},
    }

    def __init__(self, master, state: str = 'pending', size: int = 20, **kwargs):
        super().__init__(master, width=size, height=size, fg_color="transparent", **kwargs)

        self.size = size
        self._state = state

        self.label = ctk.CTkLabel(
            self,
            text=self.STATES[state]['icon'],
            text_color=self.STATES[state]['color'],
            font=ctk.CTkFont(size=size - 4),
            width=size,
            height=size
        )
        self.label.pack(expand=True)

    def set_state(self, state: str) -> None:
        """Set the status state."""
        if state in self.STATES:
            self._state = state
            self.label.configure(
                text=self.STATES[state]['icon'],
                text_color=self.STATES[state]['color']
            )

    def get_state(self) -> str:
        """Get the current state."""
        return self._state


class ScrollableFrame(ctk.CTkScrollableFrame):
    """Enhanced scrollable frame with smooth scrolling."""

    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self._widgets: List[tk.Widget] = []

    def add_widget(self, widget: tk.Widget) -> None:
        """Track a widget added to the frame."""
        self._widgets.append(widget)

    def clear(self) -> None:
        """Remove all widgets from the frame."""
        for widget in self._widgets:
            widget.destroy()
        self._widgets.clear()


class LogViewer(ctk.CTkFrame):
    """Real-time log viewer with filtering and search."""

    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

        self.max_lines = 1000
        self._auto_scroll = True

        # Configure colors for different log levels
        self.level_colors = {
            'DEBUG': '#6b7280',
            'INFO': '#e5e7eb',
            'SUCCESS': '#10b981',
            'WARNING': '#f59e0b',
            'ERROR': '#ef4444',
            'CRITICAL': '#ec4899',
        }

        self._setup_ui()

    def _setup_ui(self) -> None:
        """Set up the UI components."""
        # Toolbar
        toolbar = ctk.CTkFrame(self, fg_color="transparent", height=30)
        toolbar.pack(fill="x", padx=5, pady=(5, 0))

        # Level filter
        self.level_var = ctk.StringVar(value="ALL")
        level_menu = ctk.CTkOptionMenu(
            toolbar,
            values=["ALL", "DEBUG", "INFO", "SUCCESS", "WARNING", "ERROR"],
            variable=self.level_var,
            command=self._filter_logs,
            width=100,
            height=28
        )
        level_menu.pack(side="left", padx=(0, 5))
        ToolTip(level_menu, "Filter logs by level")

        # Search entry
        self.search_var = ctk.StringVar()
        self.search_var.trace_add("write", lambda *args: self._filter_logs())
        search_entry = ctk.CTkEntry(
            toolbar,
            placeholder_text="Search logs...",
            textvariable=self.search_var,
            width=200,
            height=28
        )
        search_entry.pack(side="left", padx=(0, 5))

        # Auto-scroll toggle
        self.autoscroll_var = ctk.BooleanVar(value=True)
        autoscroll_cb = ctk.CTkCheckBox(
            toolbar,
            text="Auto-scroll",
            variable=self.autoscroll_var,
            command=self._toggle_autoscroll,
            width=100,
            height=28
        )
        autoscroll_cb.pack(side="left", padx=(0, 5))

        # Clear button
        clear_btn = ctk.CTkButton(
            toolbar,
            text="Clear",
            command=self.clear,
            width=60,
            height=28
        )
        clear_btn.pack(side="right")

        # Text area
        text_frame = ctk.CTkFrame(self)
        text_frame.pack(fill="both", expand=True, padx=5, pady=5)

        self.text = tk.Text(
            text_frame,
            wrap="word",
            font=("Consolas", 10),
            bg="#1a1a2e",
            fg="#e5e7eb",
            insertbackground="#ffffff",
            selectbackground="#3b82f6",
            relief="flat",
            padx=10,
            pady=10
        )
        self.text.pack(side="left", fill="both", expand=True)

        scrollbar = ctk.CTkScrollbar(text_frame, command=self.text.yview)
        scrollbar.pack(side="right", fill="y")
        self.text.configure(yscrollcommand=scrollbar.set)

        # Configure tags for colors
        for level, color in self.level_colors.items():
            self.text.tag_configure(level, foreground=color)
        self.text.tag_configure("timestamp", foreground="#6b7280")

        # Store all entries for filtering
        self._entries: List[tuple] = []

    def _toggle_autoscroll(self) -> None:
        """Toggle auto-scroll feature."""
        self._auto_scroll = self.autoscroll_var.get()

    def _filter_logs(self, *args) -> None:
        """Filter displayed logs based on level and search."""
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")

        level_filter = self.level_var.get()
        search_text = self.search_var.get().lower()

        for timestamp, level, message in self._entries:
            if level_filter != "ALL" and level != level_filter:
                continue
            if search_text and search_text not in message.lower():
                continue

            self._insert_entry(timestamp, level, message)

        self.text.configure(state="disabled")
        if self._auto_scroll:
            self.text.see("end")

    def _insert_entry(self, timestamp: str, level: str, message: str) -> None:
        """Insert a single log entry."""
        self.text.insert("end", f"[{timestamp}] ", "timestamp")
        self.text.insert("end", f"[{level}] ", level)
        self.text.insert("end", f"{message}\n")

    def add_log(self, message: str, level: str = "INFO") -> None:
        """Add a log entry."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self._entries.append((timestamp, level, message))

        # Trim old entries
        if len(self._entries) > self.max_lines:
            self._entries = self._entries[-self.max_lines:]

        # Apply current filter
        level_filter = self.level_var.get()
        search_text = self.search_var.get().lower()

        should_show = (level_filter == "ALL" or level == level_filter)
        should_show = should_show and (not search_text or search_text in message.lower())

        if should_show:
            self.text.configure(state="normal")
            self._insert_entry(timestamp, level, message)
            self.text.configure(state="disabled")
            if self._auto_scroll:
                self.text.see("end")

    def clear(self) -> None:
        """Clear all logs."""
        self._entries.clear()
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.configure(state="disabled")

    def export(self, file_path: str) -> None:
        """Export logs to a file."""
        with open(file_path, 'w', encoding='utf-8') as f:
            for timestamp, level, message in self._entries:
                f.write(f"[{timestamp}] [{level}] {message}\n")


class StepCard(ctk.CTkFrame):
    """Card widget for displaying a workflow step."""

    def __init__(self, master, step_id: str, name: str, description: str,
                 enabled: bool = True, on_toggle: Optional[Callable] = None, **kwargs):
        super().__init__(master, **kwargs)

        self.step_id = step_id
        self.on_toggle = on_toggle

        # Configure frame
        self.configure(corner_radius=8, fg_color=("#f3f4f6", "#1f2937"))

        # Status icon
        self.status_icon = StatusIcon(self, state='pending', size=24)
        self.status_icon.pack(side="left", padx=(15, 10), pady=15)

        # Info section
        info_frame = ctk.CTkFrame(self, fg_color="transparent")
        info_frame.pack(side="left", fill="both", expand=True, pady=10)

        name_label = ctk.CTkLabel(
            info_frame,
            text=name,
            font=ctk.CTkFont(size=14, weight="bold"),
            anchor="w"
        )
        name_label.pack(fill="x")

        desc_label = ctk.CTkLabel(
            info_frame,
            text=description,
            font=ctk.CTkFont(size=11),
            text_color=("#6b7280", "#9ca3af"),
            anchor="w"
        )
        desc_label.pack(fill="x")

        # Duration label (hidden by default)
        self.duration_label = ctk.CTkLabel(
            info_frame,
            text="",
            font=ctk.CTkFont(size=10),
            text_color=("#9ca3af", "#6b7280"),
            anchor="w"
        )
        self.duration_label.pack(fill="x")

        # Enable toggle
        self.enabled_var = ctk.BooleanVar(value=enabled)
        toggle = ctk.CTkSwitch(
            self,
            text="",
            variable=self.enabled_var,
            command=self._on_toggle,
            width=40
        )
        toggle.pack(side="right", padx=15)
        ToolTip(toggle, "Enable/disable this step")

    def _on_toggle(self) -> None:
        """Handle toggle event."""
        if self.on_toggle:
            self.on_toggle(self.step_id, self.enabled_var.get())

    def set_status(self, status: str) -> None:
        """Set the step status."""
        self.status_icon.set_state(status)

    def set_duration(self, seconds: float) -> None:
        """Set the duration display."""
        if seconds > 0:
            self.duration_label.configure(text=f"Duration: {seconds:.2f}s")
        else:
            self.duration_label.configure(text="")

    def is_enabled(self) -> bool:
        """Check if step is enabled."""
        return self.enabled_var.get()
