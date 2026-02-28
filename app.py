"""
╔══════════════════════════════════════════════════════════════════╗
║                  CUBE DATA CHANGER AIO                          ║
║                                                                  ║
║  All-in-One: Generate + Process cube/mortar test data            ║
║                                                                  ║
║  Developer : Sandeep (https://github.com/Sandeep2062)           ║
║  Repository: github.com/Sandeep2062/Cube-Data-Changer-AIO      ║
║                                                                  ║
║  © 2026 Sandeep — All Rights Reserved                           ║
╚══════════════════════════════════════════════════════════════════╝
"""

import os
import sys
import threading
import webbrowser

import customtkinter as ctk
from tkinter import filedialog, messagebox

import settings as app_settings
from generator import CONCRETE_GRADES, MORTAR_TYPES, ALL_TYPES, grade_display_name
from processor import process

# ── Appearance ──────────────────────────────────────────────────────────────

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# Colour palette (dark theme constants)
BG_DARK       = "#0f0f0f"
BG_SIDEBAR    = "#161618"
BG_CARD       = "#1c1c1e"
BG_CARD_HOVER = "#242426"
ACCENT        = "#3b82f6"    # blue-500
ACCENT_HOVER  = "#2563eb"    # blue-600
GREEN         = "#22c55e"
GREEN_HOVER   = "#16a34a"
RED           = "#ef4444"
RED_HOVER     = "#dc2626"
ORANGE        = "#f59e0b"
TEXT_PRIMARY   = "#f5f5f7"
TEXT_SECONDARY = "#a1a1aa"
TEXT_DIM       = "#71717a"
BORDER_COLOR   = "#27272a"

VERSION = "1.0.0"


def resource_path(relative_path):
    """Get path to bundled resource (works inside PyInstaller)."""
    try:
        base = sys._MEIPASS
    except AttributeError:
        base = os.path.abspath(".")
    return os.path.join(base, relative_path)


# ── Main Application ────────────────────────────────────────────────────────

class CubeDataChangerAIO:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title("Cube Data Changer AIO")
        self.root.geometry("1200x800")
        self.root.minsize(1050, 700)
        self.root.configure(fg_color=BG_DARK)

        try:
            self.root.iconbitmap(resource_path("icon.ico"))
        except Exception:
            pass

        # State
        self._load_settings()
        self.grade_vars = {}          # grade -> BooleanVar
        self.processing = False

        # Build UI
        self._build_ui()

    # ── Settings persistence ────────────────────────────────────────────────

    def _load_settings(self):
        s = app_settings.load()
        self.office_path    = ctk.StringVar(value=s.get("office_path", ""))
        self.output_path    = ctk.StringVar(value=s.get("output_path", ""))
        self.calendar_path  = ctk.StringVar(value=s.get("calendar_path", ""))
        self.mode_var       = ctk.StringVar(value=s.get("mode", "generate+date"))
        self.saved_grades   = s.get("selected_grades", [])
        self.saved_grade_files = [f for f in s.get("grade_files", []) if os.path.exists(f)]

    def _save_settings(self):
        selected = [g for g, v in self.grade_vars.items() if v.get()]
        app_settings.save({
            "office_path":     self.office_path.get(),
            "output_path":     self.output_path.get(),
            "calendar_path":   self.calendar_path.get(),
            "mode":            self.mode_var.get(),
            "selected_grades": selected,
            "grade_files":     getattr(self, "legacy_grade_files", []),
        })

    # ── UI Construction ─────────────────────────────────────────────────────

    def _build_ui(self):
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(0, weight=1)

        self._build_sidebar()
        self._build_main()

    # Sidebar ────────────────────────────────────────────────────────────────

    def _build_sidebar(self):
        sb = ctk.CTkFrame(self.root, width=280, corner_radius=0, fg_color=BG_SIDEBAR,
                          border_width=0)
        sb.grid(row=0, column=0, sticky="nsew")
        sb.grid_propagate(False)
        sb.grid_columnconfigure(0, weight=1)
        sb.grid_rowconfigure(20, weight=1)  # spacer

        r = 0

        # Logo / branding
        logo_frame = ctk.CTkFrame(sb, fg_color="transparent")
        logo_frame.grid(row=r, column=0, padx=20, pady=(30, 5)); r += 1
        try:
            from PIL import Image
            img = Image.open(resource_path("logo.png")).resize((64, 64), Image.Resampling.LANCZOS)
            self._logo_photo = ctk.CTkImage(light_image=img, dark_image=img, size=(64, 64))
            ctk.CTkLabel(logo_frame, image=self._logo_photo, text="").pack()
        except Exception:
            ctk.CTkLabel(logo_frame, text="◆", font=ctk.CTkFont(size=48),
                         text_color=ACCENT).pack()

        ctk.CTkLabel(sb, text="CUBE DATA\nCHANGER AIO",
                     font=ctk.CTkFont(size=20, weight="bold"),
                     text_color=TEXT_PRIMARY).grid(row=r, column=0, padx=20, pady=(2, 5)); r += 1

        ctk.CTkLabel(sb, text=f"v{VERSION}",
                     font=ctk.CTkFont(size=11),
                     text_color=TEXT_DIM).grid(row=r, column=0, padx=20, pady=(0, 20)); r += 1

        # Divider
        ctk.CTkFrame(sb, height=1, fg_color=BORDER_COLOR).grid(
            row=r, column=0, sticky="ew", padx=20, pady=(0, 15)); r += 1

        # ── Mode selection ──────────────────────────────────────────────────
        ctk.CTkLabel(sb, text="PROCESSING MODE",
                     font=ctk.CTkFont(size=11, weight="bold"),
                     text_color=TEXT_SECONDARY, anchor="w").grid(
            row=r, column=0, padx=24, pady=(0, 8), sticky="w"); r += 1

        modes = [
            ("generate+date", "⚡  Auto Generate + Date"),
            ("generate",      "🔄  Auto Generate Only"),
            ("date_only",     "📅  Date Only"),
            ("grade_files+date", "📁  Files + Date (Legacy)"),
            ("grade_files",      "📁  Files Only (Legacy)"),
        ]
        for val, label in modes:
            rb = ctk.CTkRadioButton(
                sb, text=label, variable=self.mode_var, value=val,
                command=self._on_mode_change,
                font=ctk.CTkFont(size=12),
                text_color=TEXT_PRIMARY,
                fg_color=ACCENT, border_color=TEXT_DIM,
                hover_color=ACCENT_HOVER,
            )
            rb.grid(row=r, column=0, padx=28, pady=5, sticky="w"); r += 1

        # Divider
        ctk.CTkFrame(sb, height=1, fg_color=BORDER_COLOR).grid(
            row=r, column=0, sticky="ew", padx=20, pady=15); r += 1

        # ── Grade selector (for auto-generate modes) ───────────────────────
        self._grade_header_row = r
        self._grade_label = ctk.CTkLabel(
            sb, text="SELECT GRADES TO GENERATE",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=TEXT_SECONDARY, anchor="w")
        self._grade_label.grid(row=r, column=0, padx=24, pady=(0, 8), sticky="w"); r += 1

        # Quick select buttons
        self._grade_btn_frame = ctk.CTkFrame(sb, fg_color="transparent")
        self._grade_btn_frame.grid(row=r, column=0, padx=24, pady=(0, 6), sticky="ew"); r += 1

        ctk.CTkButton(self._grade_btn_frame, text="All", width=55, height=26,
                       font=ctk.CTkFont(size=11), fg_color=ACCENT, hover_color=ACCENT_HOVER,
                       command=lambda: self._toggle_all_grades(True)).pack(side="left", padx=(0, 4))
        ctk.CTkButton(self._grade_btn_frame, text="None", width=55, height=26,
                       font=ctk.CTkFont(size=11), fg_color="#3f3f46", hover_color="#52525b",
                       command=lambda: self._toggle_all_grades(False)).pack(side="left", padx=(0, 4))
        ctk.CTkButton(self._grade_btn_frame, text="Concrete", width=72, height=26,
                       font=ctk.CTkFont(size=11), fg_color="#3f3f46", hover_color="#52525b",
                       command=self._select_concrete_only).pack(side="left", padx=(0, 4))
        ctk.CTkButton(self._grade_btn_frame, text="Mortar", width=60, height=26,
                       font=ctk.CTkFont(size=11), fg_color="#3f3f46", hover_color="#52525b",
                       command=self._select_mortar_only).pack(side="left")

        # Scrollable checkbox area
        self._grade_scroll = ctk.CTkScrollableFrame(
            sb, height=170, fg_color=BG_CARD, corner_radius=8,
            border_width=1, border_color=BORDER_COLOR,
            scrollbar_button_color="#3f3f46", scrollbar_button_hover_color="#52525b")
        self._grade_scroll.grid(row=r, column=0, padx=20, pady=(0, 10), sticky="ew"); r += 1

        for g in ALL_TYPES:
            var = ctk.BooleanVar(value=(g in self.saved_grades))
            self.grade_vars[g] = var
            cb = ctk.CTkCheckBox(
                self._grade_scroll, text=grade_display_name(g),
                variable=var, font=ctk.CTkFont(size=12),
                text_color=TEXT_PRIMARY,
                fg_color=ACCENT, hover_color=ACCENT_HOVER,
                border_color=TEXT_DIM, checkmark_color="white",
            )
            cb.pack(anchor="w", padx=8, pady=3)

        self._grade_widgets_end_row = r

        # ── Legacy grade files area ─────────────────────────────────────────
        self.legacy_grade_files = list(self.saved_grade_files)

        self._legacy_label = ctk.CTkLabel(
            sb, text="GRADE FILES (LEGACY)",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=TEXT_SECONDARY, anchor="w")
        self._legacy_label.grid(row=r, column=0, padx=24, pady=(0, 8), sticky="w"); r += 1

        self._legacy_listbox = ctk.CTkTextbox(
            sb, height=90, font=ctk.CTkFont(size=11),
            fg_color=BG_CARD, border_width=1, border_color=BORDER_COLOR,
            text_color=TEXT_PRIMARY)
        self._legacy_listbox.grid(row=r, column=0, padx=20, pady=(0, 6), sticky="ew"); r += 1

        self._legacy_btn_frame = ctk.CTkFrame(sb, fg_color="transparent")
        self._legacy_btn_frame.grid(row=r, column=0, padx=20, pady=(0, 10), sticky="ew"); r += 1

        ctk.CTkButton(self._legacy_btn_frame, text="+ Add Files", width=100, height=28,
                       font=ctk.CTkFont(size=11, weight="bold"),
                       fg_color=ACCENT, hover_color=ACCENT_HOVER,
                       command=self._add_legacy_files).pack(side="left", padx=(0, 6))
        ctk.CTkButton(self._legacy_btn_frame, text="Clear", width=70, height=28,
                       font=ctk.CTkFont(size=11, weight="bold"),
                       fg_color=RED, hover_color=RED_HOVER,
                       command=self._clear_legacy_files).pack(side="left")

        self._update_legacy_listbox()

        # Spacer (pushes footer down)
        sb.grid_rowconfigure(20, weight=1)

        # ── Footer links ───────────────────────────────────────────────────
        ctk.CTkFrame(sb, height=1, fg_color=BORDER_COLOR).grid(
            row=21, column=0, sticky="ew", padx=20, pady=(10, 10))

        link_frame = ctk.CTkFrame(sb, fg_color="transparent")
        link_frame.grid(row=22, column=0, padx=20, pady=(0, 20), sticky="ew")

        ctk.CTkButton(link_frame, text="GitHub", width=100, height=30,
                       font=ctk.CTkFont(size=11), fg_color="#3f3f46", hover_color="#52525b",
                       command=lambda: webbrowser.open(
                           "https://github.com/Sandeep2062/Cube-Data-Changer-AIO")
                       ).pack(side="left", padx=(0, 6))
        ctk.CTkButton(link_frame, text="Instagram", width=100, height=30,
                       font=ctk.CTkFont(size=11),
                       fg_color="#E1306C", hover_color="#C13584",
                       command=lambda: webbrowser.open(
                           "https://www.instagram.com/sandeep._.2062/")
                       ).pack(side="left")

        # Initial visibility
        self._on_mode_change()

    # Main content ───────────────────────────────────────────────────────────

    def _build_main(self):
        main = ctk.CTkFrame(self.root, fg_color=BG_DARK, corner_radius=0)
        main.grid(row=0, column=1, sticky="nsew", padx=(0, 0), pady=0)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(4, weight=1)

        pad_x = 24
        pad_y = 10

        # Header
        header = ctk.CTkFrame(main, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=pad_x, pady=(24, 8))
        ctk.CTkLabel(header, text="Cube Data Changer AIO",
                     font=ctk.CTkFont(size=26, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(side="left")
        ctk.CTkLabel(header, text="Generate · Process · Done",
                     font=ctk.CTkFont(size=13),
                     text_color=TEXT_DIM).pack(side="left", padx=(12, 0), pady=(6, 0))

        # ── File picker cards ───────────────────────────────────────────────
        cards_frame = ctk.CTkFrame(main, fg_color="transparent")
        cards_frame.grid(row=1, column=0, sticky="ew", padx=pad_x, pady=pad_y)
        cards_frame.grid_columnconfigure(0, weight=1)

        # Calendar
        self.calendar_card = self._file_card(
            cards_frame, row=0, icon="📅", label="Calendar File",
            var=self.calendar_path, placeholder="Select calendar Excel file...",
            browse_cmd=lambda: self._browse_file(self.calendar_path, "Select Calendar File"))

        # Office template
        self._file_card(
            cards_frame, row=1, icon="📄", label="Office Template",
            var=self.office_path, placeholder="Select office template Excel...",
            browse_cmd=lambda: self._browse_file(self.office_path, "Select Office Template"))

        # Output folder
        self._folder_card(
            cards_frame, row=2, icon="💾", label="Output Folder",
            var=self.output_path, placeholder="Select output destination...",
            browse_cmd=self._browse_output)

        # ── Action buttons ──────────────────────────────────────────────────
        btn_frame = ctk.CTkFrame(main, fg_color="transparent")
        btn_frame.grid(row=2, column=0, sticky="ew", padx=pad_x, pady=(14, 0))
        btn_frame.grid_columnconfigure(0, weight=1)

        self.start_btn = ctk.CTkButton(
            btn_frame, text="▶   START PROCESSING",
            font=ctk.CTkFont(size=18, weight="bold"), height=56,
            fg_color=GREEN, hover_color=GREEN_HOVER,
            text_color="white", corner_radius=12,
            command=self._run)
        self.start_btn.grid(row=0, column=0, sticky="ew")

        # Progress
        self.progress = ctk.CTkProgressBar(
            main, height=6, corner_radius=3,
            fg_color=BORDER_COLOR, progress_color=ACCENT)
        self.progress.grid(row=3, column=0, sticky="ew", padx=pad_x, pady=(12, 0))
        self.progress.set(0)

        # ── Log area ────────────────────────────────────────────────────────
        log_header = ctk.CTkFrame(main, fg_color="transparent")
        log_header.grid(row=4, column=0, sticky="new", padx=pad_x, pady=(14, 4))
        ctk.CTkLabel(log_header, text="Processing Log",
                     font=ctk.CTkFont(size=14, weight="bold"),
                     text_color=TEXT_SECONDARY).pack(side="left")

        self.log_box = ctk.CTkTextbox(
            main, font=ctk.CTkFont(family="Consolas", size=12),
            fg_color=BG_CARD, border_width=1, border_color=BORDER_COLOR,
            text_color="#d4d4d8", corner_radius=10,
            wrap="word")
        self.log_box.grid(row=5, column=0, sticky="nsew", padx=pad_x, pady=(0, 10))
        main.grid_rowconfigure(5, weight=1)

        # Footer
        ctk.CTkLabel(main,
                     text="© 2026 Sandeep  ·  github.com/Sandeep2062/Cube-Data-Changer-AIO",
                     font=ctk.CTkFont(size=11), text_color=TEXT_DIM
                     ).grid(row=6, column=0, pady=(2, 12))

    # ── Card helpers ────────────────────────────────────────────────────────

    def _file_card(self, parent, row, icon, label, var, placeholder, browse_cmd):
        card = ctk.CTkFrame(parent, fg_color=BG_CARD, corner_radius=10,
                            border_width=1, border_color=BORDER_COLOR)
        card.grid(row=row, column=0, sticky="ew", pady=6)
        card.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(card, text=f"{icon}  {label}",
                     font=ctk.CTkFont(size=13, weight="bold"),
                     text_color=TEXT_PRIMARY).grid(row=0, column=0, padx=16, pady=14, sticky="w")

        entry = ctk.CTkEntry(card, textvariable=var, placeholder_text=placeholder,
                             height=38, font=ctk.CTkFont(size=12),
                             fg_color="#27272a", border_color=BORDER_COLOR,
                             text_color=TEXT_PRIMARY, placeholder_text_color=TEXT_DIM)
        entry.grid(row=0, column=1, padx=(4, 8), pady=14, sticky="ew")

        btn = ctk.CTkButton(card, text="Browse", width=90, height=36,
                            font=ctk.CTkFont(size=12, weight="bold"),
                            fg_color=ACCENT, hover_color=ACCENT_HOVER,
                            corner_radius=8, command=browse_cmd)
        btn.grid(row=0, column=2, padx=(0, 14), pady=14)
        return card

    def _folder_card(self, parent, row, icon, label, var, placeholder, browse_cmd):
        return self._file_card(parent, row, icon, label, var, placeholder, browse_cmd)

    def _browse_file(self, var, title="Select File"):
        path = filedialog.askopenfilename(
            title=title, filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")])
        if path:
            var.set(path)

    def _browse_output(self):
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_path.set(folder)

    # ── Mode switching ──────────────────────────────────────────────────────

    def _on_mode_change(self):
        mode = self.mode_var.get()
        is_generate = "generate" in mode
        is_legacy   = "grade_files" in mode
        is_date     = "date" in mode

        # Grade checkboxes
        for w in (self._grade_label, self._grade_btn_frame, self._grade_scroll):
            w.grid() if is_generate else w.grid_remove()

        # Legacy file list
        for w in (self._legacy_label, self._legacy_listbox, self._legacy_btn_frame):
            w.grid() if is_legacy else w.grid_remove()

        # Calendar card
        if hasattr(self, "calendar_card"):
            self.calendar_card.grid() if is_date else self.calendar_card.grid_remove()

    # ── Grade selection helpers ─────────────────────────────────────────────

    def _toggle_all_grades(self, state):
        for v in self.grade_vars.values():
            v.set(state)

    def _select_concrete_only(self):
        for g, v in self.grade_vars.items():
            v.set(g in CONCRETE_GRADES)

    def _select_mortar_only(self):
        for g, v in self.grade_vars.items():
            v.set(g in MORTAR_TYPES)

    # ── Legacy grade file management ────────────────────────────────────────

    def _add_legacy_files(self):
        files = filedialog.askopenfilenames(
            title="Select Grade Excel Files",
            filetypes=[("Excel Files", "*.xlsx")])
        for f in files:
            if f not in self.legacy_grade_files:
                self.legacy_grade_files.append(f)
        self._update_legacy_listbox()

    def _clear_legacy_files(self):
        self.legacy_grade_files.clear()
        self._update_legacy_listbox()

    def _update_legacy_listbox(self):
        self._legacy_listbox.delete("0.0", "end")
        if not self.legacy_grade_files:
            self._legacy_listbox.insert("end", "  No files selected\n")
        else:
            for f in self.legacy_grade_files:
                self._legacy_listbox.insert("end", f"  📄 {os.path.basename(f)}\n")

    # ── Logging ─────────────────────────────────────────────────────────────

    def _log(self, msg):
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")
        self.root.update_idletasks()

    def _set_progress(self, val):
        self.progress.set(max(0, min(1, val)))
        self.root.update_idletasks()

    # ── Processing ──────────────────────────────────────────────────────────

    def _validate(self):
        mode = self.mode_var.get()

        if not self.office_path.get():
            messagebox.showerror("Missing Input", "Please select an Office Template file.")
            return False
        if not self.output_path.get():
            messagebox.showerror("Missing Input", "Please select an Output Folder.")
            return False

        if "generate" in mode:
            selected = [g for g, v in self.grade_vars.items() if v.get()]
            if not selected:
                messagebox.showerror("Missing Input",
                                     "Please select at least one grade to generate.")
                return False

        if "grade_files" in mode:
            if not self.legacy_grade_files:
                messagebox.showerror("Missing Input",
                                     "Please add grade Excel files for legacy processing.")
                return False

        if "date" in mode:
            if not self.calendar_path.get():
                messagebox.showerror("Missing Input",
                                     "Please select a Calendar file for date processing.")
                return False

        return True

    def _run(self):
        if self.processing:
            return
        if not self._validate():
            return

        self.processing = True
        self.start_btn.configure(state="disabled", text="⏳  Processing...",
                                 fg_color="#3f3f46")
        self.log_box.delete("0.0", "end")
        self.progress.set(0)
        self._save_settings()

        mode = self.mode_var.get()
        selected_grades = [g for g, v in self.grade_vars.items() if v.get()] \
                          if "generate" in mode else None
        grade_files = self.legacy_grade_files if "grade_files" in mode else None
        calendar = self.calendar_path.get() if "date" in mode else None

        def worker():
            try:
                total = process(
                    office_file=self.office_path.get(),
                    output_folder=self.output_path.get(),
                    mode=mode,
                    log=self._log,
                    selected_grades=selected_grades,
                    grade_files=grade_files,
                    calendar_file=calendar,
                    progress_cb=self._set_progress,
                )
                self.root.after(0, lambda: self._on_done(total))
            except Exception as e:
                self.root.after(0, lambda: self._on_error(str(e)))

        threading.Thread(target=worker, daemon=True).start()

    def _on_done(self, total):
        self.processing = False
        self.progress.set(1.0)
        self.start_btn.configure(state="normal", text="▶   START PROCESSING",
                                 fg_color=GREEN)
        self._log(f"\n✅ Processing complete — {total} operations performed")

        # Sound (Windows only, silently ignored elsewhere)
        try:
            import winsound
            winsound.MessageBeep()
        except Exception:
            pass

        messagebox.showinfo("✓ Complete",
                            f"Processing finished!\n\nTotal operations: {total}")
        self.progress.set(0)

    def _on_error(self, err):
        self.processing = False
        self.progress.set(0)
        self.start_btn.configure(state="normal", text="▶   START PROCESSING",
                                 fg_color=GREEN)
        self._log(f"\n✖ ERROR: {err}")
        messagebox.showerror("Error", f"Processing failed:\n{err}")

    # ── Run ─────────────────────────────────────────────────────────────────

    def run(self):
        self.root.mainloop()


# ── Entry Point ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = CubeDataChangerAIO()
    app.run()
