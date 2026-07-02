"""Word Helper — a small dark-themed GUI for batch operations on .docx files.

Tabs:
  * Find & Replace  — batch text find/replace with no 255-character limit.
  * Replace Image   — swap an embedded image (old file -> new file) across many
                      documents at once.

The layout is modular: each feature is a build_*_tab method, and the file list
(FileSelector) and image pickers (ImagePicker) are reusable widgets.
"""

from __future__ import annotations

import io
import os
import queue
import shutil
import threading
from tkinter import filedialog

import customtkinter as ctk
from PIL import Image as PILImage

from docx_replace import replace_in_document
from docx_image_replace import replace_images_in_document, read_image_info

APP_NAME = "Word Helper"
ACCENT = "#3d7eff"
ACCENT_HOVER = "#2f66d6"
GREY = "#8a8f98"
CARD = "#2b2d31"

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

DOCX_TYPES = [("Word documents", "*.docx"), ("All files", "*.*")]
IMAGE_TYPES = [
    ("Images", "*.png *.jpg *.jpeg *.gif *.bmp *.tif *.tiff *.emf *.wmf"),
    ("All files", "*.*"),
]


# --------------------------------------------------------------------- widgets
class FileSelector(ctk.CTkFrame):
    """A titled, scrollable list of selected .docx files with add/clear/remove."""

    def __init__(self, master, title="Word files", height=140):
        super().__init__(master, fg_color="transparent")
        self.files: list[str] = []

        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.pack(fill="x")
        ctk.CTkLabel(
            bar, text=title, font=ctk.CTkFont(size=15, weight="bold")
        ).pack(side="left")
        ctk.CTkButton(
            bar, text="Clear", width=70, fg_color="#3a3d44",
            hover_color="#4a4e57", command=self.clear,
        ).pack(side="right")
        ctk.CTkButton(
            bar, text="+ Add files", width=110, fg_color=ACCENT,
            hover_color=ACCENT_HOVER, command=self.add,
        ).pack(side="right", padx=(0, 8))

        self._list = ctk.CTkScrollableFrame(self, height=height, label_text="")
        self._list.pack(fill="both", expand=True, pady=(6, 0))
        self._render()

    def add(self):
        for p in filedialog.askopenfilenames(
            title="Select Word documents", filetypes=DOCX_TYPES
        ):
            if p not in self.files:
                self.files.append(p)
        self._render()

    def clear(self):
        self.files.clear()
        self._render()

    def remove(self, path):
        if path in self.files:
            self.files.remove(path)
        self._render()

    def _render(self):
        for child in self._list.winfo_children():
            child.destroy()
        if not self.files:
            ctk.CTkLabel(
                self._list, text="No files selected.", text_color=GREY
            ).pack(anchor="w", padx=6, pady=6)
            return
        for path in self.files:
            row = ctk.CTkFrame(self._list, fg_color=CARD)
            row.pack(fill="x", pady=2, padx=2)
            ctk.CTkButton(
                row, text="✕", width=28, height=24, fg_color="#3a3d44",
                hover_color="#a13d3d", command=lambda p=path: self.remove(p),
            ).pack(side="right", padx=4, pady=2)
            ctk.CTkLabel(row, text=os.path.basename(path), anchor="w").pack(
                side="left", padx=8, pady=2
            )


class ImagePicker(ctk.CTkFrame):
    """A card that lets the user choose an image file and shows a thumbnail +
    format/dimension info. ``bytes`` holds the chosen file's raw data."""

    def __init__(self, master, title):
        super().__init__(master, fg_color=CARD, corner_radius=8)
        self.path: str | None = None
        self.bytes: bytes | None = None
        self._imgref = None

        ctk.CTkLabel(
            self, text=title, font=ctk.CTkFont(size=14, weight="bold")
        ).pack(pady=(10, 6))
        self.preview = ctk.CTkLabel(
            self, text="no image", width=140, height=140,
            fg_color="#1f2124", corner_radius=6, text_color=GREY,
        )
        self.preview.pack(padx=12)
        self.info = ctk.CTkLabel(
            self, text="", text_color=GREY, font=ctk.CTkFont(size=12),
            justify="center",
        )
        self.info.pack(pady=(6, 4))
        ctk.CTkButton(
            self, text="Choose…", width=120, fg_color=ACCENT,
            hover_color=ACCENT_HOVER, command=self._choose,
        ).pack(pady=(0, 12))

    def valid(self) -> bool:
        return self.bytes is not None

    def _choose(self):
        path = filedialog.askopenfilename(title="Select image", filetypes=IMAGE_TYPES)
        if not path:
            return
        try:
            with open(path, "rb") as fh:
                data = fh.read()
        except Exception as exc:
            self.info.configure(text=f"could not read file:\n{exc}")
            return
        self.path = path
        self.bytes = data

        try:
            info = read_image_info(data)
            self.info.configure(
                text=f"{os.path.basename(path)}\n"
                     f"{info.px_width}×{info.px_height} px · {info.ext.upper()}"
            )
        except Exception:
            self.info.configure(text=f"{os.path.basename(path)}\n(format not recognised)")

        try:
            im = PILImage.open(io.BytesIO(data))
            im.thumbnail((130, 130))
            self._imgref = ctk.CTkImage(light_image=im, dark_image=im, size=im.size)
            self.preview.configure(image=self._imgref, text="")
        except Exception:
            self._imgref = None
            self.preview.configure(image="", text="(no preview)")


# ------------------------------------------------------------------------- app
class WordHelperApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("980x860")
        self.minsize(820, 680)

        self._msg_queue: queue.Queue = queue.Queue()
        self._running = False

        self._build_header()

        self.tabview = ctk.CTkTabview(self, fg_color="transparent")
        self.tabview.pack(fill="both", expand=True, padx=16, pady=(0, 12))

        self.tabview.add("Find & Replace")
        self.build_find_replace_tab(self.tabview.tab("Find & Replace"))

        self.tabview.add("Replace Image")
        self.build_image_replace_tab(self.tabview.tab("Replace Image"))

        self.after(100, self._drain_queue)

    # ------------------------------------------------------------------ header
    def _build_header(self):
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=16, pady=(14, 6))
        ctk.CTkLabel(
            header, text="Word Helper", font=ctk.CTkFont(size=24, weight="bold")
        ).pack(side="left")
        ctk.CTkLabel(
            header, text="Batch tools for .docx files",
            font=ctk.CTkFont(size=13), text_color=GREY,
        ).pack(side="left", padx=(12, 0), pady=(6, 0))

    # ------------------------------------------------ shared run-panel builder
    def _build_run_panel(self, tab, row, button_text, command):
        """Create a run button + progress bar (one row) and a log box (next
        row, expanding). Returns (run_btn, progress, log_box)."""
        run_bar = ctk.CTkFrame(tab, fg_color="transparent")
        run_bar.grid(row=row, column=0, sticky="ew", pady=(6, 6))
        run_btn = ctk.CTkButton(
            run_bar, text=button_text, width=170, height=40,
            font=ctk.CTkFont(size=15, weight="bold"),
            fg_color=ACCENT, hover_color=ACCENT_HOVER, command=command,
        )
        run_btn.pack(side="left")
        progress = ctk.CTkProgressBar(run_bar)
        progress.set(0)
        progress.pack(side="left", fill="x", expand=True, padx=(14, 0))

        log_box = ctk.CTkTextbox(tab, wrap="word", font=ctk.CTkFont(size=12))
        log_box.grid(row=row + 1, column=0, sticky="nsew", pady=(6, 4))
        log_box.configure(state="disabled")
        return run_btn, progress, log_box

    # -------------------------------------------------------- find & replace UI
    def build_find_replace_tab(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)  # file list grows
        tab.grid_rowconfigure(5, weight=1)  # log grows

        self.fr_files = FileSelector(tab)
        self.fr_files.grid(row=0, column=0, sticky="nsew", pady=(8, 10))

        io_frame = ctk.CTkFrame(tab, fg_color="transparent")
        io_frame.grid(row=1, column=0, sticky="ew")
        io_frame.grid_columnconfigure(0, weight=1)
        io_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(io_frame, text="Find").grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(io_frame, text="Replace with").grid(
            row=0, column=1, sticky="w", padx=(10, 0)
        )
        self.fr_find = ctk.CTkTextbox(io_frame, height=110, wrap="word")
        self.fr_find.grid(row=1, column=0, sticky="nsew", pady=(2, 0))
        self.fr_replace = ctk.CTkTextbox(io_frame, height=110, wrap="word")
        self.fr_replace.grid(row=1, column=1, sticky="nsew", padx=(10, 0), pady=(2, 0))

        opts = ctk.CTkFrame(tab, fg_color="transparent")
        opts.grid(row=2, column=0, sticky="ew", pady=(10, 4))
        self.fr_match_case = ctk.CTkSwitch(opts, text="Match case")
        self.fr_match_case.pack(side="left")
        self.fr_scope = ctk.StringVar(value="everywhere")
        ctk.CTkLabel(opts, text="Scope:").pack(side="left", padx=(20, 6))
        ctk.CTkOptionMenu(
            opts, width=180, values=["everywhere", "body"], variable=self.fr_scope,
            fg_color="#3a3d44", button_color="#3a3d44", button_hover_color="#4a4e57",
        ).pack(side="left")
        ctk.CTkLabel(
            opts, text="Originals are backed up to .bak before overwriting.",
            text_color=GREY, font=ctk.CTkFont(size=12),
        ).pack(side="right")

        self.fr_run, self.fr_progress, self.fr_log = self._build_run_panel(
            tab, 4, "Run replace", self._start_text_run
        )
        self._log(self.fr_log,
                  "Add .docx files, enter your Find and Replace text, then Run.")

    # --------------------------------------------------------- replace image UI
    def build_image_replace_tab(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(2, weight=1)  # file list grows
        tab.grid_rowconfigure(5, weight=1)  # log grows

        pickers = ctk.CTkFrame(tab, fg_color="transparent")
        pickers.grid(row=0, column=0, sticky="ew", pady=(8, 6))
        pickers.grid_columnconfigure(0, weight=1)
        pickers.grid_columnconfigure(1, weight=0)
        pickers.grid_columnconfigure(2, weight=1)
        self.im_old = ImagePicker(pickers, "Old image (find)")
        self.im_old.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ctk.CTkLabel(pickers, text="→", font=ctk.CTkFont(size=28)).grid(
            row=0, column=1, padx=8
        )
        self.im_new = ImagePicker(pickers, "New image (replace with)")
        self.im_new.grid(row=0, column=2, sticky="ew", padx=(6, 0))

        opts = ctk.CTkFrame(tab, fg_color="transparent")
        opts.grid(row=1, column=0, sticky="ew", pady=(8, 4))
        self.im_dims = ctk.CTkSwitch(
            opts, text="Also match by size if not an exact match"
        )
        self.im_dims.pack(side="left")
        ctk.CTkLabel(
            opts,
            text="Matches images identical to the old file; new image keeps the "
                 "original's size. Originals backed up to .bak.",
            text_color=GREY, font=ctk.CTkFont(size=12), wraplength=380, justify="right",
        ).pack(side="right")

        self.im_files = FileSelector(tab)
        self.im_files.grid(row=2, column=0, sticky="nsew", pady=(8, 10))

        self.im_run, self.im_progress, self.im_log = self._build_run_panel(
            tab, 4, "Replace image", self._start_image_run
        )
        self._log(self.im_log,
                  "Choose the old and new image, add .docx files, then Replace.")

    # ----------------------------------------------------------- run: text tab
    def _start_text_run(self):
        if self._running:
            return
        find = self.fr_find.get("1.0", "end-1c")
        replace = self.fr_replace.get("1.0", "end-1c")
        if not self.fr_files.files:
            self._log(self.fr_log, "⚠ No files selected.")
            return
        if not find:
            self._log(self.fr_log, "⚠ The Find box is empty.")
            return

        self._set_running(self.fr_run, True, "Running…")
        self.fr_progress.set(0)
        self._log(self.fr_log, f"\n── Starting: {len(self.fr_files.files)} file(s) ──")
        threading.Thread(
            target=self._text_worker, daemon=True,
            kwargs=dict(
                files=list(self.fr_files.files), find=find, replace=replace,
                match_case=bool(self.fr_match_case.get()), scope=self.fr_scope.get(),
            ),
        ).start()

    def _text_worker(self, files, find, replace, match_case, scope):
        total = len(files)
        grand = changed = 0
        for i, path in enumerate(files, start=1):
            name = os.path.basename(path)
            document, result = replace_in_document(
                path, find, replace, match_case=match_case, scope=scope
            )
            if not result.ok:
                self._logq(self.fr_log, f"✖ {name}: {result.error}")
            elif result.replacements == 0:
                self._logq(self.fr_log, f"•  {name}: no matches")
            else:
                saved = self._save_with_backup(path, document, self.fr_log, name,
                                               f"{result.replacements} replaced")
                if saved:
                    loc = ", ".join(f"{k}: {v}" for k, v in result.locations.items())
                    self._logq(self.fr_log,
                               f"    ({loc})")
                    grand += result.replacements
                    changed += 1
            self._progq(self.fr_progress, i / total)
        self._logq(self.fr_log,
                   f"── Done: {grand} replacement(s) across {changed} file(s). ──")
        self._doneq(self.fr_run, "Run replace")

    # ---------------------------------------------------------- run: image tab
    def _start_image_run(self):
        if self._running:
            return
        if not self.im_old.valid() or not self.im_new.valid():
            self._log(self.im_log, "⚠ Choose both an old and a new image.")
            return
        if not self.im_files.files:
            self._log(self.im_log, "⚠ No files selected.")
            return

        self._set_running(self.im_run, True, "Working…")
        self.im_progress.set(0)
        self._log(self.im_log, f"\n── Starting: {len(self.im_files.files)} file(s) ──")
        threading.Thread(
            target=self._image_worker, daemon=True,
            kwargs=dict(
                files=list(self.im_files.files),
                old_bytes=self.im_old.bytes, new_bytes=self.im_new.bytes,
                match_dims=bool(self.im_dims.get()),
            ),
        ).start()

    def _image_worker(self, files, old_bytes, new_bytes, match_dims):
        total = len(files)
        grand = changed = 0
        for i, path in enumerate(files, start=1):
            name = os.path.basename(path)
            document, result = replace_images_in_document(
                path, old_bytes, new_bytes, match_by_dimensions=match_dims
            )
            if not result.ok:
                self._logq(self.im_log, f"✖ {name}: {result.error}")
            elif result.total == 0:
                self._logq(self.im_log, f"•  {name}: no matching image")
            else:
                parts = []
                if result.exact:
                    parts.append(f"{result.exact} exact")
                if result.dimensions:
                    parts.append(f"{result.dimensions} by size")
                detail = ", ".join(parts)
                if result.format_changed:
                    detail += ", format updated"
                saved = self._save_with_backup(path, document, self.im_log, name,
                                               f"{result.total} image(s) — {detail}")
                if saved:
                    grand += result.total
                    changed += 1
            self._progq(self.im_progress, i / total)
        self._logq(self.im_log,
                   f"── Done: {grand} image(s) replaced across {changed} file(s). ──")
        self._doneq(self.im_run, "Replace image")

    # ---------------------------------------------------- shared save + backup
    def _save_with_backup(self, path, document, log_box, name, summary):
        """Back up original to .bak (once) then overwrite. Returns True on
        success; logs the outcome either way."""
        try:
            bak = path + ".bak"
            if not os.path.exists(bak):
                shutil.copy2(path, bak)
                note = "backed up"
            else:
                note = ".bak already existed, kept original"
            document.save(path)
            self._logq(log_box, f"✔ {name}: {summary} — {note}")
            return True
        except PermissionError:
            self._logq(log_box, f"✖ {name}: could not save (is it open in Word?)")
        except Exception as exc:
            self._logq(log_box, f"✖ {name}: save failed — {exc}")
        return False

    # -------------------------------------------------- thread-safe UI helpers
    def _ui(self, fn):
        self._msg_queue.put(fn)

    def _logq(self, log_box, text):
        self._ui(lambda: self._log(log_box, text))

    def _progq(self, progress, value):
        self._ui(lambda: progress.set(value))

    def _doneq(self, run_btn, idle_text):
        def finish():
            self._running = False
            run_btn.configure(state="normal", text=idle_text)
        self._ui(finish)

    def _drain_queue(self):
        try:
            while True:
                self._msg_queue.get_nowait()()
        except queue.Empty:
            pass
        self.after(100, self._drain_queue)

    def _set_running(self, run_btn, running, busy_text):
        self._running = running
        run_btn.configure(state="disabled" if running else "normal",
                          text=busy_text)

    def _log(self, log_box, text):
        log_box.configure(state="normal")
        log_box.insert("end", text + "\n")
        log_box.see("end")
        log_box.configure(state="disabled")


if __name__ == "__main__":
    WordHelperApp().mainloop()
