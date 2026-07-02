"""Word Helper — a small dark-themed GUI for batch operations on .docx files.

Currently provides a Find & Replace tab that runs across multiple selected Word
documents with no 255-character limit. The tab layout is modular so more tabs
(each its own build_* method) can be added later.
"""

from __future__ import annotations

import os
import queue
import shutil
import threading
from tkinter import filedialog

import customtkinter as ctk

from docx_replace import replace_in_document

APP_NAME = "Word Helper"
ACCENT = "#3d7eff"
ACCENT_HOVER = "#2f66d6"

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")


class WordHelperApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("940x820")
        self.minsize(760, 640)

        self.files: list[str] = []
        self._msg_queue: queue.Queue = queue.Queue()
        self._running = False

        self._build_header()

        self.tabview = ctk.CTkTabview(self, fg_color="transparent")
        self.tabview.pack(fill="both", expand=True, padx=16, pady=(0, 12))

        self.tabview.add("Find & Replace")
        self.build_find_replace_tab(self.tabview.tab("Find & Replace"))

        # Future features get their own tab + build_* method, e.g.:
        # self.tabview.add("Merge Fields")
        # self.build_merge_tab(self.tabview.tab("Merge Fields"))

        self.after(100, self._drain_queue)

    # ------------------------------------------------------------------ header
    def _build_header(self):
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=16, pady=(14, 6))
        ctk.CTkLabel(
            header,
            text="Word Helper",
            font=ctk.CTkFont(size=24, weight="bold"),
        ).pack(side="left")
        ctk.CTkLabel(
            header,
            text="Batch tools for .docx files",
            font=ctk.CTkFont(size=13),
            text_color="#8a8f98",
        ).pack(side="left", padx=(12, 0), pady=(6, 0))

    # -------------------------------------------------------- find & replace UI
    def build_find_replace_tab(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=1)  # file list grows
        tab.grid_rowconfigure(5, weight=1)  # log grows

        # --- Files section ---
        files_bar = ctk.CTkFrame(tab, fg_color="transparent")
        files_bar.grid(row=0, column=0, sticky="ew", pady=(8, 4))
        ctk.CTkLabel(
            files_bar, text="Word files", font=ctk.CTkFont(size=15, weight="bold")
        ).pack(side="left")
        ctk.CTkButton(
            files_bar, text="Clear", width=70, fg_color="#3a3d44",
            hover_color="#4a4e57", command=self._clear_files,
        ).pack(side="right")
        ctk.CTkButton(
            files_bar, text="+ Add files", width=110, fg_color=ACCENT,
            hover_color=ACCENT_HOVER, command=self._add_files,
        ).pack(side="right", padx=(0, 8))

        self.file_frame = ctk.CTkScrollableFrame(tab, height=140,
                                                 label_text="")
        self.file_frame.grid(row=1, column=0, sticky="nsew", pady=(0, 10))
        self._render_file_list()

        # --- Find / Replace inputs ---
        io_frame = ctk.CTkFrame(tab, fg_color="transparent")
        io_frame.grid(row=2, column=0, sticky="ew")
        io_frame.grid_columnconfigure(0, weight=1)
        io_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(io_frame, text="Find").grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(io_frame, text="Replace with").grid(row=0, column=1,
                                                          sticky="w", padx=(10, 0))
        self.find_box = ctk.CTkTextbox(io_frame, height=110, wrap="word")
        self.find_box.grid(row=1, column=0, sticky="nsew", pady=(2, 0))
        self.replace_box = ctk.CTkTextbox(io_frame, height=110, wrap="word")
        self.replace_box.grid(row=1, column=1, sticky="nsew", padx=(10, 0),
                              pady=(2, 0))

        # --- Options ---
        opts = ctk.CTkFrame(tab, fg_color="transparent")
        opts.grid(row=3, column=0, sticky="ew", pady=(10, 4))
        self.match_case = ctk.CTkSwitch(opts, text="Match case")
        self.match_case.pack(side="left")

        self.scope_var = ctk.StringVar(value="everywhere")
        ctk.CTkLabel(opts, text="Scope:").pack(side="left", padx=(20, 6))
        ctk.CTkOptionMenu(
            opts, width=180,
            values=["everywhere", "body"],
            variable=self.scope_var,
            fg_color="#3a3d44", button_color="#3a3d44",
            button_hover_color="#4a4e57",
        ).pack(side="left")
        ctk.CTkLabel(
            opts,
            text="Originals are backed up to .bak before overwriting.",
            text_color="#8a8f98", font=ctk.CTkFont(size=12),
        ).pack(side="right")

        # --- Run + progress ---
        run_bar = ctk.CTkFrame(tab, fg_color="transparent")
        run_bar.grid(row=4, column=0, sticky="ew", pady=(6, 6))
        self.run_btn = ctk.CTkButton(
            run_bar, text="Run replace", width=160, height=40,
            font=ctk.CTkFont(size=15, weight="bold"),
            fg_color=ACCENT, hover_color=ACCENT_HOVER,
            command=self._start_run,
        )
        self.run_btn.pack(side="left")
        self.progress = ctk.CTkProgressBar(run_bar)
        self.progress.set(0)
        self.progress.pack(side="left", fill="x", expand=True, padx=(14, 0))

        # --- Log ---
        self.log_box = ctk.CTkTextbox(tab, wrap="word", font=ctk.CTkFont(size=12))
        self.log_box.grid(row=5, column=0, sticky="nsew", pady=(6, 4))
        self.log_box.configure(state="disabled")
        self._log("Add .docx files, enter your Find and Replace text, then Run.")

    # ---------------------------------------------------------------- file list
    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="Select Word documents",
            filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
        )
        for p in paths:
            if p not in self.files:
                self.files.append(p)
        self._render_file_list()

    def _clear_files(self):
        self.files.clear()
        self._render_file_list()

    def _remove_file(self, path):
        if path in self.files:
            self.files.remove(path)
        self._render_file_list()

    def _render_file_list(self):
        for child in self.file_frame.winfo_children():
            child.destroy()
        if not self.files:
            ctk.CTkLabel(
                self.file_frame, text="No files selected.",
                text_color="#8a8f98",
            ).pack(anchor="w", padx=6, pady=6)
            return
        for path in self.files:
            row = ctk.CTkFrame(self.file_frame, fg_color="#2b2d31")
            row.pack(fill="x", pady=2, padx=2)
            ctk.CTkButton(
                row, text="✕", width=28, height=24,
                fg_color="#3a3d44", hover_color="#a13d3d",
                command=lambda p=path: self._remove_file(p),
            ).pack(side="right", padx=4, pady=2)
            ctk.CTkLabel(
                row, text=os.path.basename(path), anchor="w",
            ).pack(side="left", padx=8, pady=2)

    # ----------------------------------------------------------------- run flow
    def _start_run(self):
        if self._running:
            return
        find = self.find_box.get("1.0", "end-1c")
        replace = self.replace_box.get("1.0", "end-1c")

        if not self.files:
            self._log("⚠ No files selected.")
            return
        if not find:
            self._log("⚠ The Find box is empty.")
            return

        self._running = True
        self.run_btn.configure(state="disabled", text="Running…")
        self.progress.set(0)
        self._log(f"\n── Starting: {len(self.files)} file(s) ──")

        args = dict(
            files=list(self.files),
            find=find,
            replace=replace,
            match_case=bool(self.match_case.get()),
            scope=self.scope_var.get(),
        )
        threading.Thread(target=self._worker, kwargs=args, daemon=True).start()

    def _worker(self, files, find, replace, match_case, scope):
        total = len(files)
        total_replacements = 0
        changed_files = 0
        for i, path in enumerate(files, start=1):
            name = os.path.basename(path)
            document, result = replace_in_document(
                path, find, replace, match_case=match_case, scope=scope
            )
            if not result.ok:
                self._post(("log", f"✖ {name}: {result.error}"))
            elif result.replacements == 0:
                self._post(("log", f"•  {name}: no matches"))
            else:
                try:
                    bak = path + ".bak"
                    if not os.path.exists(bak):
                        shutil.copy2(path, bak)
                        backup_note = "backed up"
                    else:
                        backup_note = ".bak already existed, kept original"
                    document.save(path)
                    loc = ", ".join(f"{k}: {v}" for k, v in result.locations.items())
                    self._post((
                        "log",
                        f"✔ {name}: {result.replacements} replaced ({loc}) — {backup_note}",
                    ))
                    total_replacements += result.replacements
                    changed_files += 1
                except PermissionError:
                    self._post((
                        "log",
                        f"✖ {name}: could not save (is it open in Word?)",
                    ))
                except Exception as exc:
                    self._post(("log", f"✖ {name}: save failed — {exc}"))
            self._post(("progress", i / total))

        self._post((
            "log",
            f"── Done: {total_replacements} replacement(s) across "
            f"{changed_files} file(s). ──",
        ))
        self._post(("done", None))

    # -------------------------------------------------- thread-safe UI updates
    def _post(self, msg):
        self._msg_queue.put(msg)

    def _drain_queue(self):
        try:
            while True:
                kind, payload = self._msg_queue.get_nowait()
                if kind == "log":
                    self._log(payload)
                elif kind == "progress":
                    self.progress.set(payload)
                elif kind == "done":
                    self._running = False
                    self.run_btn.configure(state="normal", text="Run replace")
        except queue.Empty:
            pass
        self.after(100, self._drain_queue)

    def _log(self, text):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", text + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")


if __name__ == "__main__":
    WordHelperApp().mainloop()
