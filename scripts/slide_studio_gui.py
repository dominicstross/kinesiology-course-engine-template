import os
import sys
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Any, Dict, List

SCRIPT_DIR = os.path.abspath(os.path.dirname(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

import build_slides  # type: ignore


class SlideStudioApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Kinesiology OLM Slide Studio - PPTX to JSON Converter")
        self.root.geometry("1100x670")
        self.root.minsize(920, 560)

        template_root = os.path.dirname(SCRIPT_DIR)
        default_workspace = os.path.dirname(template_root)

        self.workspace_var = tk.StringVar(value=default_workspace)
        self.filter_var = tk.StringVar(value="ALL")
        self.text_only_var = tk.BooleanVar(value=False)
        self.force_images_var = tk.BooleanVar(value=False)
        self.format_var = tk.StringVar(value="webp")

        self.decks: List[Dict[str, Any]] = []
        self.is_running = False

        self._build_ui()
        self.refresh_decks()
        self.root.bind("<FocusIn>", self._on_window_focus)

    def _on_window_focus(self, event: Any) -> None:
        if event.widget == self.root and not self.is_running:
            self.refresh_decks()

    def _build_ui(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        top_frame = ttk.Frame(self.root, padding=(12, 10))
        top_frame.pack(fill=tk.X)

        ttk.Label(top_frame, text="Course Repos Folder:", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT)
        ttk.Entry(top_frame, textvariable=self.workspace_var, width=65).pack(side=tk.LEFT, padx=(8, 6), fill=tk.X, expand=True)
        ttk.Button(top_frame, text="Browse...", command=self.browse_folder).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(top_frame, text="Rescan", command=self.refresh_decks).pack(side=tk.LEFT)

        filter_frame = ttk.Frame(self.root, padding=(12, 4))
        filter_frame.pack(fill=tk.X)

        ttk.Label(filter_frame, text="Show:").pack(side=tk.LEFT, padx=(0, 6))
        filter_options = [
            ("All Decks", "ALL"),
            ("New & Modified Only", "ACTIONABLE"),
            ("Missing Notes Only", "MISSING_NOTES"),
            ("Up to Date", "UP TO DATE"),
        ]
        for label, val in filter_options:
            ttk.Radiobutton(
                filter_frame,
                text=label,
                variable=self.filter_var,
                value=val,
                command=self._populate_tree,
            ).pack(side=tk.LEFT, padx=(0, 12))

        self.summary_label = ttk.Label(filter_frame, text="", font=("Segoe UI", 9, "bold"), foreground="#1d4ed8")
        self.summary_label.pack(side=tk.RIGHT)

        tree_frame = ttk.Frame(self.root, padding=(12, 6))
        tree_frame.pack(fill=tk.BOTH, expand=True)

        columns = ("status", "repo", "chapter", "module", "file", "pngs", "notes")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="extended")
        self.tree.heading("status", text="Status")
        self.tree.heading("repo", text="Repository")
        self.tree.heading("chapter", text="Chapter")
        self.tree.heading("module", text="Module")
        self.tree.heading("file", text="PowerPoint File")
        self.tree.heading("pngs", text="Exported Images")
        self.tree.heading("notes", text="Speaker Notes")

        self.tree.column("status", width=120, anchor=tk.CENTER)
        self.tree.column("repo", width=180, anchor=tk.W)
        self.tree.column("chapter", width=80, anchor=tk.CENTER)
        self.tree.column("module", width=75, anchor=tk.CENTER)
        self.tree.column("file", width=250, anchor=tk.W)
        self.tree.column("pngs", width=115, anchor=tk.CENTER)
        self.tree.column("notes", width=115, anchor=tk.CENTER)

        self.tree.tag_configure("NEW", background="#fef3c7", foreground="#92400e")
        self.tree.tag_configure("NEEDS UPDATE", background="#dbeafe", foreground="#1e40af")
        self.tree.tag_configure("UP TO DATE", background="#f0fdf4", foreground="#166534")

        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        ctrl_frame = ttk.Frame(self.root, padding=(12, 8))
        ctrl_frame.pack(fill=tk.X)

        self.btn_convert_smart = ttk.Button(
            ctrl_frame,
            text="Convert New & Modified Decks",
            command=self.run_new_and_modified,
        )
        self.btn_convert_smart.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_convert_selected = ttk.Button(
            ctrl_frame,
            text="Convert Selected",
            command=self.run_selected,
        )
        self.btn_convert_selected.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_refresh_downloads = ttk.Button(
            ctrl_frame,
            text="Refresh Downloads",
            command=self.run_refresh_downloads,
        )
        self.btn_refresh_downloads.pack(side=tk.LEFT, padx=(0, 12))

        ttk.Label(ctrl_frame, text="Format:").pack(side=tk.LEFT, padx=(0, 4))
        fmt_combo = ttk.Combobox(
            ctrl_frame,
            textvariable=self.format_var,
            values=["webp", "png"],
            width=6,
            state="readonly",
        )
        fmt_combo.pack(side=tk.LEFT, padx=(0, 12))

        ttk.Checkbutton(
            ctrl_frame,
            text="Text / Notes Only",
            variable=self.text_only_var,
            command=self._on_mode_toggle,
        ).pack(side=tk.LEFT, padx=(0, 10))

        ttk.Checkbutton(
            ctrl_frame,
            text="Force Re-Render Images",
            variable=self.force_images_var,
            command=self._on_force_toggle,
        ).pack(side=tk.LEFT)

        log_frame = ttk.LabelFrame(self.root, text="Conversion Log", padding=(8, 6))
        log_frame.pack(fill=tk.BOTH, expand=False, padx=12, pady=(0, 12))

        self.log_text = tk.Text(log_frame, height=10, wrap=tk.WORD, font=("Consolas", 9), bg="#0f172a", fg="#e2e8f0")
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def _on_mode_toggle(self) -> None:
        if self.text_only_var.get():
            self.force_images_var.set(False)

    def _on_force_toggle(self) -> None:
        if self.force_images_var.get():
            self.text_only_var.set(False)

    def browse_folder(self) -> None:
        chosen = filedialog.askdirectory(initialdir=self.workspace_var.get(), title="Select Course Repos or Chapter Folder")
        if chosen:
            self.workspace_var.set(chosen)
            self.refresh_decks()

    def append_log(self, message: str) -> None:
        def _insert() -> None:
            self.log_text.insert(tk.END, message + "\n")
            self.log_text.see(tk.END)
        self.root.after(0, _insert)

    def refresh_decks(self) -> None:
        target = os.path.normpath(self.workspace_var.get().strip())
        self.workspace_var.set(target)
        if not os.path.exists(target):
            self.append_log(f"[WARN] Folder does not exist: {target}")
            return
        try:
            self.decks = build_slides.scan_decks(target)
        except Exception as exc:
            self.append_log(f"[ERROR] Scan failed: {exc}")
            self.decks = []
        self._populate_tree()

    def _populate_tree(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)

        mode = self.filter_var.get()
        new_count = sum(1 for d in self.decks if d["status"] == "NEW")
        mod_count = sum(1 for d in self.decks if d["status"] == "NEEDS UPDATE")
        ok_count = sum(1 for d in self.decks if d["status"] == "UP TO DATE")
        no_notes_count = sum(1 for d in self.decks if d.get("notes_count", 0) < d.get("total_slides", 0))

        self.summary_label.config(
            text=f"Total: {len(self.decks)}  |  New: {new_count}  |  Modified: {mod_count}  |  Need Notes: {no_notes_count}  |  Up to Date: {ok_count}"
        )

        for idx, d in enumerate(self.decks):
            status = d["status"]
            n_cnt = d.get("notes_count", 0)
            t_cnt = d.get("total_slides", 0)

            if mode == "ACTIONABLE" and status == "UP TO DATE":
                continue
            if mode == "UP TO DATE" and status != "UP TO DATE":
                continue
            if mode == "MISSING_NOTES" and t_cnt > 0 and n_cnt >= t_cnt:
                continue

            repo_name = os.path.basename(d["repo_root"])
            notes_label = f"{n_cnt} / {t_cnt}" if t_cnt > 0 else "0 / ?"

            self.tree.insert(
                "",
                tk.END,
                iid=str(idx),
                values=(
                    status,
                    repo_name,
                    d["chapter"],
                    d["mod_id"],
                    d["filename"],
                    f"{d['png_count']} slides" if d["png_count"] else "None",
                    notes_label,
                ),
                tags=(status,),
            )

    def _set_busy(self, busy: bool) -> None:
        self.is_running = busy
        state = tk.DISABLED if busy else tk.NORMAL
        self.btn_convert_smart.config(state=state)
        self.btn_convert_selected.config(state=state)

    def run_new_and_modified(self) -> None:
        if self.is_running:
            return
        targets = [d for d in self.decks if d["status"] in ("NEW", "NEEDS UPDATE")]
        if not targets:
            messagebox.showinfo("All Up to Date", "All scanned PowerPoint decks are already converted and up to date!")
            return
        self._start_worker(targets)

    def run_refresh_downloads(self) -> None:
        from build_downloads import build_downloads

        roots = set()
        for d in getattr(self, "decks", []):
            if d.get("repo_root"):
                roots.add(os.path.abspath(d["repo_root"]))

        ws = os.path.abspath(self.workspace_var.get().strip())
        if not roots and os.path.isdir(ws):
            if os.path.isdir(os.path.join(ws, "content")):
                roots.add(ws)
            else:
                for entry in os.listdir(ws):
                    cand = os.path.join(ws, entry)
                    if os.path.isdir(os.path.join(cand, "content")):
                        roots.add(cand)

        if not roots:
            messagebox.showwarning("No Course Repo Found", "No course repository containing a 'content' folder was found under the selected path.")
            return

        updated = []
        for r in sorted(roots):
            try:
                build_downloads(r)
                repo_name = os.path.basename(r)
                updated.append(repo_name)
                if hasattr(self, "log"):
                    self.log(f"[DOWNLOADS] Updated {repo_name}/downloads/manifest.json")
                elif hasattr(self, "append_log"):
                    self.append_log(f"[DOWNLOADS] Updated {repo_name}/downloads/manifest.json")
            except Exception as exc:
                messagebox.showerror("Downloads Error", f"Failed to build downloads for {r}:\n{exc}")
                return

        messagebox.showinfo("Downloads Refreshed", "Updated downloads/manifest.json for:\n• " + "\n• ".join(updated))

    def run_selected(self) -> None:
        if self.is_running:
            return
        selected_ids = self.tree.selection()
        if not selected_ids:
            messagebox.showwarning("No Selection", "Select one or more decks in the table first (or click 'Convert New & Modified Decks').")
            return
        targets = [self.decks[int(i)] for i in selected_ids]
        self._start_worker(targets)

    def _start_worker(self, targets: List[Dict[str, Any]]) -> None:
        self._set_busy(True)
        text_only = self.text_only_var.get()
        force_images = self.force_images_var.get()
        img_format = self.format_var.get()

        def worker() -> None:
            self.append_log(f"\n=== Starting batch conversion ({len(targets)} deck(s), format={img_format.upper()}) ===")
            for d in targets:
                try:
                    build_slides.convert_deck(
                        d["pptx_path"],
                        repo_root=d["repo_root"],
                        force_images=force_images,
                        text_only=text_only,
                        img_format=img_format,
                        log_fn=self.append_log,
                    )
                except Exception as exc:
                    self.append_log(f"[ERROR] Failed on {d['filename']}: {exc}")
            self.append_log("=== Batch conversion complete ===")
            self.root.after(0, self._on_worker_done)

        threading.Thread(target=worker, daemon=True).start()

    def _on_worker_done(self) -> None:
        self._set_busy(False)
        self.refresh_decks()


def main() -> None:
    try:
        root = tk.Tk()
        root.lift()
        root.attributes("-topmost", True)
        root.after(300, lambda: root.attributes("-topmost", False))
        SlideStudioApp(root)
        root.mainloop()
    except Exception:
        traceback.print_exc()
        input("\nPress Enter to exit...")


if __name__ == "__main__":
    main()
