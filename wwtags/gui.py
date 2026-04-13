import sys
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
from importlib.metadata import version, PackageNotFoundError
from pathlib import Path
from PIL import Image, ImageTk


def get_version():
    try:
        return version("wwtags")
    except PackageNotFoundError:
        return "unknown"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("wwtags – Tag Import Generator")
        self.resizable(True, True)
        self.minsize(560, 480)
        logo_path = Path(__file__).parent.parent / "Official Quad Plus Brand Logo.png"
        if logo_path.exists():
            img = Image.open(logo_path)
            h = 64
            w = round(img.width * h / img.height)
            img = img.resize((w, h), Image.LANCZOS)
            self._logo = ImageTk.PhotoImage(img)
            self.iconphoto(True, self._logo)
        self._action_buttons = []
        self._build_ui()

    def _build_ui(self):
        main = ttk.Frame(self, padding=10)
        main.grid(row=0, column=0, sticky="nsew")
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        main.columnconfigure(1, weight=1)

        r = 0

        # Header
        ttk.Label(main, text=f"wwtags v{get_version()}", font=("", 10, "bold")).grid(
            row=r, column=0, columnspan=3, sticky="w", pady=(0, 8)
        )
        r += 1

        # Workbook
        ttk.Label(main, text="Workbook:").grid(row=r, column=0, sticky="w", pady=2)
        self.workbook_var = tk.StringVar()
        ttk.Entry(main, textvariable=self.workbook_var).grid(
            row=r, column=1, sticky="ew", padx=(4, 4)
        )
        ttk.Button(main, text="Browse…", command=self._browse_workbook).grid(row=r, column=2)
        r += 1

        # Output
        ttk.Label(main, text="Output:").grid(row=r, column=0, sticky="w", pady=2)
        self.output_var = tk.StringVar(value="ww_tag_import.csv")
        ttk.Entry(main, textvariable=self.output_var).grid(
            row=r, column=1, sticky="ew", padx=(4, 4)
        )
        ttk.Button(main, text="Browse…", command=self._browse_output).grid(row=r, column=2)
        r += 1

        ttk.Separator(main, orient="horizontal").grid(
            row=r, column=0, columnspan=3, sticky="ew", pady=6
        )
        r += 1

        # Checkboxes
        opts = ttk.Frame(main)
        opts.grid(row=r, column=0, columnspan=3, sticky="w")
        self.strict_var = tk.BooleanVar()
        self.dry_run_var = tk.BooleanVar()
        ttk.Checkbutton(opts, text="Strict", variable=self.strict_var).pack(
            side="left", padx=(0, 12)
        )
        ttk.Checkbutton(opts, text="Dry run", variable=self.dry_run_var).pack(side="left")
        r += 1

        # Filter
        ttk.Label(main, text="Filter:").grid(row=r, column=0, sticky="w", pady=2)
        self.filter_var = tk.StringVar()
        ttk.Entry(main, textvariable=self.filter_var).grid(
            row=r, column=1, sticky="ew", padx=(4, 4)
        )
        ttk.Label(main, text="COL=VAL", foreground="gray").grid(row=r, column=2, sticky="w")
        r += 1

        ttk.Separator(main, orient="horizontal").grid(
            row=r, column=0, columnspan=3, sticky="ew", pady=6
        )
        r += 1

        # Action buttons
        btns = ttk.Frame(main)
        btns.grid(row=r, column=0, columnspan=3, sticky="w")
        self.generate_btn = ttk.Button(btns, text="Generate", command=self._run_generate)
        self.generate_btn.pack(side="left", padx=(0, 8))
        self.list_templates_btn = ttk.Button(
            btns, text="List Templates", command=self._run_list_templates
        )
        self.list_templates_btn.pack(side="left", padx=(0, 8))
        self.list_columns_btn = ttk.Button(
            btns, text="List Columns", command=self._run_list_columns
        )
        self.list_columns_btn.pack(side="left")
        self._action_buttons = [
            self.generate_btn,
            self.list_templates_btn,
            self.list_columns_btn,
        ]
        r += 1

        ttk.Separator(main, orient="horizontal").grid(
            row=r, column=0, columnspan=3, sticky="ew", pady=6
        )
        r += 1

        # Log panel
        ttk.Label(main, text="Log:").grid(row=r, column=0, columnspan=3, sticky="w")
        r += 1
        self.log = scrolledtext.ScrolledText(main, height=12, state="disabled", wrap="word")
        self.log.grid(row=r, column=0, columnspan=3, sticky="nsew")
        main.rowconfigure(r, weight=1)

        self.workbook_var.trace_add("write", self._on_workbook_change)
        self._on_workbook_change()

    # ------------------------------------------------------------------
    # File dialogs
    # ------------------------------------------------------------------

    def _browse_workbook(self):
        path = filedialog.askopenfilename(
            title="Select Workbook",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")],
        )
        if path:
            self.workbook_var.set(path)

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            title="Save Output As",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialfile=self.output_var.get(),
        )
        if path:
            self.output_var.set(path)

    # ------------------------------------------------------------------
    # Button state management
    # ------------------------------------------------------------------

    def _on_workbook_change(self, *_):
        state = "normal" if self.workbook_var.get().strip() else "disabled"
        for btn in self._action_buttons:
            btn.config(state=state)

    def _set_busy(self, busy):
        if busy:
            state = "disabled"
        else:
            state = "normal" if self.workbook_var.get().strip() else "disabled"
        for btn in self._action_buttons:
            btn.config(state=state)

    # ------------------------------------------------------------------
    # Command builders
    # ------------------------------------------------------------------

    def _generate_cmd(self):
        cmd = [sys.executable, "-m", "wwtags.cli", self.workbook_var.get().strip()]
        output = self.output_var.get().strip()
        if output:
            cmd += ["--output", output]
        if self.strict_var.get():
            cmd.append("--strict")
        if self.dry_run_var.get():
            cmd.append("--dry-run")
        filter_val = self.filter_var.get().strip()
        if filter_val:
            cmd += ["--filter", filter_val]
        return cmd

    def _workbook_only_cmd(self, flag):
        return [sys.executable, "-m", "wwtags.cli", self.workbook_var.get().strip(), flag]

    # ------------------------------------------------------------------
    # Button handlers
    # ------------------------------------------------------------------

    def _run_generate(self):
        self._run_cmd(self._generate_cmd())

    def _run_list_templates(self):
        self._run_cmd(self._workbook_only_cmd("--list-templates"))

    def _run_list_columns(self):
        self._run_cmd(self._workbook_only_cmd("--list-columns"))

    def _run_cmd(self, cmd):
        self._set_busy(True)
        self._clear_log()
        # Show a readable summary (skip python path and -m flag)
        display = "wwtags " + " ".join(cmd[3:])
        self._append_log(f"$ {display}\n\n")
        threading.Thread(target=self._exec, args=(cmd,), daemon=True).start()

    # ------------------------------------------------------------------
    # Subprocess execution (runs in background thread)
    # ------------------------------------------------------------------

    def _exec(self, cmd):
        try:
            proc = subprocess.run(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
            )
            self.after(0, self._append_log, proc.stdout or "")
            if proc.returncode != 0:
                self.after(0, self._append_log, f"\n[Exited with code {proc.returncode}]\n")
        except Exception as e:
            self.after(0, self._append_log, f"Failed to run command: {e}\n")
        finally:
            self.after(0, self._set_busy, False)

    # ------------------------------------------------------------------
    # Log helpers
    # ------------------------------------------------------------------

    def _clear_log(self):
        self.log.config(state="normal")
        self.log.delete("1.0", "end")
        self.log.config(state="disabled")

    def _append_log(self, text):
        self.log.config(state="normal")
        self.log.insert("end", text)
        self.log.see("end")
        self.log.config(state="disabled")


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
