import sys
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
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
        self.title("wwtags – Wonderware Tag Import Generator")
        self.resizable(True, True)
        self.minsize(560, 480)
        self._header_logo = None
        header_path = Path(__file__).parent / "qp_logo.png"
        if header_path.exists():
            img = Image.open(header_path)
            h = 40
            w = round(img.width * h / img.height)
            self._header_logo = ImageTk.PhotoImage(img.resize((w, h), Image.LANCZOS))

        icon_path = Path(__file__).parent / "qp_logo_icon.png"
        if icon_path.exists():
            try:
                icon_img = Image.open(icon_path)
                self._icons = [
                    ImageTk.PhotoImage(icon_img.resize((s, s), Image.LANCZOS))
                    for s in (16, 32, 48)
                ]
                self.iconphoto(True, *self._icons)
            except Exception:
                pass
        self._action_buttons = []
        self._build_menu()
        self._build_ui()

    def _build_ui(self):
        main = ttk.Frame(self, padding=10)
        main.grid(row=0, column=0, sticky="nsew")
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        main.columnconfigure(1, weight=1)

        r = 0

        # Header
        header = ttk.Frame(main)
        header.grid(row=r, column=0, columnspan=3, sticky="w", pady=(0, 8))
        if self._header_logo:
            tk.Label(header, image=self._header_logo, borderwidth=0).pack(side="left", padx=(0, 8))
        # ttk.Label(header, text=f"wwtags v{get_version()}", font=("", 10)).pack(
        #     side="left"
        # )
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
        ttk.Label(main, text="Output File:").grid(row=r, column=0, sticky="w", pady=2)
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
        self.dry_run_var = tk.BooleanVar()
        ttk.Checkbutton(opts, text="Dry run", variable=self.dry_run_var).pack(side="left", padx=(0, 16))
        self.ww2023_var = tk.BooleanVar()
        ttk.Checkbutton(opts, text="Wonderware 2023+", variable=self.ww2023_var).pack(side="left")
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
        self.list_columns_btn.pack(side="left", padx=(0, 8))
        self.tag_length_btn = ttk.Button(
            btns, text="Check Tag Length", command=self._run_tag_length
        )
        self.tag_length_btn.pack(side="left")
        self._action_buttons = [
            self.generate_btn,
            self.list_templates_btn,
            self.list_columns_btn,
            self.tag_length_btn,
        ]
        r += 1

        ttk.Separator(main, orient="horizontal").grid(
            row=r, column=0, columnspan=3, sticky="ew", pady=6
        )
        r += 1

        # UDT import
        ttk.Label(main, text="UDT (.L5X):").grid(row=r, column=0, sticky="w", pady=2)
        self.l5x_var = tk.StringVar()
        ttk.Entry(main, textvariable=self.l5x_var).grid(
            row=r, column=1, sticky="ew", padx=(4, 4)
        )
        ttk.Button(main, text="Browse…", command=self._browse_l5x).grid(row=r, column=2)
        r += 1

        udt_btns = ttk.Frame(main)
        udt_btns.grid(row=r, column=0, columnspan=3, sticky="w")
        self.import_udt_btn = ttk.Button(
            udt_btns, text="Import UDT", command=self._run_import_udt
        )
        self.import_udt_btn.pack(side="left")
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
        self.l5x_var.trace_add("write", self._on_l5x_change)
        self._on_workbook_change()

    # ------------------------------------------------------------------
    # Menu bar
    # ------------------------------------------------------------------

    def _build_menu(self):
        menubar = tk.Menu(self)
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="How to Use", command=self._show_help)
        help_menu.add_separator()
        help_menu.add_command(label="About", command=self._show_about)
        menubar.add_cascade(label="Help", menu=help_menu)
        self.config(menu=menubar)

    def _show_help(self):
        win = tk.Toplevel(self)
        win.title("How to Use — wwtags")
        win.resizable(True, True)
        win.minsize(480, 420)

        text = scrolledtext.ScrolledText(win, wrap="word", padx=12, pady=10)
        text.pack(fill="both", expand=True)

        help_text = (
            "wwtags — Wonderware Tag Import Generator\n"
            "========================================\n\n"
            "OVERVIEW\n"
            "--------\n"
            "This tool reads an Excel workbook containing a DEVICE_LIST sheet\n"
            "and one or more template sheets, then generates a CSV file\n"
            "ready to import into Wonderware.\n\n"
            "FIELDS\n"
            "------\n"
            "Workbook\n"
            "  The Excel workbook to process.\n"
            "  Click Browse… to select the file.\n\n"
            "Output\n"
            "  The name and location of the generated CSV file.\n"
            "  Default: ww_tag_import.csv (saved to the current directory).\n"
            "  Click Browse… to choose a different name or location.\n\n"
            "Dry run\n"
            "  When checked, validates the workbook and counts tags without\n"
            "  writing any output file. Use this to check for errors before\n"
            "  generating.\n\n"
            "Filter\n"
            "  Process only rows where a column matches a specific value.\n"
            "  Enter in COL=VAL format, for example:\n"
            "    DEVICE_TYPE=VFD\n"
            "    ALARM_GROUP=Pumps\n"
            "  Leave blank to process all rows.\n\n"
            "BUTTONS\n"
            "-------\n"
            "Generate\n"
            "  Runs the tag export with the current settings and writes the\n"
            "  output CSV file (unless Dry run is checked).\n\n"
            "List Templates\n"
            "  Lists all template sheets found in the workbook. Use this to\n"
            "  confirm sheet names match the DEVICE_TYPE values in\n"
            "  DEVICE_LIST.\n\n"
            "List Columns\n"
            "  Lists the columns present in DEVICE_LIST, showing which are\n"
            "  required, optional, or unrecognised.\n\n"
            "Check Tag Length\n"
            "  Checks each template sheet and reports the maximum number of\n"
            "  characters the HMI_TAG value can be, given the fixed characters\n"
            "  surrounding it in each tag row. Prompts for Wonderware version\n"
            "  (2020 and earlier = 32-char limit; 2023 and later = 128-char\n"
            "  limit). Use this as a precheck before generating tags.\n\n"
            "UDT IMPORT\n"
            "----------\n"
            "UDT (.L5X)\n"
            "  A Studio5000 User Defined Data Type export file (.L5X).\n"
            "  Click Browse… to select the file.\n\n"
            "Import UDT\n"
            "  Parses the selected .L5X file and adds a new template sheet\n"
            "  to the workbook, named after the UDT (e.g. QP_DRV_ETH_04_v01).\n"
            "  Each UDT member becomes a tag row in the correct Wonderware\n"
            "  section (:IODisc, :IOReal, or :IOInt).\n"
            "  SINT packed-bit containers and TIMER members are excluded.\n"
            "  Errors if a sheet with the UDT name already exists.\n\n"
            "LOG PANEL\n"
            "---------\n"
            "All output — including tag counts, validation warnings, and\n"
            "errors — is shown in the log panel at the bottom of the window.\n"
        )

        text.insert("1.0", help_text)
        text.config(state="disabled")

    def _show_about(self):
        messagebox.showinfo(
            "About wwtags",
            f"wwtags v{get_version()}\n\nWonderware tag import generator.\n\n"
            "Generates tag import CSV files from Excel workbook templates.",
        )

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

    def _browse_l5x(self):
        path = filedialog.askopenfilename(
            title="Select UDT L5X File",
            filetypes=[("L5X files", "*.L5X *.l5x"), ("XML files", "*.xml"), ("All files", "*.*")],
        )
        if path:
            self.l5x_var.set(path)

    # ------------------------------------------------------------------
    # Button state management
    # ------------------------------------------------------------------

    def _on_workbook_change(self, *_):
        state = "normal" if self.workbook_var.get().strip() else "disabled"
        for btn in self._action_buttons:
            btn.config(state=state)
        self._update_import_udt_state()

    def _on_l5x_change(self, *_):
        self._update_import_udt_state()

    def _update_import_udt_state(self):
        enabled = bool(self.workbook_var.get().strip() and self.l5x_var.get().strip())
        self.import_udt_btn.config(state="normal" if enabled else "disabled")

    def _set_busy(self, busy):
        if busy:
            for btn in self._action_buttons:
                btn.config(state="disabled")
            self.import_udt_btn.config(state="disabled")
        else:
            state = "normal" if self.workbook_var.get().strip() else "disabled"
            for btn in self._action_buttons:
                btn.config(state=state)
            self._update_import_udt_state()

    # ------------------------------------------------------------------
    # Command builders
    # ------------------------------------------------------------------

    def _ww_version_args(self):
        return ["--ww-version", "2023" if self.ww2023_var.get() else "2020"]

    def _generate_cmd(self):
        cmd = [sys.executable, "-m", "wwtags.cli", self.workbook_var.get().strip()]
        output = self.output_var.get().strip()
        if output:
            cmd += ["--output", output]
        if self.dry_run_var.get():
            cmd.append("--dry-run")
        filter_val = self.filter_var.get().strip()
        if filter_val:
            cmd += ["--filter", filter_val]
        cmd += self._ww_version_args()
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

    def _run_tag_length(self):
        cmd = [
            sys.executable, "-m", "wwtags.cli",
            self.workbook_var.get().strip(),
            "--tag-length",
        ] + self._ww_version_args()
        self._run_cmd(cmd)

    def _run_import_udt(self):
        cmd = [
            sys.executable, "-m", "wwtags.cli",
            self.workbook_var.get().strip(),
            "--import-udt", self.l5x_var.get().strip(),
        ]
        self._run_cmd(cmd)

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
