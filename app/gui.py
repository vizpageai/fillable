from __future__ import annotations

import os
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from app.codex_cli import CodexCli, CodexCliError
from app.template_engine import (
    create_template_from_user_template,
    fill_template,
    fill_template_multiple,
    generate_template,
)
from app.utils import AppConfig, load_config, load_json, save_config
from app.version import APP_NAME, APP_VERSION


class FillableApp(tk.Tk):
    def __init__(self, initial_generate: str | None = None, initial_template: str | None = None):
        super().__init__()
        self.title(f"{APP_NAME} {APP_VERSION} - Codex Document Autofill")
        self.geometry("1080x760")
        self.minsize(960, 700)

        self.config_data = load_config()
        self.source_var = tk.StringVar(value=initial_generate or "")
        self.template_var = tk.StringVar(value=initial_template or "")
        self.source_has_placeholders_var = tk.BooleanVar(value=False)
        self.batch_data_var = tk.StringVar(value="")
        self.codex_command_var = tk.StringVar(value=self.config_data.codex_command_template)

        self._init_styles()
        self._build_ui()

    def _init_styles(self) -> None:
        self.colors = {
            "bg": "#F4F7FB",
            "card": "#FFFFFF",
            "text": "#0F172A",
            "muted": "#475569",
            "line": "#D8E0EB",
            "accent": "#0B7285",
            "accent_hover": "#095C6B",
        }
        self.configure(bg=self.colors["bg"])

        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("App.TFrame", background=self.colors["bg"])
        style.configure("Card.TFrame", background=self.colors["card"], relief="flat")
        style.configure(
            "Title.TLabel",
            background=self.colors["bg"],
            foreground=self.colors["text"],
            font=("Segoe UI Semibold", 18),
        )
        style.configure(
            "Subtitle.TLabel",
            background=self.colors["bg"],
            foreground=self.colors["muted"],
            font=("Segoe UI", 10),
        )
        style.configure(
            "Section.TLabelframe",
            background=self.colors["card"],
            borderwidth=1,
            relief="solid",
            bordercolor=self.colors["line"],
        )
        style.configure(
            "Section.TLabelframe.Label",
            background=self.colors["card"],
            foreground=self.colors["text"],
            font=("Segoe UI Semibold", 10),
        )
        style.configure(
            "TLabel",
            background=self.colors["card"],
            foreground=self.colors["text"],
            font=("Segoe UI", 10),
        )
        style.configure(
            "Muted.TLabel",
            background=self.colors["card"],
            foreground=self.colors["muted"],
            font=("Segoe UI", 9),
        )
        style.configure(
            "TEntry",
            fieldbackground="#FFFFFF",
            background="#FFFFFF",
            bordercolor=self.colors["line"],
            lightcolor=self.colors["line"],
            darkcolor=self.colors["line"],
            padding=7,
        )
        style.configure(
            "TButton",
            font=("Segoe UI Semibold", 10),
            padding=(12, 7),
            borderwidth=0,
        )
        style.configure(
            "Primary.TButton",
            background=self.colors["accent"],
            foreground="#FFFFFF",
        )
        style.map(
            "Primary.TButton",
            background=[("active", self.colors["accent_hover"])],
            foreground=[("active", "#FFFFFF")],
        )
        style.configure(
            "TCheckbutton",
            background=self.colors["card"],
            foreground=self.colors["text"],
            font=("Segoe UI", 10),
        )

    def _build_ui(self) -> None:
        root = ttk.Frame(self, style="App.TFrame", padding=(18, 14))
        root.pack(fill=tk.BOTH, expand=True)

        header = ttk.Frame(root, style="App.TFrame")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        ttk.Label(header, text="Fillable", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="Generate template placeholders and fill documents with Codex.",
            style="Subtitle.TLabel",
        ).pack(anchor="w")

        content = ttk.Frame(root, style="App.TFrame")
        content.grid(row=1, column=0, sticky="nsew")

        left = ttk.Frame(content, style="Card.TFrame", padding=12)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        right = ttk.Frame(content, style="Card.TFrame", padding=12)
        right.grid(row=0, column=1, sticky="nsew", padx=(8, 0))

        root.rowconfigure(1, weight=1)
        root.columnconfigure(0, weight=1)
        content.columnconfigure(0, weight=3)
        content.columnconfigure(1, weight=2)
        content.rowconfigure(0, weight=1)

        source_card = ttk.LabelFrame(left, text="Template Source", style="Section.TLabelframe", padding=10)
        source_card.grid(row=0, column=0, sticky="ew")
        ttk.Label(source_card, text="Source file (.docx/.pptx/.pdf)").grid(row=0, column=0, sticky="w")
        ttk.Entry(source_card, textvariable=self.source_var).grid(row=1, column=0, sticky="ew", pady=(4, 0))
        ttk.Button(source_card, text="Browse", command=self.browse_source).grid(row=1, column=1, padx=(8, 0))
        ttk.Checkbutton(
            source_card,
            text="Source already contains placeholders",
            variable=self.source_has_placeholders_var,
        ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(8, 0))
        source_card.columnconfigure(0, weight=1)

        run_card = ttk.LabelFrame(left, text="Run Actions", style="Section.TLabelframe", padding=10)
        run_card.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        actions = ttk.Frame(run_card, style="Card.TFrame")
        actions.grid(row=0, column=0, sticky="w")
        ttk.Button(
            actions,
            text="Generate Placeholder Template",
            command=self.on_generate_template,
            style="Primary.TButton",
        ).pack(side=tk.LEFT)
        ttk.Button(actions, text="Fill Template", command=self.on_fill).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Label(
            run_card,
            text="Tip: edit the generated template file, then fill from the same JSON.",
            style="Muted.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(8, 0))

        settings_card = ttk.LabelFrame(left, text="Settings", style="Section.TLabelframe", padding=10)
        settings_card.grid(row=2, column=0, sticky="nsew", pady=(10, 0))
        ttk.Label(settings_card, text="Codex command template").grid(row=0, column=0, sticky="w")
        ttk.Entry(settings_card, textvariable=self.codex_command_var).grid(
            row=1, column=0, columnspan=2, sticky="ew", pady=(4, 10)
        )
        ttk.Label(settings_card, text="Active template JSON").grid(row=2, column=0, sticky="w")
        ttk.Entry(settings_card, textvariable=self.template_var, state="readonly").grid(
            row=3, column=0, columnspan=2, sticky="ew", pady=(4, 10)
        )
        ttk.Label(settings_card, text="Batch data file (.csv/.json, optional)").grid(row=4, column=0, sticky="w")
        ttk.Entry(settings_card, textvariable=self.batch_data_var).grid(row=5, column=0, sticky="ew", pady=(4, 0))
        ttk.Button(settings_card, text="Browse", command=self.browse_batch_data).grid(row=5, column=1, padx=(8, 0))
        settings_card.columnconfigure(0, weight=1)
        settings_card.rowconfigure(6, weight=1)

        context_card = ttk.LabelFrame(right, text="Context Files", style="Section.TLabelframe", padding=10)
        context_card.grid(row=0, column=0, sticky="nsew")
        list_frame = ttk.Frame(context_card, style="Card.TFrame")
        list_frame.grid(row=0, column=0, sticky="nsew")
        self.context_list = tk.Listbox(
            list_frame,
            height=7,
            background="#FFFFFF",
            foreground=self.colors["text"],
            highlightthickness=1,
            highlightbackground=self.colors["line"],
            selectbackground="#D8EEF2",
        )
        self.context_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        side = ttk.Frame(list_frame, style="Card.TFrame")
        side.pack(side=tk.LEFT, padx=(8, 0), fill=tk.Y)
        ttk.Button(side, text="Add", command=self.add_context).pack(fill=tk.X)
        ttk.Button(side, text="Remove", command=self.remove_context).pack(fill=tk.X, pady=(6, 0))
        context_card.columnconfigure(0, weight=1)
        context_card.rowconfigure(0, weight=1)

        instructions_card = ttk.LabelFrame(right, text="Extra Instructions", style="Section.TLabelframe", padding=10)
        instructions_card.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
        self.instructions_text = tk.Text(
            instructions_card,
            height=8,
            wrap=tk.WORD,
            font=("Segoe UI", 10),
            background="#FFFFFF",
            foreground=self.colors["text"],
            highlightthickness=1,
            highlightbackground=self.colors["line"],
            insertbackground=self.colors["text"],
            padx=8,
            pady=8,
        )
        self.instructions_text.grid(row=0, column=0, sticky="nsew")
        instructions_card.columnconfigure(0, weight=1)
        instructions_card.rowconfigure(0, weight=1)

        logs_card = ttk.LabelFrame(right, text="Logs", style="Section.TLabelframe", padding=10)
        logs_card.grid(row=2, column=0, sticky="nsew", pady=(10, 0))
        self.log_text = tk.Text(
            logs_card,
            height=12,
            wrap=tk.WORD,
            font=("Consolas", 10),
            background="#0D1B2A",
            foreground="#DCE6F2",
            insertbackground="#DCE6F2",
            padx=8,
            pady=8,
        )
        self.log_text.grid(row=0, column=0, sticky="nsew")
        logs_card.columnconfigure(0, weight=1)
        logs_card.rowconfigure(0, weight=1)

        left.columnconfigure(0, weight=1)
        left.rowconfigure(2, weight=1)
        right.columnconfigure(0, weight=1)
        right.rowconfigure(0, weight=2)
        right.rowconfigure(1, weight=2)
        right.rowconfigure(2, weight=3)

    def log(self, message: str) -> None:
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.update_idletasks()

    def browse_source(self) -> None:
        path = filedialog.askopenfilename(
            filetypes=[("Documents", "*.docx *.pptx *.pdf"), ("All files", "*.*")]
        )
        if path:
            self.source_var.set(path)

    def browse_batch_data(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Batch Data", "*.csv *.json"), ("All files", "*.*")])
        if path:
            self.batch_data_var.set(path)

    def add_context(self) -> None:
        paths = filedialog.askopenfilenames(
            filetypes=[("Documents", "*.docx *.pptx *.pdf *.txt *.md"), ("All files", "*.*")]
        )
        for p in paths:
            self.context_list.insert(tk.END, p)

    def remove_context(self) -> None:
        selected = list(self.context_list.curselection())
        selected.reverse()
        for idx in selected:
            self.context_list.delete(idx)

    def _build_codex(self) -> CodexCli:
        cmd = self.codex_command_var.get().strip()
        if not cmd:
            raise ValueError("Codex command template is required")
        if "{prompt}" not in cmd and "{prompt_file}" not in cmd:
            raise ValueError("Codex command template must include {prompt} or {prompt_file}")
        self.config_data = AppConfig(codex_command_template=cmd)
        save_config(self.config_data)
        return CodexCli(self.config_data)

    def _resolve_template_path(self, force_regenerate: bool) -> Path:
        source_raw = self.source_var.get().strip()
        if not source_raw:
            raise ValueError("Select a source file first.")
        source_path = Path(source_raw).resolve()

        current_template_raw = self.template_var.get().strip()
        if current_template_raw and not force_regenerate:
            current_template = Path(current_template_raw).resolve()
            if current_template.exists():
                return current_template

        if self.source_has_placeholders_var.get():
            template_path = create_template_from_user_template(source_path, log=self.log)
        else:
            codex = self._build_codex()
            template_path = generate_template(source_path, codex, log=self.log)

        self.template_var.set(str(template_path))
        return template_path

    @staticmethod
    def _template_document_from_json(template_json: Path) -> Path:
        data = load_json(template_json.resolve())
        raw = str(data.get("template_file", "")).strip()
        if not raw:
            raise ValueError("Template JSON does not contain template_file.")
        template_file = Path(raw)
        if not template_file.is_absolute():
            template_file = template_json.resolve().parent / template_file
        return template_file.resolve()

    def on_generate_template(self) -> None:
        try:
            template_path = self._resolve_template_path(force_regenerate=True)
            template_doc = self._template_document_from_json(template_path)
            messagebox.showinfo("Done", f"Template ready:\n{template_path}")
            if template_doc.exists() and messagebox.askyesno(
                "Edit template", "Open the generated template document now for manual edits?"
            ):
                os.startfile(str(template_doc))
                self.log(f"Opened template document: {template_doc}")
        except (CodexCliError, Exception) as exc:
            messagebox.showerror("Error", str(exc))
            self.log(f"ERROR: {exc}")

    def on_fill(self) -> None:
        context = [Path(self.context_list.get(i)) for i in range(self.context_list.size())]
        instructions = self.instructions_text.get("1.0", tk.END).strip()

        try:
            template_json = self._resolve_template_path(force_regenerate=False)
            batch_data = self.batch_data_var.get().strip()
            codex = self._build_codex()

            if batch_data:
                outputs = fill_template_multiple(
                    template_json,
                    codex,
                    Path(batch_data),
                    context_files=context,
                    extra_instructions=instructions,
                    log=self.log,
                )
                messagebox.showinfo("Done", f"Created {len(outputs)} filled files:\n{outputs[0].parent}")
                return

            output = fill_template(
                template_json,
                codex,
                context_files=context,
                extra_instructions=instructions,
                log=self.log,
            )
            messagebox.showinfo("Done", f"Filled file created:\n{output}")
        except (CodexCliError, Exception) as exc:
            messagebox.showerror("Error", str(exc))
            self.log(f"ERROR: {exc}")
