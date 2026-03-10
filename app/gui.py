from __future__ import annotations

import os
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from app.codex_cli import (
    SUBSCRIPTION_TOKEN_TARGET,
    USER_OPENAI_KEY_TARGET,
    CodexCli,
    CodexCliError,
)
from app.secure_store import get_secret, get_user_openai_key, set_secret, set_user_openai_key
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
        self.title(f"{APP_NAME} {APP_VERSION} - FillableDOC Document Autofill")
        self.geometry("1080x760")
        self.minsize(960, 700)

        self.config_data = load_config()
        self.source_var = tk.StringVar(value=initial_generate or "")
        self.template_var = tk.StringVar(value=initial_template or "")
        self.source_has_placeholders_var = tk.BooleanVar(value=False)
        self.batch_data_var = tk.StringVar(value="")
        self.ai_mode_var = tk.StringVar(value=self.config_data.ai_mode)
        self.openai_model_var = tk.StringVar(value=self.config_data.openai_model)
        self.openai_api_base_var = tk.StringVar(value=self.config_data.openai_api_base)
        self.subscription_api_base_var = tk.StringVar(value=self.config_data.subscription_api_base)
        self.user_api_key_var = tk.StringVar(value=get_user_openai_key(USER_OPENAI_KEY_TARGET) or "")
        self.subscription_token_var = tk.StringVar(value=get_secret(SUBSCRIPTION_TOKEN_TARGET) or "")
        self.onboarding_api_key_var = tk.StringVar(value="")
        self.settings_window: tk.Toplevel | None = None
        self.user_api_key_entry = None
        self.subscription_api_base_entry = None
        self.subscription_token_entry = None

        self._init_styles()
        self.bg_canvas = tk.Canvas(self, highlightthickness=0, bd=0)
        self.bg_canvas.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.page_container = ttk.Frame(self, style="App.TFrame")
        self.page_container.place(relx=0.02, rely=0.02, relwidth=0.96, relheight=0.96)
        self.page_container.lift()
        self.bind("<Configure>", self._on_resize)
        self.onboarding_page = ttk.Frame(self.page_container, style="App.TFrame")
        self.app_page = ttk.Frame(self.page_container, style="App.TFrame")
        self._build_onboarding_ui()
        self._build_ui(self.app_page)
        self._refresh_mode_ui()
        self._show_initial_page()

    def _init_styles(self) -> None:
        self.colors = {
            "bg": "#F7F2EC",
            "bg_start": "#F8EFE6",
            "bg_end": "#E6EFF7",
            "card": "#FFFFFF",
            "text": "#1C1B1F",
            "muted": "#5E6B78",
            "line": "#D9D0C7",
            "accent": "#C65F3D",
            "accent_hover": "#B15435",
            "glow_1": "#FADFD1",
            "glow_2": "#D6EEF3",
            "glow_3": "#F5E7D3",
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
            font=("Bahnschrift SemiBold", 20),
        )
        style.configure(
            "Subtitle.TLabel",
            background=self.colors["bg"],
            foreground=self.colors["muted"],
            font=("Bahnschrift", 11),
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
            font=("Bahnschrift SemiBold", 10),
        )
        style.configure(
            "TLabel",
            background=self.colors["card"],
            foreground=self.colors["text"],
            font=("Bahnschrift", 10),
        )
        style.configure(
            "Muted.TLabel",
            background=self.colors["card"],
            foreground=self.colors["muted"],
            font=("Bahnschrift", 9),
        )
        style.configure(
            "TEntry",
            fieldbackground="#FFFFFF",
            background="#FFFFFF",
            bordercolor=self.colors["line"],
            lightcolor=self.colors["line"],
            darkcolor=self.colors["line"],
            padding=8,
        )
        style.configure(
            "TButton",
            font=("Bahnschrift SemiBold", 10),
            padding=(12, 8),
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
            font=("Bahnschrift", 10),
        )

    def _on_resize(self, event: tk.Event) -> None:
        if event.widget is self:
            self._draw_background(event.width, event.height)

    def _draw_background(self, width: int, height: int) -> None:
        self.bg_canvas.delete("bg")
        if width < 2 or height < 2:
            return
        start = self.colors["bg_start"].lstrip("#")
        end = self.colors["bg_end"].lstrip("#")
        r1, g1, b1 = int(start[0:2], 16), int(start[2:4], 16), int(start[4:6], 16)
        r2, g2, b2 = int(end[0:2], 16), int(end[2:4], 16), int(end[4:6], 16)
        for i in range(height):
            ratio = i / max(height - 1, 1)
            r = int(r1 + (r2 - r1) * ratio)
            g = int(g1 + (g2 - g1) * ratio)
            b = int(b1 + (b2 - b1) * ratio)
            color = f"#{r:02x}{g:02x}{b:02x}"
            self.bg_canvas.create_line(0, i, width, i, fill=color, tags="bg")
        self.bg_canvas.create_oval(-120, -120, 260, 260, fill=self.colors["glow_1"], outline="", tags="bg")
        self.bg_canvas.create_oval(width - 320, 80, width + 80, 480, fill=self.colors["glow_2"], outline="", tags="bg")
        self.bg_canvas.create_oval(
            int(width * 0.35),
            height - 200,
            int(width * 0.85),
            height + 200,
            fill=self.colors["glow_3"],
            outline="",
            tags="bg",
        )

    def _build_onboarding_ui(self) -> None:
        root = ttk.Frame(self.onboarding_page, style="App.TFrame", padding=(24, 20))
        root.pack(fill=tk.BOTH, expand=True)

        card = ttk.Frame(root, style="Card.TFrame", padding=20)
        card.place(relx=0.5, rely=0.45, anchor="center")

        ttk.Label(card, text=f"Welcome to {APP_NAME}", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            card,
            text="Enter your OpenAI API key to get started. You can change this later in Settings.",
            style="Subtitle.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(4, 12))
        ttk.Label(card, text="OpenAI API key").grid(row=2, column=0, sticky="w")
        ttk.Entry(card, textvariable=self.onboarding_api_key_var, width=62, show="*").grid(
            row=3, column=0, sticky="ew", pady=(4, 10)
        )
        actions = ttk.Frame(card, style="Card.TFrame")
        actions.grid(row=4, column=0, sticky="e")
        ttk.Button(actions, text="Continue", style="Primary.TButton", command=self._complete_onboarding).pack(
            side=tk.LEFT
        )
        ttk.Button(actions, text="Skip for now", command=self._skip_onboarding).pack(side=tk.LEFT, padx=(8, 0))
        card.columnconfigure(0, weight=1)

    def _build_ui(self, parent: ttk.Frame) -> None:
        root = ttk.Frame(parent, style="App.TFrame", padding=(18, 14))
        root.pack(fill=tk.BOTH, expand=True)

        header = ttk.Frame(root, style="App.TFrame")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        ttk.Label(header, text=APP_NAME, style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="Generate template placeholders and fill documents with FillableDOC.",
            style="Subtitle.TLabel",
        ).pack(anchor="w")
        ttk.Button(header, text="Settings", command=self.open_settings_window).pack(anchor="e", pady=(6, 0))

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

        template_card = ttk.LabelFrame(left, text="Template & Batch", style="Section.TLabelframe", padding=10)
        template_card.grid(row=2, column=0, sticky="nsew", pady=(10, 0))
        ttk.Label(template_card, text="Active template JSON").grid(row=0, column=0, sticky="w")
        ttk.Entry(template_card, textvariable=self.template_var, state="readonly").grid(
            row=1, column=0, columnspan=2, sticky="ew", pady=(4, 10)
        )
        ttk.Label(template_card, text="Batch data file (.csv/.json, optional)").grid(row=2, column=0, sticky="w")
        ttk.Entry(template_card, textvariable=self.batch_data_var).grid(row=3, column=0, sticky="ew", pady=(4, 0))
        ttk.Button(template_card, text="Browse", command=self.browse_batch_data).grid(row=3, column=1, padx=(8, 0))
        template_card.columnconfigure(0, weight=1)
        template_card.rowconfigure(4, weight=1)

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

    def _show_initial_page(self) -> None:
        has_user_key = bool((get_user_openai_key(USER_OPENAI_KEY_TARGET) or "").strip())
        has_subscription_token = bool((get_secret(SUBSCRIPTION_TOKEN_TARGET) or "").strip())
        mode = self.ai_mode_var.get().strip().lower()
        if has_user_key or (mode == "app_subscription" and has_subscription_token):
            self._show_app_page()
            return
        self._show_onboarding_page()

    def _show_onboarding_page(self) -> None:
        self.app_page.pack_forget()
        self.onboarding_page.pack(fill=tk.BOTH, expand=True)

    def _show_app_page(self) -> None:
        self.onboarding_page.pack_forget()
        self.app_page.pack(fill=tk.BOTH, expand=True)

    def _complete_onboarding(self) -> None:
        api_key = self.onboarding_api_key_var.get().strip()
        if not api_key:
            messagebox.showerror("Missing key", "Enter an OpenAI API key or click 'Skip for now'.")
            return
        set_user_openai_key(USER_OPENAI_KEY_TARGET, api_key)
        self.user_api_key_var.set(api_key)
        self.ai_mode_var.set("user_key")
        self._refresh_mode_ui()
        self._show_app_page()

    def _skip_onboarding(self) -> None:
        self._show_app_page()

    def open_settings_window(self) -> None:
        if self.settings_window is not None and self.settings_window.winfo_exists():
            self.settings_window.lift()
            self.settings_window.focus_force()
            return

        win = tk.Toplevel(self)
        win.title("Settings")
        win.geometry("720x430")
        win.minsize(640, 390)
        win.configure(bg=self.colors["bg"])
        self.settings_window = win

        container = ttk.Frame(win, style="App.TFrame", padding=14)
        container.pack(fill=tk.BOTH, expand=True)
        card = ttk.LabelFrame(container, text="AI Settings", style="Section.TLabelframe", padding=10)
        card.pack(fill=tk.BOTH, expand=True)

        ttk.Label(card, text="AI mode").grid(row=0, column=0, sticky="w")
        mode_combo = ttk.Combobox(
            card,
            textvariable=self.ai_mode_var,
            state="readonly",
            values=["user_key", "app_subscription"],
        )
        mode_combo.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(4, 8))
        mode_combo.bind("<<ComboboxSelected>>", lambda _: self._refresh_mode_ui())

        ttk.Label(card, text="OpenAI model").grid(row=2, column=0, sticky="w")
        ttk.Entry(card, textvariable=self.openai_model_var).grid(row=3, column=0, columnspan=2, sticky="ew", pady=(4, 8))

        ttk.Label(card, text="OpenAI API base URL").grid(row=4, column=0, sticky="w")
        ttk.Entry(card, textvariable=self.openai_api_base_var).grid(
            row=5, column=0, columnspan=2, sticky="ew", pady=(4, 8)
        )

        ttk.Label(card, text="Your OpenAI API key (secured with Windows DPAPI)").grid(
            row=6, column=0, sticky="w"
        )
        self.user_api_key_entry = ttk.Entry(card, textvariable=self.user_api_key_var, show="*")
        self.user_api_key_entry.grid(row=7, column=0, columnspan=2, sticky="ew", pady=(4, 8))

        ttk.Label(card, text="Subscription backend URL").grid(row=8, column=0, sticky="w")
        self.subscription_api_base_entry = ttk.Entry(card, textvariable=self.subscription_api_base_var)
        self.subscription_api_base_entry.grid(row=9, column=0, columnspan=2, sticky="ew", pady=(4, 8))

        ttk.Label(card, text="Subscription access token (stored in Credential Manager)").grid(
            row=10, column=0, sticky="w"
        )
        self.subscription_token_entry = ttk.Entry(card, textvariable=self.subscription_token_var, show="*")
        self.subscription_token_entry.grid(row=11, column=0, columnspan=2, sticky="ew", pady=(4, 10))

        actions = ttk.Frame(card, style="Card.TFrame")
        actions.grid(row=12, column=0, columnspan=2, sticky="e")
        ttk.Button(actions, text="Save", style="Primary.TButton", command=self._save_settings).pack(side=tk.LEFT)
        ttk.Button(actions, text="Close", command=win.destroy).pack(side=tk.LEFT, padx=(8, 0))
        card.columnconfigure(0, weight=1)
        self._refresh_mode_ui()

    def _save_settings(self) -> None:
        try:
            self._build_codex()
            messagebox.showinfo("Settings", "Settings saved.")
        except Exception as exc:
            messagebox.showerror("Settings", str(exc))

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

    def _refresh_mode_ui(self) -> None:
        mode = self.ai_mode_var.get().strip().lower()
        user_mode = mode != "app_subscription"
        if not self.user_api_key_entry or not self.subscription_api_base_entry or not self.subscription_token_entry:
            return
        if not self.user_api_key_entry.winfo_exists():
            return
        if user_mode:
            self.user_api_key_entry.state(["!disabled"])
            self.subscription_api_base_entry.state(["disabled"])
            self.subscription_token_entry.state(["disabled"])
        else:
            self.user_api_key_entry.state(["disabled"])
            self.subscription_api_base_entry.state(["!disabled"])
            self.subscription_token_entry.state(["!disabled"])

    def _build_codex(self) -> CodexCli:
        mode = self.ai_mode_var.get().strip().lower()
        if mode not in {"user_key", "app_subscription"}:
            raise ValueError("AI mode must be user_key or app_subscription")
        model = self.openai_model_var.get().strip()
        if not model:
            raise ValueError("OpenAI model is required")
        openai_api_base = self.openai_api_base_var.get().strip() or "https://api.openai.com/v1"
        subscription_api_base = self.subscription_api_base_var.get().strip()

        user_key = self.user_api_key_var.get().strip()
        if user_key:
            set_user_openai_key(USER_OPENAI_KEY_TARGET, user_key)
        token = self.subscription_token_var.get().strip()
        if token:
            set_secret(SUBSCRIPTION_TOKEN_TARGET, token, username="fillable-subscription")

        self.config_data = AppConfig(
            ai_mode=mode,
            openai_model=model,
            openai_api_base=openai_api_base,
            subscription_api_base=subscription_api_base,
            codex_command_template=self.config_data.codex_command_template,
        )
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
            codex = None
            try:
                codex = self._build_codex()
            except Exception as exc:
                self.log(f"AI unavailable for template import; using local mapping: {exc}")
            template_path = create_template_from_user_template(source_path, codex=codex, log=self.log)
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
