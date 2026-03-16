from __future__ import annotations

import json
import os
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from urllib import error, request
import uuid

from app.backend_auth import BackendAuthError, login as backend_login, refresh as backend_refresh, register as backend_register
from app.codex_cli import USER_OPENAI_KEY_TARGET, CodexCli, CodexCliError
from app.firebase_auth import is_token_expired, jwt_expiry_epoch
from app.secure_store import (
    delete_firebase_refresh_token,
    get_firebase_refresh_token,
    get_user_openai_key,
    set_firebase_refresh_token,
    set_user_openai_key,
)
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
        self.openai_model_var = tk.StringVar(value=self.config_data.openai_model)
        self.backend_api_base_var = tk.StringVar(value=self.config_data.backend_api_base)
        self.firebase_email_var = tk.StringVar(value=self.config_data.firebase_email)
        self.firebase_uid_var = tk.StringVar(value=self.config_data.firebase_uid)
        self.firebase_id_token_var = tk.StringVar(value=self.config_data.firebase_id_token)
        self.firebase_password_var = tk.StringVar(value="")
        self.user_api_key_var = tk.StringVar(value=get_user_openai_key(USER_OPENAI_KEY_TARGET) or "")
        self.account_status_var = tk.StringVar(value=self._format_account_status())
        self.subscription_status_var = tk.StringVar(value="Credits: --")
        self.subscription_hint_var = tk.StringVar(value="")
        self.onboarding_api_key_var = tk.StringVar(value="")
        self.user_api_key_entry = None
        self.firebase_email_entry = None
        self.firebase_password_entry = None
        self.account_action_frame = None
        self.auth_window: tk.Toplevel | None = None
        self.auth_button = None
        self.upgrade_button = None
        self.credit_poll_after_upgrade = 0
        self.oauth_session_id: str | None = None
        self.oauth_poll_attempts = 0

        self._init_styles()
        self._set_window_icon()
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
        self._refresh_subscription_ui()
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

    @staticmethod
    def _resource_path(filename: str) -> Path:
        base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[1]))
        return base / filename

    def _set_window_icon(self) -> None:
        ico_path = self._resource_path("fillableicon.ico")
        if ico_path.exists() and sys.platform.startswith("win"):
            try:
                self.iconbitmap(str(ico_path))
                return
            except Exception:
                pass
        return

    def _build_onboarding_ui(self) -> None:
        root = ttk.Frame(self.onboarding_page, style="App.TFrame", padding=(24, 20))
        root.pack(fill=tk.BOTH, expand=True)

        card = ttk.Frame(root, style="Card.TFrame", padding=20)
        card.place(relx=0.5, rely=0.45, anchor="center")

        ttk.Label(card, text=f"Welcome to {APP_NAME}", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            card,
            text="Enter your OpenAI API key to get started. You can change it later on the main screen.",
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
        header_actions = ttk.Frame(header, style="App.TFrame")
        header_actions.pack(anchor="e", pady=(6, 0))
        self.auth_button = ttk.Button(
            header_actions,
            text=self._auth_button_label(),
            command=self._on_auth_button,
        )
        self.auth_button.pack(side=tk.LEFT, padx=(0, 8))
        self.upgrade_button = ttk.Button(header_actions, text="Buy credits", command=self._upgrade)
        self.upgrade_button.pack(side=tk.LEFT, padx=(0, 8))
        # Settings removed per requirements; model selection is on the main page.
        ttk.Label(header, textvariable=self.subscription_status_var, style="Muted.TLabel").pack(anchor="e", pady=(4, 0))

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

        ai_card = ttk.LabelFrame(left, text="AI Settings", style="Section.TLabelframe", padding=10)
        ai_card.grid(row=0, column=0, sticky="ew")
        ttk.Label(ai_card, text="Model").grid(row=0, column=0, sticky="w")
        model_combo = ttk.Combobox(
            ai_card,
            textvariable=self.openai_model_var,
            state="readonly",
            values=["gpt-5.4", "gpt-5.3-codex", "gpt-5.2", "gpt-4o"],
        )
        model_combo.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(4, 8))
        self.free_key_label = ttk.Label(ai_card, text="OpenAI API key (free plan)")
        self.free_key_label.grid(row=2, column=0, sticky="w")
        self.user_api_key_entry = ttk.Entry(ai_card, textvariable=self.user_api_key_var, show="*")
        self.user_api_key_entry.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(4, 8))
        self.save_key_button = ttk.Button(ai_card, text="Save key", command=self._save_openai_key)
        self.save_key_button.grid(row=4, column=0, sticky="w")
        ttk.Label(ai_card, textvariable=self.subscription_hint_var, style="Muted.TLabel").grid(
            row=5, column=0, columnspan=2, sticky="w", pady=(6, 0)
        )
        ai_card.columnconfigure(0, weight=1)

        source_card = ttk.LabelFrame(left, text="Template Source", style="Section.TLabelframe", padding=10)
        source_card.grid(row=1, column=0, sticky="ew")
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
        run_card.grid(row=2, column=0, sticky="ew", pady=(10, 0))
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
        template_card.grid(row=3, column=0, sticky="nsew", pady=(10, 0))
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
        left.rowconfigure(3, weight=1)
        right.columnconfigure(0, weight=1)
        right.rowconfigure(0, weight=2)
        right.rowconfigure(1, weight=2)
        right.rowconfigure(2, weight=3)

    def _show_initial_page(self) -> None:
        has_user_key = bool((get_user_openai_key(USER_OPENAI_KEY_TARGET) or "").strip())
        if has_user_key:
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
        self._show_app_page()

    def _skip_onboarding(self) -> None:
        self._show_app_page()

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
        model = self.openai_model_var.get().strip()
        if model not in {"gpt-5.4", "gpt-5.3-codex", "gpt-5.2", "gpt-4o"}:
            raise ValueError("Model must be one of: gpt-5.4, gpt-5.3-codex, gpt-5.2, gpt-4o.")
        if not model:
            raise ValueError("OpenAI model is required")
        backend_api_base = self._backend_base()
        firebase_id_token = self.firebase_id_token_var.get().strip()
        use_subscription = bool(self.config_data.credit_balance > 0 and firebase_id_token)

        user_key = self.user_api_key_var.get().strip()
        if use_subscription:
            if not backend_api_base:
                raise ValueError("Backend URL is required for credit mode.")
            if not firebase_id_token:
                raise ValueError("Firebase login or ID token is required for credit mode.")
        else:
            if not user_key:
                raise ValueError("OpenAI API key is required for free mode.")
            set_user_openai_key(USER_OPENAI_KEY_TARGET, user_key)

        self.config_data = AppConfig(
            credit_balance=self.config_data.credit_balance,
            openai_model=model,
            backend_api_base=backend_api_base,
            firebase_id_token=firebase_id_token,
            firebase_token_expiry_utc=self.config_data.firebase_token_expiry_utc,
            firebase_email=self.firebase_email_var.get().strip(),
            firebase_uid=self.firebase_uid_var.get().strip(),
            codex_command_template=self.config_data.codex_command_template,
        )
        save_config(self.config_data)
        return CodexCli(self.config_data)

    def _refresh_subscription_ui(self) -> None:
        if not self.user_api_key_entry:
            return
        if not self.user_api_key_entry.winfo_exists():
            return
        if self.config_data.credit_balance > 0:
            self.user_api_key_entry.state(["disabled"])
            self.free_key_label.grid_remove()
            self.user_api_key_entry.grid_remove()
            self.save_key_button.grid_remove()
            self.subscription_hint_var.set(f"Credits available: {self.config_data.credit_balance:.2f}")
        else:
            self.user_api_key_entry.state(["!disabled"])
            self.free_key_label.grid()
            self.user_api_key_entry.grid()
            self.save_key_button.grid()
            self.subscription_hint_var.set("")
        self._sync_auth_button()

    def _auth_button_label(self) -> str:
        return "Sign out" if self.config_data.firebase_uid else "Sign in"

    def _sync_auth_button(self) -> None:
        if self.auth_button is not None:
            self.auth_button.config(text=self._auth_button_label())
        signed_in = bool(self.config_data.firebase_uid)
        if self.upgrade_button is not None:
            self.upgrade_button.state(["!disabled"] if signed_in else ["disabled"])

    def _on_auth_button(self) -> None:
        if self.config_data.firebase_uid:
            self._sign_out()
            return
        self._open_auth_window()

    def _open_auth_window(self) -> None:
        if self.auth_window is not None and self.auth_window.winfo_exists():
            self.auth_window.lift()
            self.auth_window.focus_force()
            return

        win = tk.Toplevel(self)
        win.title("Account")
        win.geometry("560x420")
        win.minsize(520, 380)
        win.configure(bg=self.colors["bg"])
        self.auth_window = win

        container = ttk.Frame(win, style="App.TFrame", padding=14)
        container.pack(fill=tk.BOTH, expand=True)

        auth_card = ttk.LabelFrame(container, text="Sign In / Register", style="Section.TLabelframe", padding=10)
        auth_card.pack(fill=tk.X, expand=False)
        ttk.Label(auth_card, text="Email").grid(row=0, column=0, sticky="w")
        self.firebase_email_entry = ttk.Entry(auth_card, textvariable=self.firebase_email_var)
        self.firebase_email_entry.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(4, 8))
        ttk.Label(auth_card, text="Password").grid(row=2, column=0, sticky="w")
        self.firebase_password_entry = ttk.Entry(auth_card, textvariable=self.firebase_password_var, show="*")
        self.firebase_password_entry.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(4, 8))
        actions = ttk.Frame(auth_card, style="Card.TFrame")
        actions.grid(row=4, column=0, columnspan=2, sticky="w")
        ttk.Button(actions, text="Register", style="Primary.TButton", command=self._register_account).pack(side=tk.LEFT)
        ttk.Button(actions, text="Sign in", command=self._sign_in).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(actions, text="Sign out", command=self._sign_out).pack(side=tk.LEFT, padx=(8, 0))
        oauth_actions = ttk.Frame(auth_card, style="Card.TFrame")
        oauth_actions.grid(row=5, column=0, columnspan=2, sticky="w", pady=(8, 0))
        ttk.Button(oauth_actions, text="Sign in with Google", command=lambda: self._oauth_sign_in("google")).pack(
            side=tk.LEFT
        )
        auth_card.columnconfigure(0, weight=1)

        status_card = ttk.LabelFrame(container, text="Account Status", style="Section.TLabelframe", padding=10)
        status_card.pack(fill=tk.BOTH, expand=True, pady=(12, 0))
        ttk.Label(status_card, textvariable=self.account_status_var, style="Muted.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        billing_actions = ttk.Frame(status_card, style="Card.TFrame")
        billing_actions.grid(row=1, column=0, sticky="w", pady=(10, 0))
        ttk.Button(billing_actions, text="Check credits", command=self._check_subscription).pack(side=tk.LEFT)
        ttk.Button(billing_actions, text="Buy credits", style="Primary.TButton", command=self._upgrade).pack(
            side=tk.LEFT, padx=(8, 0)
        )
        status_card.columnconfigure(0, weight=1)

    def _format_account_status(self) -> str:
        if self.config_data.firebase_uid:
            email = self.config_data.firebase_email or "unknown"
            return f"Signed in as {email} ({self.config_data.firebase_uid})."
        return "Not signed in."

    def _set_account_status(self, message: str) -> None:
        self.account_status_var.set(message)

    def _save_openai_key(self) -> None:
        api_key = self.user_api_key_var.get().strip()
        if not api_key:
            messagebox.showerror("OpenAI key", "Enter an OpenAI API key first.")
            return
        set_user_openai_key(USER_OPENAI_KEY_TARGET, api_key)
        messagebox.showinfo("OpenAI key", "Key saved.")

    def _oauth_sign_in(self, provider: str) -> None:
        provider = provider.strip().lower()
        if provider != "google":
            messagebox.showerror("Sign in", "Unsupported provider.")
            return
        if self.oauth_session_id:
            messagebox.showinfo("Sign in", "Sign-in is already in progress.")
            return
        base = self._backend_base()
        if not base:
            messagebox.showerror("Sign in", "Backend URL is required.")
            return
        self.oauth_session_id = str(uuid.uuid4())
        self.oauth_poll_attempts = 0
        os.startfile(f"{base.rstrip('/')}/auth/google?session_id={self.oauth_session_id}")
        self.after(2500, self._poll_oauth_session)

    def _poll_oauth_session(self) -> None:
        if not self.oauth_session_id:
            return
        self.oauth_poll_attempts += 1
        if self.oauth_poll_attempts > 20:
            self.oauth_session_id = None
            messagebox.showerror("Sign in", "Sign-in timed out. Please try again.")
            return
        base = self._backend_base()
        if not base:
            self.oauth_session_id = None
            return
        url = f"{base.rstrip('/')}/v1/auth/poll?session_id={self.oauth_session_id}"
        try:
            req = request.Request(url, method="GET")
            with request.urlopen(req, timeout=15) as resp:
                text = resp.read().decode("utf-8", errors="ignore")
            data = json.loads(text)
        except Exception:
            self.after(2500, self._poll_oauth_session)
            return
        if not isinstance(data, dict) or data.get("status") != "ok":
            self.after(2500, self._poll_oauth_session)
            return
        self.oauth_session_id = None
        self._handle_oauth_payload(data)

    def _handle_oauth_payload(self, payload: dict) -> None:
        id_token = str(payload.get("id_token", "")).strip()
        refresh_token = str(payload.get("refresh_token", "")).strip()
        email_out = str(payload.get("email", "")).strip()
        uid = str(payload.get("uid", "")).strip()
        if not id_token or not uid:
            messagebox.showerror("Sign in", "OAuth login did not return a valid token.")
            return
        exp_epoch = jwt_expiry_epoch(id_token) or 0
        self.config_data.firebase_id_token = id_token
        self.config_data.firebase_token_expiry_utc = exp_epoch
        self.config_data.firebase_email = email_out
        self.config_data.firebase_uid = uid
        if refresh_token:
            set_firebase_refresh_token("Fillable.Firebase.RefreshToken", refresh_token)
        self.firebase_id_token_var.set(id_token)
        self.firebase_email_var.set(email_out)
        self.firebase_uid_var.set(uid)
        save_config(self.config_data)
        self._set_account_status(self._format_account_status())
        self._sync_auth_button()
        self._check_subscription(silent=True)

    def _save_auth_state(
        self,
        *,
        id_token: str,
        refresh_token: str,
        expires_in: int,
        email: str,
        uid: str,
    ) -> None:
        exp_epoch = int(__import__("time").time()) + int(expires_in or 0)
        if not exp_epoch:
            exp_epoch = jwt_expiry_epoch(id_token) or 0
        self.config_data.firebase_id_token = id_token
        self.config_data.firebase_token_expiry_utc = exp_epoch
        self.config_data.firebase_email = email
        self.config_data.firebase_uid = uid
        if refresh_token:
            set_firebase_refresh_token("Fillable.Firebase.RefreshToken", refresh_token)
        self.firebase_id_token_var.set(id_token)
        self.firebase_email_var.set(email)
        self.firebase_uid_var.set(uid)
        save_config(self.config_data)
        self._set_account_status(self._format_account_status())
        self._sync_auth_button()

    def _current_token(self) -> str:
        token = (self.firebase_id_token_var.get().strip() or self.config_data.firebase_id_token).strip()
        exp_epoch = int(self.config_data.firebase_token_expiry_utc or 0)
        if token and not is_token_expired(exp_epoch):
            return token
        refresh = get_firebase_refresh_token("Fillable.Firebase.RefreshToken") or ""
        base = self._backend_base()
        if not refresh or not base:
            return token
        try:
            refreshed = backend_refresh(base, refresh)
        except Exception:
            return token
        token = str(refreshed.get("id_token", "") or refreshed.get("idToken", "")).strip()
        refresh = str(refreshed.get("refresh_token", "") or refreshed.get("refreshToken", "")).strip()
        expires_in = int(refreshed.get("expires_in", 0) or refreshed.get("expiresIn", 0) or 0)
        if token:
            exp_epoch = int(__import__("time").time()) + expires_in if expires_in else (jwt_expiry_epoch(token) or 0)
            self.config_data.firebase_id_token = token
            if refresh:
                set_firebase_refresh_token("Fillable.Firebase.RefreshToken", refresh)
            self.config_data.firebase_token_expiry_utc = exp_epoch
            self.firebase_id_token_var.set(token)
            save_config(self.config_data)
        return token

    def _backend_base(self) -> str:
        base = (self.backend_api_base_var.get().strip() or self.config_data.backend_api_base).strip()
        if (
            not base
            or "localhost" in base
            or "127.0.0.1" in base
            or "[::1]" in base
        ):
            base = AppConfig().backend_api_base
        if self.backend_api_base_var.get().strip() != base:
            self.backend_api_base_var.set(base)
        if self.config_data.backend_api_base != base:
            self.config_data.backend_api_base = base
            save_config(self.config_data)
        return base

    def _backend_json(self, path: str, *, method: str = "GET", body: dict | None = None) -> dict:
        base = self._backend_base()
        if not base:
            raise ValueError("Backend URL is required.")
        url = base.rstrip("/") + path
        headers = {}
        data = None
        if body is not None:
            data = json.dumps(body).encode("utf-8")
            headers["Content-Type"] = "application/json"
        token = self._current_token()
        if not token:
            raise ValueError("Firebase ID token is required. Sign in first.")
        headers["Authorization"] = f"Bearer {token}"
        req = request.Request(url, data=data, headers=headers, method=method.upper())
        try:
            with request.urlopen(req, timeout=60) as resp:
                text = resp.read().decode("utf-8", errors="ignore")
        except error.HTTPError as exc:
            details = exc.read().decode("utf-8", errors="ignore")
            raise ValueError(f"HTTP {exc.code} from {url}: {details[:800]}") from exc
        except error.URLError as exc:
            raise ValueError(f"Network error calling {url}: {exc}") from exc
        data = json.loads(text)
        if not isinstance(data, dict):
            raise ValueError("Backend response was not a JSON object.")
        return data

    def _sign_in(self) -> None:
        email = self.firebase_email_var.get().strip()
        password = self.firebase_password_var.get().strip()
        if not email or not password:
            messagebox.showerror("Sign in", "Email and password are required.")
            return
        try:
            base = self._backend_base()
            if not base:
                raise ValueError("Backend URL is required.")
            result = backend_login(base, email, password)
            id_token = str(result.get("id_token", "") or result.get("idToken", "")).strip()
            refresh_token = str(result.get("refresh_token", "") or result.get("refreshToken", "")).strip()
            expires_in = int(result.get("expires_in", 0) or result.get("expiresIn", 0) or 0)
            uid = str(result.get("uid", "") or result.get("localId", "")).strip()
            email_out = str(result.get("email", email)).strip()
            if not id_token or not uid:
                raise BackendAuthError("Sign-in response missing token or uid.")
            self._save_auth_state(
                id_token=id_token,
                refresh_token=refresh_token,
                expires_in=expires_in,
                email=email_out,
                uid=uid,
            )
            messagebox.showinfo("Sign in", "Signed in successfully.")
            self._check_subscription(silent=True)
        except Exception as exc:
            messagebox.showerror("Sign in", str(exc))

    def _register_account(self) -> None:
        email = self.firebase_email_var.get().strip()
        password = self.firebase_password_var.get().strip()
        if not email or not password:
            messagebox.showerror("Register", "Email and password are required.")
            return
        if not self._password_ok(password):
            return
        try:
            base = self._backend_base()
            if not base:
                raise ValueError("Backend URL is required.")
            result = backend_register(base, email, password)
            id_token = str(result.get("id_token", "") or result.get("idToken", "")).strip()
            refresh_token = str(result.get("refresh_token", "") or result.get("refreshToken", "")).strip()
            expires_in = int(result.get("expires_in", 0) or result.get("expiresIn", 0) or 0)
            uid = str(result.get("uid", "") or result.get("localId", "")).strip()
            email_out = str(result.get("email", email)).strip()
            if not id_token or not uid:
                raise BackendAuthError("Register response missing token or uid.")
            self._save_auth_state(
                id_token=id_token,
                refresh_token=refresh_token,
                expires_in=expires_in,
                email=email_out,
                uid=uid,
            )
            messagebox.showinfo("Register", "Account created and signed in.")
            self._check_subscription(silent=True)
        except Exception as exc:
            messagebox.showerror("Register", str(exc))

    def _password_ok(self, password: str) -> bool:
        if len(password) < 8:
            messagebox.showerror("Register", "Password must be at least 8 characters.")
            return False
        if password.lower() == password or password.upper() == password:
            messagebox.showerror("Register", "Password must include upper and lower case letters.")
            return False
        if not any(ch.isdigit() for ch in password):
            messagebox.showerror("Register", "Password must include a number.")
            return False
        if not any(not ch.isalnum() for ch in password):
            messagebox.showerror("Register", "Password must include a special character.")
            return False
        return True

    def _sign_out(self) -> None:
        self.config_data.firebase_id_token = ""
        self.config_data.firebase_token_expiry_utc = 0
        self.config_data.firebase_email = ""
        self.config_data.firebase_uid = ""
        self.config_data.credit_balance = 0.0
        delete_firebase_refresh_token("Fillable.Firebase.RefreshToken")
        self.firebase_id_token_var.set("")
        self.firebase_email_var.set("")
        self.firebase_uid_var.set("")
        save_config(self.config_data)
        self._set_account_status(self._format_account_status())
        self._sync_auth_button()
        self.subscription_status_var.set("Credits: --")
        self._refresh_subscription_ui()
        messagebox.showinfo("Sign out", "Signed out.")

    def _check_subscription(self, *, silent: bool = False) -> None:
        try:
            data = self._backend_json("/v1/credits", method="GET")
            credits = float(data.get("credits", 0.0) or 0.0)
            self.subscription_status_var.set(f"Credits: {credits:.2f}")
            self.config_data.credit_balance = credits
            save_config(self.config_data)
            self._refresh_subscription_ui()
            status_message = f"Credits remaining: {credits:.2f}"
            self._set_account_status(status_message)
            if not silent:
                messagebox.showinfo("Credits", status_message)
        except Exception as exc:
            if not silent:
                messagebox.showerror("Credits", str(exc))

    def _upgrade(self) -> None:
        try:
            data = self._backend_json("/v1/billing/create-checkout-session", method="POST", body={})
            url = str(data.get("url", "")).strip()
            if not url:
                raise ValueError("Checkout URL not returned.")
            os.startfile(url)
            self.credit_poll_after_upgrade = 5
            self.after(15000, self._poll_credits_after_upgrade)
        except Exception as exc:
            messagebox.showerror("Upgrade", str(exc))

    def _poll_credits_after_upgrade(self) -> None:
        if self.credit_poll_after_upgrade <= 0:
            return
        self.credit_poll_after_upgrade -= 1
        try:
            self._check_subscription(silent=True)
        except Exception:
            pass
        if self.credit_poll_after_upgrade > 0:
            self.after(20000, self._poll_credits_after_upgrade)


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
