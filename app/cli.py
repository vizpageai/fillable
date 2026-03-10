from __future__ import annotations

import argparse
from pathlib import Path

from app.codex_cli import SUBSCRIPTION_TOKEN_TARGET, USER_OPENAI_KEY_TARGET, CodexCli
from app.context_menu import install_context_menu, uninstall_context_menu
from app.template_engine import (
    create_template_from_user_template,
    fill_template,
    fill_template_multiple,
    generate_template,
)
from app.utils import load_config
from app.models import AppConfig
from app.secure_store import delete_secret, delete_user_openai_key, set_secret, set_user_openai_key


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="fillable")
    parser.add_argument("--generate-template", type=Path)
    parser.add_argument("--import-template-file", type=Path)
    parser.add_argument("--fill-template", type=Path)
    parser.add_argument("--batch-data", type=Path)
    parser.add_argument("--batch-output-dir", type=Path)
    parser.add_argument(
        "--context",
        type=str,
        default="",
        help="Semicolon-separated context file paths",
    )
    parser.add_argument("--instructions", type=str, default="")
    parser.add_argument(
        "--prompt-instructions",
        action="store_true",
        help="Prompt in terminal for fill instructions before calling AI.",
    )
    parser.add_argument("--print-config", action="store_true")
    parser.add_argument("--install-context-menu", action="store_true")
    parser.add_argument("--uninstall-context-menu", action="store_true")
    parser.add_argument("--context-menu-exe", type=str, default="")
    parser.add_argument("--no-prompt-in-context-menu", action="store_true")
    parser.add_argument("--set-ai-mode", choices=["user_key", "app_subscription"])
    parser.add_argument("--set-openai-model", type=str)
    parser.add_argument("--set-openai-api-base", type=str)
    parser.add_argument("--set-subscription-api-base", type=str)
    parser.add_argument("--set-user-openai-key", type=str)
    parser.add_argument("--clear-user-openai-key", action="store_true")
    parser.add_argument("--set-subscription-token", type=str)
    parser.add_argument("--clear-subscription-token", action="store_true")
    return parser


def _resolve_fill_template_target(path: Path) -> Path:
    candidate = path.resolve()
    if candidate.name.lower().endswith(".fillable.json"):
        return candidate

    # Allow right-click fill directly from template files such as:
    # report.template.docx, slides.template.pptx, form.template.pdf, notes.template.txt
    stem_lower = candidate.stem.lower()
    if stem_lower.endswith(".template"):
        base_stem = candidate.stem[: -len(".template")]
        json_path = candidate.with_name(f"{base_stem}.fillable.json")
        if json_path.exists():
            return json_path.resolve()
        raise FileNotFoundError(
            f"Could not find matching template JSON: {json_path}\n"
            f"Expected for template file: {candidate}"
        )

    return candidate


def run_cli(args: argparse.Namespace) -> int:
    config = load_config()
    config_changed = False

    if args.set_ai_mode:
        config.ai_mode = args.set_ai_mode
        config_changed = True
    if args.set_openai_model:
        config.openai_model = args.set_openai_model.strip()
        config_changed = True
    if args.set_openai_api_base:
        config.openai_api_base = args.set_openai_api_base.strip()
        config_changed = True
    if args.set_subscription_api_base is not None:
        config.subscription_api_base = args.set_subscription_api_base.strip()
        config_changed = True
    if config_changed:
        from app.utils import save_config

        save_config(config)
        print("Saved AI configuration.")

    if args.set_user_openai_key is not None:
        set_user_openai_key(USER_OPENAI_KEY_TARGET, args.set_user_openai_key.strip())
        print("Saved user OpenAI API key using Windows DPAPI.")
    if args.clear_user_openai_key:
        delete_user_openai_key(USER_OPENAI_KEY_TARGET)
        print("Cleared user OpenAI API key from Windows DPAPI storage.")
    if args.set_subscription_token is not None:
        set_secret(
            SUBSCRIPTION_TOKEN_TARGET,
            args.set_subscription_token.strip(),
            username="fillable-subscription",
        )
        print("Saved subscription token in Windows Credential Manager.")
    if args.clear_subscription_token:
        delete_secret(SUBSCRIPTION_TOKEN_TARGET)
        print("Cleared subscription token from Windows Credential Manager.")

    if args.install_context_menu:
        install_context_menu(
            exe_override=(args.context_menu_exe.strip() or None),
            prompt_instructions=not args.no_prompt_in_context_menu,
        )
        print("Context menu installed for current user.")
        return 0

    if args.uninstall_context_menu:
        uninstall_context_menu()
        print("Context menu removed for current user.")
        return 0

    if args.print_config:
        sanitized = AppConfig(
            ai_mode=config.ai_mode,
            openai_model=config.openai_model,
            openai_api_base=config.openai_api_base,
            subscription_api_base=config.subscription_api_base,
            codex_command_template=config.codex_command_template,
        )
        print(f"ai_mode={sanitized.ai_mode}")
        print(f"openai_model={sanitized.openai_model}")
        print(f"openai_api_base={sanitized.openai_api_base}")
        print(f"subscription_api_base={sanitized.subscription_api_base}")
        print("user_key_in_dpapi_store=(hidden)")
        print("subscription_token_in_credential_manager=(hidden)")
        return 0

    def logger(message: str) -> None:
        print(message)

    if args.generate_template:
        codex = CodexCli(config)
        output = generate_template(args.generate_template, codex, log=logger)
        print(output)
        return 0

    if args.import_template_file:
        codex = None
        try:
            codex = CodexCli(config)
        except Exception:
            codex = None
        output = create_template_from_user_template(args.import_template_file, codex=codex, log=logger)
        print(output)
        return 0

    if args.fill_template:
        fill_target = _resolve_fill_template_target(args.fill_template)
        instructions = args.instructions
        if args.prompt_instructions:
            try:
                print("Enter instructions for AI fill (press Enter to skip):")
                entered = input("> ").strip()
            except EOFError:
                entered = ""
            if entered:
                instructions = f"{instructions}\n{entered}".strip() if instructions.strip() else entered

        if args.batch_data:
            codex = CodexCli(config)
            context_files = [Path(p) for p in args.context.split(";") if p.strip()]
            outputs = fill_template_multiple(
                fill_target,
                codex,
                args.batch_data,
                context_files=context_files,
                extra_instructions=instructions,
                output_dir=args.batch_output_dir,
                log=logger,
            )
            for path in outputs:
                print(path)
            return 0

        codex = CodexCli(config)
        context_files = [Path(p) for p in args.context.split(";") if p.strip()]
        output = fill_template(
            fill_target,
            codex,
            context_files=context_files,
            extra_instructions=instructions,
            log=logger,
        )
        print(output)
        return 0

    return 1
