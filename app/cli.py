from __future__ import annotations

import argparse
from pathlib import Path

from app.codex_cli import CodexCli
from app.template_engine import (
    create_template_from_user_template,
    fill_template,
    fill_template_multiple,
    generate_template,
)
from app.utils import load_config


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
    if args.print_config:
        print(f"codex_command_template={config.codex_command_template}")
        return 0

    def logger(message: str) -> None:
        print(message)

    if args.generate_template:
        codex = CodexCli(config)
        output = generate_template(args.generate_template, codex, log=logger)
        print(output)
        return 0

    if args.import_template_file:
        output = create_template_from_user_template(args.import_template_file, log=logger)
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
