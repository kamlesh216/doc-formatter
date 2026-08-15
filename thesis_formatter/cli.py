"""
cli.py — Command-line interface for thesis_formatter.
"""
from __future__ import annotations

import argparse
import logging
import os
import sys


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="format_thesis",
        description=(
            "Production-quality thesis DOCX formatter.\n"
            "Preserves all content; fixes spacing, fonts, images, and tables."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "input",
        metavar="input.docx",
        help="Path to the input DOCX file.",
    )
    parser.add_argument(
        "output",
        metavar="output.docx",
        nargs="?",
        default=None,
        help=(
            "Path for the formatted output DOCX. "
            "Defaults to <input>_formatted.docx."
        ),
    )
    parser.add_argument(
        "--report",
        action="store_true",
        default=False,
        help="Print a detailed formatting report after completion.",
    )
    parser.add_argument(
        "--analysis",
        action="store_true",
        default=False,
        help="Print document analysis before formatting.",
    )
    parser.add_argument(
        "--validate",
        action="store_true",
        default=True,
        help="Run post-save validation (default: on).",
    )
    parser.add_argument(
        "--no-validate",
        dest="validate",
        action="store_false",
        help="Skip post-save validation.",
    )
    parser.add_argument(
        "--keep-body-bold",
        action="store_true",
        default=False,
        help="Keep bold formatting in regular body text paragraphs.",
    )
    parser.add_argument(
        "--keep-shading",
        action="store_true",
        default=False,
        help="Keep background shading/highlighting in body paragraphs.",
    )
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        default=False,
        help="Enable verbose (DEBUG) logging.",
    )
    return parser


def default_output(input_path: str) -> str:
    base, ext = os.path.splitext(input_path)
    return base + "_formatted" + (ext or ".docx")


def main(argv=None) -> int:
    parser = build_parser()
    args   = parser.parse_args(argv)

    # Logging setup
    level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(levelname)s  %(name)s  %(message)s",
    )

    input_path  = args.input
    output_path = args.output or default_output(input_path)

    # Safety: never overwrite input
    if os.path.abspath(input_path) == os.path.abspath(output_path):
        print("Error: output path is the same as input path. "
              "Refusing to overwrite the original.", file=sys.stderr)
        return 1

    # Import here so CLI starts fast even if deps missing
    try:
        from thesis_formatter import config
        from thesis_formatter.formatter import format_document
        from thesis_formatter.validator import validate, print_validation

        if args.keep_body_bold:
            config.REMOVE_BODY_BOLD = False
        if args.keep_shading:
            config.REMOVE_BODY_SHADING = False
    except ImportError as exc:
        print(f"Import error: {exc}\nRun: pip install -r requirements.txt",
              file=sys.stderr)
        return 1

    try:
        report = format_document(
            input_path,
            output_path,
            show_analysis=args.analysis,
        )
    except FileNotFoundError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 2
    except Exception as exc:
        logging.exception("Unexpected error during formatting.")
        print(f"Unexpected error: {exc}", file=sys.stderr)
        return 3

    if args.report:
        report.print()

    if args.validate and report.analysis:
        warnings = validate(
            output_path,
            original_table_count  = report.analysis.table_count,
            original_para_count   = report.analysis.para_count,
            original_image_count  = report.analysis.image_count,
        )
        print_validation(warnings)
        if any("CRITICAL" in w for w in warnings):
            return 4

    print(f"\nDone: {input_path}  →  {output_path}")
    return 0
