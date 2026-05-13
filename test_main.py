#!/usr/bin/env python3
"""
Local manual test script for the DOCX Templater tool.

Usage:
    python test_main.py <template.docx> <context.json> [output.docx]

Example:
    python test_main.py template.docx context.json output.docx
"""

import argparse
import io
import json
import sys
from pathlib import Path

from docxtpl import DocxTemplate


def render_docx(template_path: str, context: dict, output_path: str) -> None:
    """Render a .docx template with the given context and save to output_path."""
    doc = DocxTemplate(template_path)
    doc.render(context)
    doc.save(output_path)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Render a .docx template with JSON context data."
    )
    parser.add_argument("template", help="Path to the .docx template file")
    parser.add_argument("context", help="Path to the JSON context file")
    parser.add_argument(
        "output",
        nargs="?",
        default="output.docx",
        help="Path for the generated .docx file (default: output.docx)",
    )
    args = parser.parse_args()

    template_path = Path(args.template)
    context_path = Path(args.context)
    output_path = Path(args.output)

    if not template_path.exists():
        print(f"Error: Template file not found: {template_path}", file=sys.stderr)
        return 1

    if not context_path.exists():
        print(f"Error: Context file not found: {context_path}", file=sys.stderr)
        return 1

    if template_path.suffix.lower() != ".docx":
        print("Error: Template file must have .docx extension", file=sys.stderr)
        return 1

    try:
        with open(context_path, "r", encoding="utf-8") as f:
            context = json.load(f)
    except json.JSONDecodeError as e:
        print(f"Error: Invalid JSON in context file - {e}", file=sys.stderr)
        return 1

    try:
        render_docx(str(template_path), context, str(output_path))
        print(f"Successfully generated: {output_path}")
    except Exception as e:
        print(f"Error rendering template: {e}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
