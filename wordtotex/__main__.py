import argparse
from pathlib import Path

from .converter import ConverterConfig, DocxToLatexConverter


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Convert a Word .docx file to LaTeX .tex")
    parser.add_argument("input", type=Path, help="Path to the .docx file to convert")
    parser.add_argument("-o", "--output", type=Path, help="Where to write the .tex output")
    parser.add_argument(
        "--no-preamble",
        action="store_true",
        help="Emit only document body without \\documentclass preamble",
    )
    parser.add_argument(
        "--no-table-border",
        action="store_true",
        help="Render tables without surrounding vertical borders",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if not args.input.exists():
        raise FileNotFoundError(f"Input file not found: {args.input}")

    output_dir = args.output.parent if args.output else args.input.parent
    config = ConverterConfig(
        include_preamble=not args.no_preamble,
        table_border=not args.no_table_border,
    )
    converter = DocxToLatexConverter(config=config)
    result = converter.convert(args.input, output_dir=output_dir)

    if args.output:
        args.output.write_text(result.latex, encoding="utf-8")
    else:
        print(result.latex)


if __name__ == "__main__":
    main()
