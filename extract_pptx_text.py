#!/usr/bin/env python3
"""extract_pptx_text.py - Standalone local tool to extract text from PowerPoint (.pptx) files."""

import argparse
import glob
import json
import logging
import os
import shutil
import sys
from pathlib import Path
from typing import List, Optional

# Ensure local package can be imported
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from pptx_extractor.errors import (
    CorruptedPPTXError,
    EncryptedPPTXError,
    FileNotFoundPPTXError,
    InvalidPPTXError,
    PPTXExtractionError,
)
from pptx_extractor.extractor import ExtractionConfig, PPTXExtractor
from pptx_extractor.models import PresentationResult


def setup_logging(verbose: bool = False) -> None:
    level = logging.DEBUG if verbose else logging.WARNING
    logging.basicConfig(
        format="%(levelname)s: %(message)s",
        level=level,
    )


def check_ocr_prerequisites() -> bool:
    """Check if Tesseract binary and required Python libraries are available."""
    if not shutil.which("tesseract"):
        sys.stderr.write(
            "Warning: 'tesseract' executable was not found in PATH.\n"
            "To use OCR, install Tesseract (e.g. 'brew install tesseract' on macOS or 'apt install tesseract-ocr' on Linux).\n"
            "Continuing extraction without OCR...\n"
        )
        return False
    try:
        import PIL  # noqa: F401
        import pytesseract  # noqa: F401
        return True
    except ImportError:
        sys.stderr.write(
            "Warning: 'Pillow' and 'pytesseract' libraries are required for OCR.\n"
            "Install them via 'pip install Pillow pytesseract'.\n"
            "Continuing extraction without OCR...\n"
        )
        return False


def determine_output_path(input_path: Path, output_arg: Optional[str], is_json: bool) -> Path:
    """Determine destination file path for extraction output."""
    ext = ".json" if is_json else ".txt"

    if not output_arg:
        return input_path.with_suffix(ext)

    out_p = Path(output_arg)
    # If the user specified a directory or existing directory
    if out_p.is_dir() or output_arg.endswith(os.sep) or output_arg.endswith("/"):
        out_p.mkdir(parents=True, exist_ok=True)
        return out_p / (input_path.stem + ext)

    # Ensure parent directory exists
    out_p.parent.mkdir(parents=True, exist_ok=True)
    return out_p


def process_single_file(
    input_file: Path,
    output_file: Path,
    extractor: PPTXExtractor,
    is_json: bool,
    include_notes: bool,
) -> bool:
    """Extract text from one presentation and save to output file."""
    try:
        result: PresentationResult = extractor.extract(str(input_file))

        if is_json:
            content = json.dumps(result.to_dict(include_notes=include_notes), indent=2, ensure_ascii=False)
        else:
            content = result.to_text(include_notes=include_notes)

        with open(output_file, "w", encoding="utf-8") as f:
            f.write(content)

        print(f"Extracted {result.slide_count} slides from '{input_file.name}' -> '{output_file}'")
        return True

    except (FileNotFoundPPTXError, InvalidPPTXError, EncryptedPPTXError, CorruptedPPTXError) as e:
        sys.stderr.write(f"Error [{input_file.name}]: {e}\n")
        return False
    except Exception as e:
        sys.stderr.write(f"Unexpected error processing '{input_file}': {e}\n")
        return False


def collect_pptx_files(target_dir: Path, recursive: bool = False) -> List[Path]:
    """Find all .pptx files in target directory, excluding temporary office lock files (~$*)."""
    pattern = "**/*.pptx" if recursive else "*.pptx"
    files: List[Path] = []
    for p in target_dir.glob(pattern):
        # Exclude temporary Office lock files (e.g. ~$presentation.pptx)
        if not p.name.startswith("~$") and p.is_file():
            files.append(p)
    return sorted(files)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Extract all text content from PowerPoint (.pptx) presentations locally.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""Examples:
  python extract_pptx_text.py presentation.pptx
  python extract_pptx_text.py presentation.pptx -o output.txt
  python extract_pptx_text.py presentation.pptx --json
  python extract_pptx_text.py presentation.pptx --include-notes
  python extract_pptx_text.py presentation.pptx --ocr
  python extract_pptx_text.py /path/to/presentations_dir/ -o /path/to/output_dir/
""",
    )
    parser.add_argument(
        "input",
        type=str,
        help="Path to a .pptx file or a directory containing .pptx files.",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=str,
        default=None,
        help="Path to output text/json file (or output directory in batch mode).",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Output structured JSON instead of plain text.",
    )
    parser.add_argument(
        "--include-notes",
        action="store_true",
        help="Include speaker notes in the extracted output.",
    )
    parser.add_argument(
        "--ocr",
        action="store_true",
        help="Enable local OCR on embedded images (requires Tesseract).",
    )
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="Search recursively for .pptx files when input is a directory.",
    )
    parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="Enable verbose debugging logs.",
    )
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    setup_logging(args.verbose)

    input_path = Path(args.input)
    if not input_path.exists():
        sys.stderr.write(f"Error: Input path '{args.input}' does not exist.\n")
        return 1

    # Check OCR capability if requested
    enable_ocr = False
    if args.ocr:
        enable_ocr = check_ocr_prerequisites()

    config = ExtractionConfig(
        include_notes=args.include_notes,
        enable_ocr=enable_ocr,
    )
    extractor = PPTXExtractor(config=config)

    # Batch directory mode
    if input_path.is_dir():
        files = collect_pptx_files(input_path, recursive=args.recursive)
        if not files:
            sys.stderr.write(f"No .pptx files found in '{input_path}'.\n")
            return 1

        print(f"Found {len(files)} PowerPoint presentation(s) in '{input_path}'.")
        success_count = 0

        # Destination directory if provided
        output_dir = Path(args.output) if args.output else None
        if output_dir:
            output_dir.mkdir(parents=True, exist_ok=True)

        for pptx_file in files:
            ext = ".json" if args.json else ".txt"
            if output_dir:
                dest = output_dir / (pptx_file.stem + ext)
            else:
                dest = pptx_file.with_suffix(ext)

            if process_single_file(
                pptx_file,
                dest,
                extractor,
                is_json=args.json,
                include_notes=args.include_notes,
            ):
                success_count += 1

        print(f"Finished: {success_count}/{len(files)} presentations processed successfully.")
        return 0 if success_count == len(files) else 1

    # Single file mode
    out_path = determine_output_path(input_path, args.output, args.json)
    success = process_single_file(
        input_path,
        out_path,
        extractor,
        is_json=args.json,
        include_notes=args.include_notes,
    )
    return 0 if success else 1


if __name__ == "__main__":
    sys.exit(main())
