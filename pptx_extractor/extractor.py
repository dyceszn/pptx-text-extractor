"""Main PPTX extraction engine."""

import os
import zipfile
from dataclasses import dataclass
from typing import List, Optional

from pptx import Presentation
from pptx.exc import PackageNotFoundError

from pptx_extractor.errors import (
    CorruptedPPTXError,
    EncryptedPPTXError,
    FileNotFoundPPTXError,
    InvalidPPTXError,
    PPTXExtractionError,
)
from pptx_extractor.handlers import extract_shape_recursive, is_title_shape
from pptx_extractor.models import ElementType, ExtractedElement, PresentationResult, SlideContent

# OLE CFBF signature (used for encrypted Office Open XML packages and legacy Office files)
OLE_HEADER = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
# Standard ZIP header
ZIP_HEADER = b"PK\x03\x04"


@dataclass
class ExtractionConfig:
    """Configuration options for PPTX extraction."""
    include_notes: bool = False
    enable_ocr: bool = False
    row_tolerance_emu: int = 180000  # ~0.2 inch tolerance for row-level alignment


class PPTXExtractor:
    """Extracts text content from a PowerPoint (.pptx) presentation."""

    def __init__(self, config: Optional[ExtractionConfig] = None):
        self.config = config or ExtractionConfig()

    def validate_file(self, file_path: str) -> None:
        """Validate that the file exists, is accessible, and is a valid PPTX package."""
        if not os.path.exists(file_path):
            raise FileNotFoundPPTXError(f"File not found: '{file_path}'")

        if os.path.isdir(file_path):
            raise InvalidPPTXError(f"Expected a file, but found a directory: '{file_path}'")

        if not os.path.isfile(file_path):
            raise InvalidPPTXError(f"Path is not a regular file: '{file_path}'")

        if os.path.getsize(file_path) == 0:
            raise CorruptedPPTXError(f"File is empty (0 bytes): '{file_path}'")

        # Check file signature
        try:
            with open(file_path, "rb") as f:
                header = f.read(8)
        except Exception as e:
            raise PPTXExtractionError(f"Cannot read file '{file_path}': {e}")

        # Check for password-protected / encrypted Office file (stored as OLE compound file)
        if header.startswith(OLE_HEADER):
            raise EncryptedPPTXError(
                f"File '{file_path}' appears to be password-protected/encrypted (or legacy .ppt format). "
                "Encrypted presentations cannot be extracted."
            )

        if not zipfile.is_zipfile(file_path):
            raise InvalidPPTXError(
                f"File '{file_path}' is not a valid PowerPoint (.pptx) file (not a valid ZIP package)."
            )

        # Quick check for pptx structure inside the zip
        try:
            with zipfile.ZipFile(file_path, "r") as zf:
                namelist = zf.namelist()
                if not any(name.startswith("ppt/") for name in namelist):
                    raise InvalidPPTXError(
                        f"File '{file_path}' is a ZIP archive, but lacks PowerPoint presentation contents ('ppt/')."
                    )
        except zipfile.BadZipFile:
            raise CorruptedPPTXError(f"File '{file_path}' is corrupted and could not be opened as a ZIP package.")

    def extract(self, file_path: str) -> PresentationResult:
        """Extract text from the PowerPoint presentation file."""
        self.validate_file(file_path)

        try:
            prs = Presentation(file_path)
        except PackageNotFoundError:
            raise CorruptedPPTXError(f"PowerPoint package error while opening '{file_path}'.")
        except Exception as e:
            # Handle potential internal corruption or XML parse failures
            raise CorruptedPPTXError(f"Failed to open '{file_path}': {e}")

        slides_content: List[SlideContent] = []

        for slide_idx, slide in enumerate(prs.slides, start=1):
            slide_content = self._extract_slide(slide, slide_idx)
            slides_content.append(slide_content)

        return PresentationResult(
            source_file=os.path.abspath(file_path),
            slide_count=len(slides_content),
            slides=slides_content,
        )

    def _extract_slide(self, slide, slide_number: int) -> SlideContent:
        """Extract all content from a single slide adhering to reading order."""
        # Retrieve all shapes on the slide
        shapes = list(slide.shapes)

        # Sort shapes into natural reading order:
        # Title shapes always come first (-1 priority)
        # Other shapes are sorted top-to-bottom (using row tolerance bands), then left-to-right
        def shape_sort_key(s):
            is_title = is_title_shape(s, slide)
            if is_title:
                return (-1, 0, 0)
            top = getattr(s, "top", None) or 0
            left = getattr(s, "left", None) or 0
            row_bucket = top // self.config.row_tolerance_emu
            return (0, row_bucket, left)

        sorted_shapes = sorted(shapes, key=shape_sort_key)

        elements: List[ExtractedElement] = []
        for shape in sorted_shapes:
            shape_elements = extract_shape_recursive(
                shape,
                slide,
                enable_ocr=self.config.enable_ocr,
            )
            elements.extend(shape_elements)

        # Extract speaker notes if available
        notes_text: Optional[str] = None
        if self.config.include_notes:
            notes_text = self._extract_notes(slide)

        return SlideContent(
            slide_number=slide_number,
            elements=elements,
            notes=notes_text,
        )

    def _extract_notes(self, slide) -> Optional[str]:
        """Safely extract speaker notes from a slide."""
        try:
            if not getattr(slide, "has_notes_slide", False):
                return None
            notes_slide = slide.notes_slide
            if not notes_slide:
                return None
            notes_text_frame = notes_slide.notes_text_frame
            if not notes_text_frame:
                return None
            text = notes_text_frame.text.strip()
            return text if text else None
        except Exception:
            return None
