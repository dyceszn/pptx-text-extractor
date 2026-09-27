"""Data models for extracted PowerPoint content."""

from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Dict, List, Optional


class ElementType(str, Enum):
    """Types of extracted slide elements."""
    TITLE = "title"
    TEXT = "text"
    TABLE = "table"
    GROUP = "group"
    CHART = "chart"
    SMARTART = "smartart"
    IMAGE_OCR = "image_ocr"
    NOTES = "notes"


@dataclass
class BoundingBox:
    """Position and dimension of an element on a slide (in EMUs)."""
    top: Optional[int] = None
    left: Optional[int] = None
    width: Optional[int] = None
    height: Optional[int] = None

    def to_dict(self) -> Dict[str, Optional[int]]:
        return {
            "top": self.top,
            "left": self.left,
            "width": self.width,
            "height": self.height,
        }


@dataclass
class ExtractedElement:
    """An individual text-containing element extracted from a slide."""
    element_type: ElementType
    text: str
    bbox: Optional[BoundingBox] = None
    metadata: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        result: Dict[str, Any] = {
            "type": self.element_type.value,
            "text": self.text,
        }
        if self.bbox is not None:
            result["bbox"] = self.bbox.to_dict()
        if self.metadata:
            result["metadata"] = self.metadata
        return result


@dataclass
class SlideContent:
    """Content extracted from a single slide."""
    slide_number: int
    elements: List[ExtractedElement] = field(default_factory=list)
    notes: Optional[str] = None

    def to_text(self, include_notes: bool = True) -> str:
        """Format slide content as readable text."""
        blocks: List[str] = []
        for elem in self.elements:
            clean_text = elem.text.strip()
            if clean_text:
                blocks.append(clean_text)

        if not blocks and (not self.notes or not include_notes):
            body_text = "[Empty Slide]"
        else:
            body_text = "\n\n".join(blocks) if blocks else "[Empty Slide]"

        if include_notes and self.notes and self.notes.strip():
            notes_section = f"\n\n--- Notes ---\n{self.notes.strip()}"
            body_text += notes_section

        return body_text

    def to_dict(self, include_notes: bool = True) -> Dict[str, Any]:
        """Convert slide content to a dictionary for JSON output."""
        data: Dict[str, Any] = {
            "slide_number": self.slide_number,
            "text": self.to_text(include_notes=include_notes),
            "elements": [elem.to_dict() for elem in self.elements],
        }
        if include_notes:
            data["notes"] = self.notes.strip() if self.notes else None
        return data


@dataclass
class PresentationResult:
    """Full extraction results for a presentation."""
    source_file: str
    slide_count: int
    slides: List[SlideContent] = field(default_factory=list)

    def to_text(self, include_notes: bool = True) -> str:
        """Render all slides as structured plain text."""
        slide_blocks: List[str] = []
        for slide in self.slides:
            slide_header = f"===== SLIDE {slide.slide_number} ====="
            slide_body = slide.to_text(include_notes=include_notes)
            slide_blocks.append(f"{slide_header}\n\n{slide_body}")

        return "\n\n".join(slide_blocks) + "\n" if slide_blocks else ""

    def to_dict(self, include_notes: bool = True) -> Dict[str, Any]:
        """Convert full presentation result to structured dictionary."""
        return {
            "presentation": self.source_file,
            "slide_count": self.slide_count,
            "slides": [slide.to_dict(include_notes=include_notes) for slide in self.slides],
        }
