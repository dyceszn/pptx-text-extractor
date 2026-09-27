"""PPTX Text Extractor package."""

from pptx_extractor.errors import (
    CorruptedPPTXError,
    EncryptedPPTXError,
    FileNotFoundPPTXError,
    InvalidPPTXError,
    OCRError,
    PPTXExtractionError,
)
from pptx_extractor.extractor import ExtractionConfig, PPTXExtractor
from pptx_extractor.models import (
    BoundingBox,
    ElementType,
    ExtractedElement,
    PresentationResult,
    SlideContent,
)

__all__ = [
    "PPTXExtractor",
    "ExtractionConfig",
    "SlideContent",
    "PresentationResult",
    "ExtractedElement",
    "ElementType",
    "BoundingBox",
    "PPTXExtractionError",
    "FileNotFoundPPTXError",
    "InvalidPPTXError",
    "EncryptedPPTXError",
    "CorruptedPPTXError",
    "OCRError",
]
