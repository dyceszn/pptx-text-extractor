"""Custom exception definitions for pptx extraction."""

class PPTXExtractionError(Exception):
    """Base exception for all PPTX extraction errors."""
    pass


class FileNotFoundPPTXError(PPTXExtractionError):
    """Raised when the specified PPTX file or path does not exist."""
    pass


class InvalidPPTXError(PPTXExtractionError):
    """Raised when the specified file is not a valid PPTX format."""
    pass


class EncryptedPPTXError(PPTXExtractionError):
    """Raised when the PPTX file is encrypted or password protected."""
    pass


class CorruptedPPTXError(PPTXExtractionError):
    """Raised when the PPTX file is corrupted and cannot be parsed."""
    pass


class OCRError(PPTXExtractionError):
    """Raised when OCR processing fails or OCR dependencies are missing."""
    pass
