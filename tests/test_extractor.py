"""Unit tests for PPTXExtractor and shape handlers."""

from pathlib import Path
import pytest

from pptx_extractor.errors import (
    CorruptedPPTXError,
    EncryptedPPTXError,
    FileNotFoundPPTXError,
    InvalidPPTXError,
)
from pptx_extractor.extractor import ExtractionConfig, PPTXExtractor
from pptx_extractor.models import ElementType, PresentationResult


def test_basic_text_and_reading_order(comprehensive_pptx: Path):
    """Test standard text extraction and reading order (top-to-bottom)."""
    extractor = PPTXExtractor()
    result: PresentationResult = extractor.extract(str(comprehensive_pptx))

    assert result.slide_count == 8

    # Slide 1: Title and Subtitle
    s1 = result.slides[0]
    assert s1.slide_number == 1
    s1_text = s1.to_text()
    assert "Introduction to Artificial Intelligence" in s1_text
    assert "Fundamentals and Applications" in s1_text

    # Slide 2: Reading order verification
    # Shape A (top=1 inch) was added second in code, Shape B (top=3 inches) was added first.
    # Reading order must ensure Top Paragraph appears BEFORE Bottom Paragraph.
    s2 = result.slides[1]
    s2_text = s2.to_text()
    top_pos = s2_text.find("Top Paragraph")
    bottom_pos = s2_text.find("Bottom Paragraph")
    assert top_pos != -1
    assert bottom_pos != -1
    assert top_pos < bottom_pos, "Top paragraph should be read before bottom paragraph"


def test_table_extraction(comprehensive_pptx: Path):
    """Test extracting and formatting tables."""
    extractor = PPTXExtractor()
    result = extractor.extract(str(comprehensive_pptx))

    s3 = result.slides[2]
    s3_text = s3.to_text()

    assert "Name" in s3_text
    assert "Score" in s3_text
    assert "Grade" in s3_text
    assert "Alice" in s3_text
    assert "Bob" in s3_text

    # Check table element metadata
    table_elems = [e for e in s3.elements if e.element_type == ElementType.TABLE]
    assert len(table_elems) == 1
    assert table_elems[0].metadata["row_count"] == 3
    assert table_elems[0].metadata["col_count"] == 3


def test_grouped_shapes_extraction(comprehensive_pptx: Path):
    """Test extracting text from within grouped shapes."""
    extractor = PPTXExtractor()
    result = extractor.extract(str(comprehensive_pptx))

    s4 = result.slides[3]
    s4_text = s4.to_text()

    assert "Grouped Box One" in s4_text
    assert "Grouped Box Two" in s4_text


def test_speaker_notes_toggle(comprehensive_pptx: Path):
    """Test that speaker notes are only included when configured."""
    # When include_notes is False (default)
    extractor_no_notes = PPTXExtractor(config=ExtractionConfig(include_notes=False))
    res_no_notes = extractor_no_notes.extract(str(comprehensive_pptx))
    s5_no_notes = res_no_notes.slides[4].to_text(include_notes=False)
    assert "Remember to emphasize key milestones" not in s5_no_notes

    # When include_notes is True
    extractor_with_notes = PPTXExtractor(config=ExtractionConfig(include_notes=True))
    res_notes = extractor_with_notes.extract(str(comprehensive_pptx))
    s5_notes = res_notes.slides[4].to_text(include_notes=True)
    assert "--- Notes ---" in s5_notes
    assert "Remember to emphasize key milestones" in s5_notes


def test_unicode_preservation(comprehensive_pptx: Path):
    """Test that Unicode characters, accents, non-Latin scripts, and symbols are preserved."""
    extractor = PPTXExtractor()
    result = extractor.extract(str(comprehensive_pptx))

    s6 = result.slides[5]
    s6_text = s6.to_text()

    assert "Café résumé naïve" in s6_text
    assert "人工智能" in s6_text
    assert "機械学習" in s6_text
    assert "∑(x) = λ × π" in s6_text
    assert "🚀" in s6_text


def test_empty_slide_handling(comprehensive_pptx: Path, empty_slides_pptx: Path):
    """Test that empty slides are explicitly reported as [Empty Slide]."""
    extractor = PPTXExtractor()
    result = extractor.extract(str(comprehensive_pptx))

    s7 = result.slides[6]
    assert s7.to_text() == "[Empty Slide]"

    # Presentation with only empty slides
    empty_res = extractor.extract(str(empty_slides_pptx))
    for s in empty_res.slides:
        assert s.to_text() == "[Empty Slide]"


def test_chart_extraction(comprehensive_pptx: Path):
    """Test extracting chart titles, series, and categories."""
    extractor = PPTXExtractor()
    result = extractor.extract(str(comprehensive_pptx))

    s8 = result.slides[7]
    s8_text = s8.to_text()

    assert "Quarterly Revenue" in s8_text
    assert "Revenue" in s8_text
    assert "Q1" in s8_text


def test_missing_file_error():
    """Test that nonexistent file raises FileNotFoundPPTXError."""
    extractor = PPTXExtractor()
    with pytest.raises(FileNotFoundPPTXError):
        extractor.extract("non_existent_file.pptx")


def test_corrupted_file_error(corrupt_pptx: Path):
    """Test that corrupt file raises CorruptedPPTXError or InvalidPPTXError."""
    extractor = PPTXExtractor()
    with pytest.raises((CorruptedPPTXError, InvalidPPTXError)):
        extractor.extract(str(corrupt_pptx))


def test_encrypted_file_error(encrypted_mock_pptx: Path):
    """Test that password-protected/encrypted file raises EncryptedPPTXError."""
    extractor = PPTXExtractor()
    with pytest.raises(EncryptedPPTXError):
        extractor.extract(str(encrypted_mock_pptx))


def test_picture_extraction_without_ocr(tmp_path: Path):
    """Test that slides with pictures extract cleanly without OCR enabled."""
    from PIL import Image
    from pptx import Presentation
    from pptx.util import Inches

    img_path = tmp_path / "sample.png"
    img = Image.new("RGB", (100, 100), color="blue")
    img.save(img_path)

    prs = Presentation()
    s = prs.slides.add_slide(prs.slide_layouts[6])
    tb = s.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    tb.text_frame.text = "Slide with image"
    s.shapes.add_picture(str(img_path), Inches(1), Inches(2), Inches(3), Inches(3))

    pres_path = tmp_path / "with_image.pptx"
    prs.save(str(pres_path))

    extractor = PPTXExtractor(config=ExtractionConfig(enable_ocr=False))
    result = extractor.extract(str(pres_path))
    assert result.slide_count == 1
    assert "Slide with image" in result.slides[0].to_text()


def test_ocr_mocked_extraction(tmp_path: Path, monkeypatch):
    """Test OCR output formatting when OCR is enabled and tesseract returns text."""
    from PIL import Image
    from pptx import Presentation
    from pptx.util import Inches

    img_path = tmp_path / "sample.png"
    img = Image.new("RGB", (100, 100), color="white")
    img.save(img_path)

    prs = Presentation()
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.shapes.add_picture(str(img_path), Inches(1), Inches(1), Inches(3), Inches(3))

    pres_path = tmp_path / "with_ocr.pptx"
    prs.save(str(pres_path))

    # Mock pytesseract.image_to_string
    import pytesseract
    monkeypatch.setattr(pytesseract, "image_to_string", lambda *args, **kwargs: "Detected Text In Image")

    extractor = PPTXExtractor(config=ExtractionConfig(enable_ocr=True))
    result = extractor.extract(str(pres_path))
    s1_text = result.slides[0].to_text()
    assert "[OCR Text]" in s1_text
    assert "Detected Text In Image" in s1_text
