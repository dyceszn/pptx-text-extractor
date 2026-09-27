"""Test fixtures and synthetic presentation generators."""

import os
from pathlib import Path
import pytest
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches, Pt


@pytest.fixture
def comprehensive_pptx(tmp_path: Path) -> Path:
    """Create a comprehensive PowerPoint presentation with various shape types."""
    prs = Presentation()
    blank_layout = prs.slide_layouts[6]
    title_layout = prs.slide_layouts[0]

    # Slide 1: Title slide
    s1 = prs.slides.add_slide(title_layout)
    s1.shapes.title.text = "Introduction to Artificial Intelligence"
    s1.placeholders[1].text = "Fundamentals and Applications"

    # Slide 2: Multi-paragraph and positional order
    s2 = prs.slides.add_slide(blank_layout)
    # Shape A (Lower on slide: top=3 inches)
    tb_b = s2.shapes.add_textbox(Inches(1), Inches(3), Inches(5), Inches(1))
    tb_b.text_frame.text = "Bottom Paragraph: Future directions in AI."
    # Shape B (Higher on slide: top=1 inch)
    tb_a = s2.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(1))
    tb_a.text_frame.text = "Top Paragraph: Neural networks are modeled on biological brains."

    # Slide 3: Table
    s3 = prs.slides.add_slide(blank_layout)
    table_shape = s3.shapes.add_table(3, 3, Inches(1), Inches(1), Inches(5), Inches(2))
    table = table_shape.table
    table.cell(0, 0).text = "Name"
    table.cell(0, 1).text = "Score"
    table.cell(0, 2).text = "Grade"
    table.cell(1, 0).text = "Alice"
    table.cell(1, 1).text = "95"
    table.cell(1, 2).text = "A"
    table.cell(2, 0).text = "Bob"
    table.cell(2, 1).text = "82"
    table.cell(2, 2).text = "B"

    # Slide 4: Grouped shape
    s4 = prs.slides.add_slide(blank_layout)
    group = s4.shapes.add_group_shape()
    g_tb1 = group.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
    g_tb1.text_frame.text = "Grouped Box One"
    g_tb2 = group.shapes.add_textbox(Inches(1), Inches(2), Inches(3), Inches(1))
    g_tb2.text_frame.text = "Grouped Box Two"

    # Slide 5: Speaker notes
    s5 = prs.slides.add_slide(blank_layout)
    s5_tb = s5.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    s5_tb.text_frame.text = "Slide with Speaker Notes"
    s5.notes_slide.notes_text_frame.text = "Remember to emphasize key milestones on this slide."

    # Slide 6: Unicode & Multilingual content
    s6 = prs.slides.add_slide(blank_layout)
    s6_tb = s6.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(2))
    s6_tb.text_frame.text = (
        "Café résumé naïve\n"
        "人工智能 (AI in Chinese)\n"
        "機械学習 (Machine Learning in Japanese)\n"
        "Symbols: ∑(x) = λ × π, 🚀 2026"
    )

    # Slide 7: Completely Empty Slide
    prs.slides.add_slide(blank_layout)

    # Slide 8: Chart
    s8 = prs.slides.add_slide(blank_layout)
    chart_data = CategoryChartData()
    chart_data.categories = ["Q1", "Q2", "Q3", "Q4"]
    chart_data.add_series("Revenue", (100, 150, 180, 220))
    chart_shape = s8.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(1),
        Inches(1),
        Inches(6),
        Inches(4),
        chart_data,
    )
    chart_shape.chart.has_title = True
    chart_shape.chart.chart_title.text_frame.text = "Quarterly Revenue"

    dest_file = tmp_path / "comprehensive_test.pptx"
    prs.save(str(dest_file))
    return dest_file


@pytest.fixture
def corrupt_pptx(tmp_path: Path) -> Path:
    """Create a corrupt file with PPTX extension."""
    dest = tmp_path / "corrupt.pptx"
    with open(dest, "wb") as f:
        f.write(b"This is not a zip or pptx file at all. Just random junk.")
    return dest


@pytest.fixture
def encrypted_mock_pptx(tmp_path: Path) -> Path:
    """Create a mock file starting with OLE header (used for password-protected Office files)."""
    dest = tmp_path / "encrypted.pptx"
    with open(dest, "wb") as f:
        # Standard OLE header followed by padding
        f.write(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 512)
    return dest


@pytest.fixture
def empty_slides_pptx(tmp_path: Path) -> Path:
    """Create a presentation with only empty slides."""
    prs = Presentation()
    blank_layout = prs.slide_layouts[6]
    prs.slides.add_slide(blank_layout)
    prs.slides.add_slide(blank_layout)
    dest = tmp_path / "empty_slides.pptx"
    prs.save(str(dest))
    return dest
