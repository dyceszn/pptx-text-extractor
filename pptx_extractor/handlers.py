"""Shape handlers for extracting text from different PowerPoint elements."""

import io
import logging
from typing import Any, Dict, List, Optional, Tuple
from xml.etree import ElementTree as ET

from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.shapes.base import BaseShape

from pptx_extractor.models import BoundingBox, ElementType, ExtractedElement
from pptx_extractor.table_formatter import format_table

logger = logging.getLogger(__name__)

# DrawingML namespace for raw XML inspection
DRAWINGML_NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}


def get_shape_bbox(shape: BaseShape) -> BoundingBox:
    """Safely retrieve shape bounding box."""
    return BoundingBox(
        top=getattr(shape, "top", None),
        left=getattr(shape, "left", None),
        width=getattr(shape, "width", None),
        height=getattr(shape, "height", None),
    )


def extract_text_frame(shape: BaseShape, is_title: bool = False) -> Optional[ExtractedElement]:
    """Extract formatted text from a shape's text frame."""
    if not getattr(shape, "has_text_frame", False) or not shape.text_frame:
        return None

    paragraphs_text: List[str] = []
    for paragraph in shape.text_frame.paragraphs:
        p_text = paragraph.text.strip()
        if p_text:
            paragraphs_text.append(p_text)

    if not paragraphs_text:
        return None

    full_text = "\n".join(paragraphs_text)
    elem_type = ElementType.TITLE if is_title else ElementType.TEXT
    return ExtractedElement(
        element_type=elem_type,
        text=full_text,
        bbox=get_shape_bbox(shape),
    )


def extract_table(shape: BaseShape) -> Optional[ExtractedElement]:
    """Extract and format tabular data from a table shape."""
    if not getattr(shape, "has_table", False) or not shape.table:
        return None

    table = shape.table
    raw_rows: List[List[str]] = []
    for row in table.rows:
        row_cells: List[str] = []
        for cell in row.cells:
            # Get text from cell text frame
            cell_text = cell.text.strip() if cell.text else ""
            row_cells.append(cell_text)
        raw_rows.append(row_cells)

    # Check if table has any non-empty text
    has_text = any(any(cell.strip() for cell in row) for row in raw_rows)
    if not has_text:
        return None

    formatted = format_table(raw_rows)
    return ExtractedElement(
        element_type=ElementType.TABLE,
        text=formatted,
        bbox=get_shape_bbox(shape),
        metadata={"rows": raw_rows, "row_count": len(raw_rows), "col_count": len(table.columns)},
    )


def extract_chart(shape: BaseShape) -> Optional[ExtractedElement]:
    """Extract accessible text, titles, categories, and series from a chart shape."""
    if not getattr(shape, "has_chart", False) or not shape.chart:
        return None

    chart = shape.chart
    parts: List[str] = []

    # Chart Title
    title_text = ""
    try:
        if getattr(chart, "has_title", False) and chart.chart_title and chart.chart_title.text_frame:
            title_text = chart.chart_title.text_frame.text.strip()
            if title_text:
                parts.append(f"[Chart: {title_text}]")
    except Exception as e:
        logger.debug(f"Failed to extract chart title: {e}")

    if not parts:
        parts.append("[Chart]")

    # Extract Series and Categories
    series_names: List[str] = []
    try:
        for series in chart.series:
            name = getattr(series, "name", None)
            if name:
                series_names.append(str(name).strip())
    except Exception as e:
        logger.debug(f"Failed to extract chart series: {e}")

    if series_names:
        parts.append(f"Series: {', '.join(series_names)}")

    try:
        plots = getattr(chart, "plots", [])
        for plot in plots:
            categories = getattr(plot, "categories", None)
            if categories:
                cat_names = [str(c).strip() for c in categories if str(c).strip()]
                if cat_names:
                    parts.append(f"Categories: {', '.join(cat_names)}")
                    break
    except Exception as e:
        logger.debug(f"Failed to extract chart categories: {e}")

    full_text = "\n".join(parts)
    return ExtractedElement(
        element_type=ElementType.CHART,
        text=full_text,
        bbox=get_shape_bbox(shape),
        metadata={"title": title_text, "series": series_names},
    )


def extract_drawingml_fallback(shape: BaseShape) -> Optional[ExtractedElement]:
    """Extract any DrawingML text (<a:t>) directly from underlying XML (useful for SmartArt)."""
    try:
        xml_element = getattr(shape, "element", None)
        if xml_element is None:
            return None

        # Search for all <a:t> text nodes in the element tree
        text_nodes = xml_element.findall(".//a:t", DRAWINGML_NS)
        if not text_nodes:
            return None

        texts: List[str] = []
        for node in text_nodes:
            if node.text and node.text.strip():
                texts.append(node.text.strip())

        if not texts:
            return None

        combined_text = "\n".join(texts)
        return ExtractedElement(
            element_type=ElementType.SMARTART,
            text=combined_text,
            bbox=get_shape_bbox(shape),
        )
    except Exception as e:
        logger.debug(f"XML fallback extraction failed: {e}")
        return None


def extract_image_ocr(shape: BaseShape) -> Optional[ExtractedElement]:
    """Perform local OCR on an image shape if Tesseract is available."""
    try:
        if not (shape.shape_type == MSO_SHAPE_TYPE.PICTURE or hasattr(shape, "image")):
            return None

        image = shape.image
        blob = image.blob
        if not blob:
            return None

        try:
            from PIL import Image
            import pytesseract
        except ImportError:
            logger.warning("Pillow and pytesseract must be installed to use OCR.")
            return None

        pil_image = Image.open(io.BytesIO(blob))
        ocr_text = pytesseract.image_to_string(pil_image).strip()

        if not ocr_text:
            return None

        formatted_ocr = f"[OCR Text]\n{ocr_text}"
        return ExtractedElement(
            element_type=ElementType.IMAGE_OCR,
            text=formatted_ocr,
            bbox=get_shape_bbox(shape),
            metadata={"ocr": True},
        )
    except Exception as e:
        logger.warning(f"OCR extraction failed for shape: {e}")
        return None


def is_title_shape(shape: BaseShape, slide: Any) -> bool:
    """Check if the given shape is a slide title."""
    try:
        if hasattr(slide.shapes, "title") and slide.shapes.title == shape:
            return True
        if shape.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER:
            ph_type = shape.placeholder_format.type
            if ph_type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE):
                return True
    except Exception:
        pass
    return False


def extract_shape_recursive(
    shape: BaseShape,
    slide: Any,
    enable_ocr: bool = False,
) -> List[ExtractedElement]:
    """Extract elements from a shape, recursing into group shapes."""
    elements: List[ExtractedElement] = []

    # 1. Group shapes
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP or hasattr(shape, "shapes"):
        try:
            child_shapes = list(shape.shapes)
            # Sort child shapes by position inside group
            child_shapes.sort(
                key=lambda s: (
                    getattr(s, "top", 0) if getattr(s, "top", None) is not None else 0,
                    getattr(s, "left", 0) if getattr(s, "left", None) is not None else 0,
                )
            )
            for child in child_shapes:
                elements.extend(extract_shape_recursive(child, slide, enable_ocr=enable_ocr))
            return elements
        except Exception as e:
            logger.debug(f"Failed to process group shape: {e}")

    # 2. Table
    if getattr(shape, "has_table", False):
        tbl_elem = extract_table(shape)
        if tbl_elem:
            elements.append(tbl_elem)
            return elements

    # 3. Chart
    if getattr(shape, "has_chart", False):
        chart_elem = extract_chart(shape)
        if chart_elem:
            elements.append(chart_elem)
            return elements

    # 4. Standard Text Frame
    if getattr(shape, "has_text_frame", False):
        is_title = is_title_shape(shape, slide)
        text_elem = extract_text_frame(shape, is_title=is_title)
        if text_elem:
            elements.append(text_elem)
            return elements

    # 5. Optional Image OCR
    if enable_ocr and (shape.shape_type == MSO_SHAPE_TYPE.PICTURE or hasattr(shape, "image")):
        ocr_elem = extract_image_ocr(shape)
        if ocr_elem:
            elements.append(ocr_elem)
            return elements

    # 6. Fallback XML inspection (SmartArt, graphic frames, ungrouped DrawingML)
    fallback_elem = extract_drawingml_fallback(shape)
    if fallback_elem:
        elements.append(fallback_elem)

    return elements
