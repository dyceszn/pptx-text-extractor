"""Formatting utilities for PowerPoint tables."""

from typing import List


def format_table(rows: List[List[str]]) -> str:
    """Format a 2D grid of cell strings into an aligned text table.

    Example output:
    Name | Score | Grade
    -----+-------+------
    John | 85    | A
    Jane | 78    | B
    """
    if not rows:
        return ""

    # Clean cell texts: collapse internal line breaks to spaces or semicolons, strip whitespace
    cleaned_rows: List[List[str]] = []
    max_cols = 0
    for row in rows:
        cleaned_row = []
        for cell in row:
            # Replace newlines with spaces within a cell for clean tabular layout
            cleaned_cell = " ".join(cell.split())
            cleaned_row.append(cleaned_cell)
        cleaned_rows.append(cleaned_row)
        if len(cleaned_row) > max_cols:
            max_cols = len(cleaned_row)

    if max_cols == 0:
        return ""

    # Normalize row lengths to max_cols
    for row in cleaned_rows:
        while len(row) < max_cols:
            row.append("")

    # Determine maximum column widths (minimum width of 1)
    col_widths = [1] * max_cols
    for row in cleaned_rows:
        for col_idx, cell in enumerate(row):
            if len(cell) > col_widths[col_idx]:
                col_widths[col_idx] = len(cell)

    lines: List[str] = []
    for row_idx, row in enumerate(cleaned_rows):
        formatted_cells = [
            cell.ljust(col_widths[col_idx]) for col_idx, cell in enumerate(row)
        ]
        line = " | ".join(formatted_cells).rstrip()
        lines.append(line)

        # After the first row (header row), add a separator
        if row_idx == 0 and len(cleaned_rows) > 1:
            separators = ["-" * col_widths[col_idx] for col_idx in range(max_cols)]
            lines.append("-+-".join(separators))

    return "\n".join(lines)
