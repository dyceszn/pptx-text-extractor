# PowerPoint Text Extraction Tool (`extract_pptx_text.py`)

A standalone, 100% local command-line tool to extract all text from PowerPoint (`.pptx`) presentations into cleanly formatted `.txt` or structured `.json` files.

**Privacy Guarantee**: All processing happens entirely on your machine. Files and text content never leave your computer and are never sent to external APIs or cloud services.

---

## Features

- **Comprehensive Element Extraction**:
  - Titles, subtitles, and standard text boxes
  - Paragraphs and bullet points (preserves line breaks and structure)
  - Tables with aligned columns and separator formatting
  - Grouped shapes (recursively traverses nested groups)
  - Charts (extracts chart titles, series names, and categories)
  - SmartArt and DrawingML XML text
  - Speaker notes / slide notes (toggleable with `--include-notes`)
  - Optional local OCR for embedded images (via Tesseract)
- **Natural Reading Order**:
  - Automatically sorts slide shapes from top-to-bottom and left-to-right using spatial coordinates with tolerance bands.
  - Prioritizes slide titles at the top of each slide.
- **Robust Error Handling**:
  - Friendly error messages for missing files, empty presentations, corrupted packages, and password-protected/encrypted presentations.
- **Multiple Output Formats**:
  - Human-readable `.txt` with clear slide dividers (`===== SLIDE X =====`).
  - Structured `.json` with bounding boxes, element types, and table row/column data.
- **Batch Processing**:
  - Pass a directory to process multiple PowerPoint presentations in one run (supports recursive search).

---

## Installation

### 1. Requirements
- Python 3.9 or higher

### 2. Setup Virtual Environment

```bash
# Clone or navigate to the repository
cd powerpoint_extraction_tool

# Create and activate virtual environment
python3 -m venv .venv
source .venv/bin/activate

# Install dependencies
pip install -r requirements.txt
```

### 3. (Optional) Local OCR Setup
If you want to extract text from images embedded inside slides using `--ocr`, install Tesseract:

- **macOS (Homebrew)**:
  ```bash
  brew install tesseract
  ```
- **Ubuntu / Debian**:
  ```bash
  sudo apt-get install tesseract-ocr
  ```
- **Windows**:
  Download and install from the [Tesseract GitHub repository](https://github.com/tesseract-ocr/tesseract) and ensure `tesseract` is added to your system `PATH`.

*Note: OCR is disabled by default. If `--ocr` is used without Tesseract installed, the tool gracefully notifies you and proceeds with standard text extraction.*

---

## Usage & CLI Options

```bash
python extract_pptx_text.py <input.pptx|directory> [options]
```

### Options

| Option | Description |
|---|---|
| `<input>` | Path to a `.pptx` file or a directory of `.pptx` files. |
| `-o, --output` | Custom output `.txt` or `.json` file path, or output directory in batch mode. If omitted, saves to `<filename>.txt` beside the input file. |
| `--json` | Output structured JSON with bounding boxes, slide metadata, and element breakdowns. |
| `--include-notes` | Include speaker notes in the extracted output. |
| `--ocr` | Enable local Tesseract OCR on slide images. |
| `--recursive` | Recursively search subdirectories when `<input>` is a directory. |
| `-v, --verbose` | Enable debug logs. |

---

## Examples

### 1. Basic Text Extraction
Extracts all text into `presentation.txt`:
```bash
python extract_pptx_text.py presentation.pptx
```

### 2. Custom Output File
```bash
python extract_pptx_text.py presentation.pptx -o /path/to/extracted_content.txt
```

### 3. Include Speaker Notes
```bash
python extract_pptx_text.py presentation.pptx --include-notes
```

Sample output:
```text
===== SLIDE 1 =====

AI and Modern Engineering

Local Extraction & Processing Demo

===== SLIDE 2 =====

Architecture Overview
High performance local pipeline for presentation analysis.

--- Notes ---
Speaker Note: Emphasize that international character sets are preserved intact without mangling.
```

### 4. Structured JSON Output
```bash
python extract_pptx_text.py presentation.pptx --json --include-notes -o output.json
```

Sample JSON structure:
```json
{
  "presentation": "/path/to/presentation.pptx",
  "slide_count": 5,
  "slides": [
    {
      "slide_number": 1,
      "text": "...",
      "elements": [
        {
          "type": "title",
          "text": "AI and Modern Engineering",
          "bbox": {
            "top": 2130425,
            "left": 685800,
            "width": 7772400,
            "height": 1470025
          }
        }
      ],
      "notes": "Speaker notes here..."
    }
  ]
}
```

### 5. Table Formatting
Tables are automatically parsed and formatted as aligned tables:
```text
===== SLIDE 3 =====

Benchmark Results

Metric  | Cloud Service   | Local Tool
--------+-----------------+---------------
Privacy | External server | 100% On-Device
Latency | ~1200ms         | <50ms
```

### 6. Batch Processing a Directory
Process all `.pptx` presentations in a folder and write the output files into a target directory:
```bash
python extract_pptx_text.py ./presentations/ -o ./extracted_texts/
```

With recursive search:
```bash
python extract_pptx_text.py ./presentations/ -o ./extracted_texts/ --recursive
```

---

## Running Tests

Run the automated test suite with `pytest`:

```bash
# Ensure virtual environment is activated
source .venv/bin/activate

pytest -v tests/
```

The test suite covers:
- Standard text, titles, subtitles, and paragraphs
- Spatial reading order sorting
- Tables formatting and extraction
- Grouped shapes and nested child shapes
- Speaker notes extraction
- Multilingual and Unicode character preservation
- Empty slide detection (`[Empty Slide]`)
- Chart metadata and title extraction
- Corrupted, missing, and password-protected file detection
- Image extraction and OCR handling
- CLI flags and batch directory execution

---

## Limitations

- **Legacy `.ppt` Files**: Microsoft PowerPoint 97-2003 `.ppt` binary files are not supported. Only modern OpenXML `.pptx` packages are supported.
- **Password Protection**: Encrypted PowerPoint files cannot be inspected without decryption credentials; the tool detects this condition and warns the user clearly.
- **Embedded Audio / Video**: Only textual descriptions or placeholder labels are extracted.
- **Complex Diagrams**: Text inside vector SmartArt or custom DrawingML tags is extracted, but complex hierarchical tree shapes are represented as top-down text blocks.
