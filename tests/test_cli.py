"""CLI invocation and integration tests for extract_pptx_text.py."""

import json
import subprocess
import sys
from pathlib import Path


def run_cli(*args: str) -> subprocess.CompletedProcess:
    """Helper to run extract_pptx_text.py with arguments."""
    cli_path = Path(__file__).parent.parent / "extract_pptx_text.py"
    cmd = [sys.executable, str(cli_path)] + list(args)
    return subprocess.run(cmd, capture_output=True, text=True)


def test_cli_default_output(comprehensive_pptx: Path, tmp_path: Path):
    """Test CLI creates default .txt output alongside input file."""
    # Copy file to isolated directory
    test_file = tmp_path / "test_pres.pptx"
    test_file.write_bytes(comprehensive_pptx.read_bytes())

    res = run_cli(str(test_file))
    assert res.returncode == 0

    expected_output = tmp_path / "test_pres.txt"
    assert expected_output.exists()

    content = expected_output.read_text(encoding="utf-8")
    assert "===== SLIDE 1 =====" in content
    assert "Introduction to Artificial Intelligence" in content


def test_cli_custom_output(comprehensive_pptx: Path, tmp_path: Path):
    """Test CLI writes to custom path specified by -o."""
    custom_out = tmp_path / "custom_dir" / "my_extracted_notes.txt"

    res = run_cli(str(comprehensive_pptx), "-o", str(custom_out))
    assert res.returncode == 0
    assert custom_out.exists()

    content = custom_out.read_text(encoding="utf-8")
    assert "Quarterly Revenue" in content


def test_cli_json_output(comprehensive_pptx: Path, tmp_path: Path):
    """Test CLI produces structured JSON with --json flag."""
    json_out = tmp_path / "result.json"

    res = run_cli(str(comprehensive_pptx), "-o", str(json_out), "--json", "--include-notes")
    assert res.returncode == 0
    assert json_out.exists()

    data = json.loads(json_out.read_text(encoding="utf-8"))
    assert "presentation" in data
    assert data["slide_count"] == 8
    assert len(data["slides"]) == 8

    # Verify slide 5 has notes
    slide_5 = data["slides"][4]
    assert "Remember to emphasize key milestones" in slide_5["notes"]


def test_cli_batch_directory_mode(comprehensive_pptx: Path, tmp_path: Path):
    """Test CLI batch mode when pointing to a directory."""
    input_dir = tmp_path / "batch_input"
    input_dir.mkdir()
    output_dir = tmp_path / "batch_output"
    output_dir.mkdir()

    # Create two presentations in input dir
    (input_dir / "pres1.pptx").write_bytes(comprehensive_pptx.read_bytes())
    (input_dir / "pres2.pptx").write_bytes(comprehensive_pptx.read_bytes())

    res = run_cli(str(input_dir), "-o", str(output_dir))
    assert res.returncode == 0
    assert (output_dir / "pres1.txt").exists()
    assert (output_dir / "pres2.txt").exists()


def test_cli_missing_file_error():
    """Test CLI exit code and error message on missing file."""
    res = run_cli("completely_nonexistent_file.pptx")
    assert res.returncode == 1
    assert "does not exist" in res.stderr.lower()


def test_cli_corrupted_file_error(corrupt_pptx: Path):
    """Test CLI exit code and error message on corrupted file."""
    res = run_cli(str(corrupt_pptx))
    assert res.returncode == 1
    assert "error" in res.stderr.lower()
