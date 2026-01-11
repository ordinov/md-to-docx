#!/usr/bin/env python3
"""
Word (.docx) to Markdown converter - Self-contained version

Usage: python docx2md.py <file.docx>
Output: Creates <file.md> in the same directory

No manual installation required - handles everything automatically.
Works on Windows, Mac, and Linux.
"""

import sys
import os
import subprocess
import platform
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent.resolve()
VENV_DIR = SCRIPT_DIR / ".venv"


def get_python_executable():
    """Get the correct Python executable for the virtual environment."""
    if platform.system() == "Windows":
        return VENV_DIR / "Scripts" / "python.exe"
    else:
        return VENV_DIR / "bin" / "python"


def get_pip_executable():
    """Get the correct pip executable for the virtual environment."""
    if platform.system() == "Windows":
        return VENV_DIR / "Scripts" / "pip.exe"
    else:
        return VENV_DIR / "bin" / "pip"


def setup_environment():
    """Create virtual environment and install dependencies if needed."""
    python_exe = get_python_executable()
    pip_exe = get_pip_executable()

    # Check if venv exists and has python-docx
    if python_exe.exists():
        # Check if python-docx is installed
        result = subprocess.run(
            [str(python_exe), "-c", "import docx"],
            capture_output=True
        )
        if result.returncode == 0:
            return python_exe  # All good

    # Create venv if needed
    if not VENV_DIR.exists():
        print("Setting up environment (first run only)...")
        subprocess.run(
            [sys.executable, "-m", "venv", str(VENV_DIR)],
            check=True
        )

    # Install python-docx
    print("Installing dependencies...")
    subprocess.run(
        [str(pip_exe), "install", "-q", "python-docx"],
        check=True
    )
    print("Setup complete!\n")

    return python_exe


def run_in_venv():
    """Re-run this script inside the virtual environment."""
    python_exe = setup_environment()

    # Re-execute this script with the venv Python
    result = subprocess.run(
        [str(python_exe), __file__] + sys.argv[1:],
        env={**os.environ, "DOCX2MD_IN_VENV": "1"}
    )
    sys.exit(result.returncode)


# === Main conversion logic ===

def docx_to_md(docx_path):
    """Convert docx file to markdown, saving in same location."""
    from docx import Document
    from docx.shared import Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    docx_path = Path(docx_path)

    if not docx_path.exists():
        print(f"Error: File not found: {docx_path}")
        sys.exit(1)

    if docx_path.suffix.lower() != '.docx':
        print(f"Warning: File doesn't have .docx extension: {docx_path}")

    md_path = docx_path.with_suffix('.md')

    doc = Document(docx_path)
    md_lines = []

    def extract_run_text(run):
        """Extract text from a run with formatting."""
        text = run.text
        if not text:
            return ""

        if run.bold and run.italic:
            return f"***{text}***"
        elif run.bold:
            return f"**{text}**"
        elif run.italic:
            return f"*{text}*"
        elif run.underline:
            # Underline often indicates links in converted docs
            return text
        return text

    def process_paragraph(para):
        """Convert a paragraph to markdown."""
        # Check heading style
        style_name = para.style.name if para.style else ""

        # Handle headings
        if style_name.startswith('Heading'):
            try:
                level = int(style_name.replace('Heading', '').strip())
                text = para.text.strip()
                if level == 0 or style_name == 'Title':
                    return f"# {text}"
                else:
                    return f"{'#' * (level + 1)} {text}"
            except ValueError:
                pass

        if style_name == 'Title':
            return f"# {para.text.strip()}"

        # Handle lists
        if style_name == 'List Bullet':
            text = get_formatted_text(para)
            indent = ""
            if para.paragraph_format.left_indent:
                indent_inches = para.paragraph_format.left_indent.inches if hasattr(para.paragraph_format.left_indent, 'inches') else 0
                if indent_inches and indent_inches >= 0.4:
                    indent = "  "
            return f"{indent}- {text}"

        if style_name == 'List Number':
            text = get_formatted_text(para)
            return f"1. {text}"

        # Handle blockquotes (indented paragraphs)
        if para.paragraph_format.left_indent:
            try:
                indent_val = para.paragraph_format.left_indent.inches if hasattr(para.paragraph_format.left_indent, 'inches') else 0
                if indent_val and indent_val >= 0.4 and style_name not in ['List Bullet', 'List Number']:
                    text = get_formatted_text(para)
                    return f"> {text}"
            except (AttributeError, TypeError):
                pass

        # Check for horizontal rule (line of dashes or similar)
        text = para.text.strip()
        if text and all(c in '─-—' for c in text) and len(text) > 10:
            return "---"

        # Regular paragraph
        formatted = get_formatted_text(para)
        return formatted if formatted else ""

    def get_formatted_text(para):
        """Get paragraph text with inline formatting."""
        result = []
        for run in para.runs:
            result.append(extract_run_text(run))
        return "".join(result)

    def process_table(table):
        """Convert a table to markdown."""
        rows = []
        for row in table.rows:
            cells = []
            for cell in row.cells:
                # Get cell text, joining paragraphs
                cell_text = " ".join(p.text.strip() for p in cell.paragraphs if p.text.strip())
                cells.append(cell_text)
            rows.append(cells)

        if not rows:
            return []

        md_table = []
        # Header row
        if rows:
            md_table.append("| " + " | ".join(rows[0]) + " |")
            # Separator
            md_table.append("| " + " | ".join(["---"] * len(rows[0])) + " |")
            # Data rows
            for row in rows[1:]:
                # Pad row if necessary
                while len(row) < len(rows[0]):
                    row.append("")
                md_table.append("| " + " | ".join(row) + " |")

        return md_table

    # Process document elements in order
    for element in doc.element.body:
        if element.tag.endswith('p'):
            # Find corresponding paragraph
            for para in doc.paragraphs:
                if para._element is element:
                    line = process_paragraph(para)
                    if line or (md_lines and md_lines[-1] != ""):
                        md_lines.append(line)
                    break
        elif element.tag.endswith('tbl'):
            # Find corresponding table
            for table in doc.tables:
                if table._tbl is element:
                    # Add blank line before table if needed
                    if md_lines and md_lines[-1] != "":
                        md_lines.append("")
                    md_lines.extend(process_table(table))
                    md_lines.append("")
                    break

    # Clean up multiple blank lines
    cleaned = []
    prev_blank = False
    for line in md_lines:
        is_blank = line == ""
        if is_blank and prev_blank:
            continue
        cleaned.append(line)
        prev_blank = is_blank

    # Remove trailing blank lines
    while cleaned and cleaned[-1] == "":
        cleaned.pop()

    content = "\n".join(cleaned)

    with open(md_path, 'w', encoding='utf-8') as f:
        f.write(content)

    print(f"Converted: {docx_path.name} -> {md_path.name}")
    return md_path


if __name__ == '__main__':
    if len(sys.argv) != 2:
        print("Usage: python docx2md.py <file.docx>")
        print("       Creates <file.md> in the same directory")
        sys.exit(1)

    # If not in venv, set it up and re-run
    if not os.environ.get("DOCX2MD_IN_VENV"):
        run_in_venv()
    else:
        docx_to_md(sys.argv[1])
