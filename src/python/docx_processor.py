# -------------------------------------------------------------------------
# Program: docx_processor.py
# Description: Extract text sections from a Word document and revise
#              the document based on AI analysis.
# Context: DocumentWriter project - GitHub Models integration
# Author: Greg Tate
# -------------------------------------------------------------------------

"""
Processes Word (.docx) documents: extracts text section outlines for AI
analysis, and revises the document based on AI-generated suggestions.

Usage:
    python docx_processor.py extract <docx_path> <output_json>
    python docx_processor.py revise  <docx_path> <analysis_json> <output_docx>
"""

#region IMPORTS
import sys
import json
import re
from pathlib import Path

from docx import Document
from docx.shared import Pt, RGBColor
#endregion


#region MAIN WORKFLOW
def main() -> None:
    """
    Main entry point - dispatches to extract or revise based on CLI args.
    """
    validate_argument(sys.argv)

    # Determine which command to run
    command = sys.argv[1].lower()

    if command == "extract":
        run_extract(sys.argv[2], sys.argv[3])
    elif command == "insert":
        run_insert(sys.argv[2], sys.argv[3], sys.argv[4])
#endregion


#region EXTRACT FUNCTIONS
def run_extract(
    docx_path: str,
    output_json: str
) -> None:
    """
    Extract text sections from a Word document and write JSON output.

    Args:
        docx_path: Path to the input .docx file
        output_json: Path to write the extracted JSON data
    """
    doc = Document(docx_path)

    # Walk paragraphs and build section list
    sections = build_section(doc)

    # Write the extracted data to JSON
    output = {
        "source_document": str(Path(docx_path).resolve()),
        "sections": sections
    }

    with open(output_json, "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2, ensure_ascii=False)

    print(f"Extracted {len(sections)} sections to {output_json}")

def build_section(
    doc: Document
) -> list:
    """
    Walk document paragraphs and group them into sections by heading.

    Args:
        doc: The loaded Document object

    Returns:
        list: List of section dictionaries with headings and text
    """
    sections = []
    current_section = None

    for paragraph in doc.paragraphs:
        style_name = paragraph.style.name if paragraph.style else ""

        # Detect heading paragraphs to start a new section
        if style_name.startswith("Heading"):
            if current_section is not None:
                sections.append(current_section)

            # Parse heading level from style name
            level = extract_heading_level(style_name)

            current_section = {
                "heading": paragraph.text.strip(),
                "heading_level": level,
                "text_content": [],
                "paragraph_index_start": get_paragraph_index(
                    doc, paragraph
                )
            }
        elif current_section is not None:
            # Accumulate body text for the current section
            if paragraph.text.strip():
                current_section["text_content"].append(
                    paragraph.text.strip()
                )

    # Append the last section if present
    if current_section is not None:
        sections.append(current_section)

    # Handle documents with no headings - treat entire body as one section
    if not sections:
        sections.append(
            create_default_section(doc)
        )

    return sections


def extract_heading_level(style_name: str) -> int:
    """
    Parse the heading level number from a Word style name.

    Args:
        style_name: The paragraph style name (e.g., 'Heading 2')

    Returns:
        int: The heading level, or 0 if not parsable
    """
    match = re.search(r"(\d+)", style_name)
    if match:
        return int(match.group(1))
    return 0


def get_paragraph_index(
    doc: Document,
    target_paragraph
) -> int:
    """
    Find the index of a paragraph within the document body.

    Args:
        doc: The Document object
        target_paragraph: The paragraph to locate

    Returns:
        int: Zero-based index of the paragraph
    """
    for i, para in enumerate(doc.paragraphs):
        if para._element is target_paragraph._element:
            return i
    return -1

def create_default_section(
    doc: Document
) -> dict:
    """
    Create a single section for documents that have no headings.

    Args:
        doc: The Document object

    Returns:
        dict: A section dictionary covering the entire document
    """
    all_text = [
        p.text.strip()
        for p in doc.paragraphs
        if p.text.strip()
    ]

    return {
        "heading": "Document Content",
        "heading_level": 0,
        "text_content": all_text,
        "paragraph_index_start": 0,
    }
#endregion


#region INSERT FUNCTIONS
def run_insert(
    docx_path: str,
    commentary_json: str,
    output_path: str
) -> None:
    """
    Insert AI-generated commentary into a copy of the Word document.

    Args:
        docx_path: Path to the original .docx file
        commentary_json: Path to JSON file containing commentary entries
        output_path: Path for the annotated output .docx file
    """
    source_path = Path(docx_path).resolve(strict=False)
    target_path = Path(output_path).resolve(strict=False)

    # Ensure the source document is never overwritten
    if source_path == target_path:
        raise ValueError(
            "Output document path must be different from input document path."
        )

    doc = Document(docx_path)

    # Load the commentary data
    with open(commentary_json, "r", encoding="utf-8") as f:
        commentary_data = json.load(f)

    # Insert commentary after each matching section heading
    insert_count = insert_commentary(doc, commentary_data)

    # Save the annotated document
    doc.save(output_path)
    print(
        f"Inserted {insert_count} commentary blocks "
        f"into {output_path}"
    )


def insert_commentary(
    doc: Document,
    commentary_data: list
) -> int:
    """
    Insert formatted commentary paragraphs after section headings.

    Args:
        doc: The Document object to modify
        commentary_data: List of dicts with 'heading' and 'commentary' keys

    Returns:
        int: Number of commentary blocks inserted
    """
    insert_count = 0

    for entry in commentary_data:
        heading_text = entry.get("heading", "")
        commentary_text = entry.get("commentary", "")

        # Skip entries with no commentary
        if not commentary_text.strip():
            continue

        # Find the heading paragraph in the document
        target_index = find_heading_paragraph(doc, heading_text)

        if target_index < 0:
            continue

        # Find the end of the section body to insert after it
        insert_after = find_section_end(doc, target_index)

        # Insert the commentary paragraph
        add_commentary_paragraph(doc, insert_after, commentary_text)
        insert_count += 1

    return insert_count


def find_heading_paragraph(
    doc: Document,
    heading_text: str
) -> int:
    """
    Find the index of a paragraph matching the given heading text.

    Args:
        doc: The Document object
        heading_text: Text content of the heading to find

    Returns:
        int: Index of the matching paragraph, or -1 if not found
    """
    for i, para in enumerate(doc.paragraphs):
        style_name = para.style.name if para.style else ""

        if (
            style_name.startswith("Heading")
            and para.text.strip() == heading_text.strip()
        ):
            return i

    return -1


def find_section_end(
    doc: Document,
    heading_index: int
) -> int:
    """
    Find the last paragraph index belonging to a section.

    Scans forward from the heading until the next heading or document end.

    Args:
        doc: The Document object
        heading_index: Index of the section heading paragraph

    Returns:
        int: Index of the last paragraph in the section
    """
    last_index = heading_index

    for i in range(heading_index + 1, len(doc.paragraphs)):
        style_name = doc.paragraphs[i].style.name if doc.paragraphs[i].style else ""

        # Stop at the next heading
        if style_name.startswith("Heading"):
            break

        last_index = i

    return last_index


def add_commentary_paragraph(
    doc: Document,
    after_index: int,
    commentary_text: str
) -> None:
    """
    Insert a visually distinct commentary paragraph after the given index.

    Args:
        doc: The Document object
        after_index: Paragraph index to insert after
        commentary_text: The commentary text to insert
    """
    # Build the new paragraph element
    target_element = doc.paragraphs[after_index]._element
    new_para = doc.add_paragraph()

    # Style the commentary paragraph
    new_para.style = doc.styles["Normal"]
    new_para.paragraph_format.space_before = Pt(6)
    new_para.paragraph_format.space_after = Pt(6)

    # Add a bold label prefix
    label_run = new_para.add_run("[AI Commentary] ")
    label_run.bold = True
    label_run.font.color.rgb = RGBColor(0x00, 0x51, 0x8A)
    label_run.font.size = Pt(10)

    # Add the commentary text
    text_run = new_para.add_run(commentary_text)
    text_run.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
    text_run.font.size = Pt(10)
    text_run.italic = True

    # Move the new paragraph element to the correct position
    new_element = new_para._element
    new_element.getparent().remove(new_element)
    target_element.addnext(new_element)
#endregion


#region VALIDATION
def validate_argument(argv: list) -> None:
    """
    Validate command-line arguments and print usage if incorrect.

    Args:
        argv: The sys.argv list
    """
    if len(argv) < 2:
        show_usage()
        sys.exit(1)

    command = argv[1].lower()

    # Validate extract command requires 2 additional args
    if command == "extract" and len(argv) != 4:
        print("Error: 'extract' requires <docx_path> <output_json>")
        show_usage()
        sys.exit(1)

    # Validate insert command requires 3 additional args
    if command == "insert" and len(argv) != 5:
        print(
            "Error: 'insert' requires "
            "<docx_path> <commentary_json> <output_docx>"
        )
        show_usage()
        sys.exit(1)

    # Validate command name
    if command not in ("extract", "insert"):
        print(f"Error: Unknown command '{command}'")
        show_usage()
        sys.exit(1)


def show_usage() -> None:
    """
    Print usage instructions to stderr.
    """
    print(
        "Usage:\n"
        "  python docx_processor.py extract "
        "<docx_path> <output_json>\n"
        "  python docx_processor.py insert  "
        "<docx_path> <commentary_json> <output_docx>",
        file=sys.stderr,
    )
#endregion


# -------------------------------------------------------------------------
# Script Entry Point
# -------------------------------------------------------------------------

if __name__ == "__main__":
    main()
