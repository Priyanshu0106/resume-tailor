"""Unpack .docx ZIP, extract editable content blocks with XML references.

A .docx is a ZIP containing XML files. We ONLY parse word/document.xml
to find editable text. All other files (styles, numbering, media, rels)
are left untouched.
"""

import os
import re
import tempfile
import zipfile
from dataclasses import dataclass, field
from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"


@dataclass
class ContentBlock:
    index: int
    section: str  # e.g., "Abbott India Limited", "Skills", "Education"
    block_type: str  # "bullet", "skill_line", "paragraph", "header", "table_row"
    text: str
    is_editable: bool
    text_elements: list = field(default_factory=list)  # refs to <w:t> lxml elements
    run_info: list = field(default_factory=list)  # list of (elem, text, has_bold) tuples


@dataclass
class ParsedResume:
    editable_blocks: list  # ContentBlock objects where is_editable=True
    all_blocks: list  # all ContentBlock objects (for context)
    temp_dir: str  # path to unpacked .docx directory
    doc_xml_path: str  # path to word/document.xml


def _get_para_text(para) -> str:
    """Get concatenated text from all w:t elements in a paragraph."""
    return "".join(t.text or "" for t in para.iter(f"{W}t"))


def _get_text_elements(para) -> list:
    """Get all w:t elements in a paragraph."""
    return list(para.iter(f"{W}t"))


def _get_run_info(para) -> list:
    """Get (element, text, has_bold) for each w:r run in a paragraph."""
    runs = []
    for r in para.iter(f"{W}r"):
        rpr = r.find(f"{W}rPr")
        has_bold = rpr is not None and rpr.find(f"{W}b") is not None
        text = "".join(t.text or "" for t in r.iter(f"{W}t"))
        runs.append((r, text, has_bold))
    return runs


def _is_in_table(para) -> bool:
    """Check if paragraph is inside a w:tbl (table)."""
    parent = para.getparent()
    while parent is not None:
        if parent.tag == f"{W}tbl":
            return True
        parent = parent.getparent()
    return False


def _has_numbering(para) -> bool:
    """Check if paragraph has w:numPr (bullet/list numbering)."""
    ppr = para.find(f"{W}pPr")
    if ppr is not None:
        return ppr.find(f"{W}numPr") is not None
    return False


def _is_section_header(text: str) -> bool:
    """Detect section headers: short text, often styled differently."""
    text = text.strip()
    if not text:
        return False
    if len(text.split()) <= 3 and len(text) < 40:
        return True
    if text.isupper():
        return True
    return False


def _is_company_date_line(text: str) -> bool:
    """Detect lines with company name + date pattern like MM/YYYY – MM/YYYY."""
    return bool(re.search(r"\d{2}/\d{2,4}\s*[–\-]\s*(\d{2}/\d{2,4}|Present)", text, re.I))


def _is_contact_info(text: str) -> bool:
    """Detect contact info lines (email, phone, URL)."""
    return bool(re.search(r"@|linkedin\.com|github\.com|https?://|\+?\d[\d\s\-\(\)]{7,}", text, re.I))


def _is_education_row(text: str) -> bool:
    """Detect education table rows with percentage patterns."""
    return bool(re.search(r"\d{2,3}\.\d+%", text))


def _is_skills_line(text: str) -> bool:
    """Detect pipe-separated skills lines."""
    return text.count("|") >= 3


def _classify_block(para, text: str) -> tuple:
    """Classify a paragraph and return (block_type, is_editable)."""
    text_stripped = text.strip()
    if not text_stripped:
        return ("empty", False)

    # Table rows (education) → locked
    if _is_in_table(para):
        return ("table_row", False)

    # Contact info → locked
    if _is_contact_info(text_stripped):
        return ("contact", False)

    # Company/date lines → locked
    if _is_company_date_line(text_stripped):
        return ("company_line", False)

    # Section headers → locked
    if _is_section_header(text_stripped):
        return ("header", False)

    # Education rows → locked
    if _is_education_row(text_stripped):
        return ("education", False)

    # Skills line → editable
    if _is_skills_line(text_stripped):
        return ("skill_line", True)

    # Bullet points (has numbering) → editable
    if _has_numbering(para):
        return ("bullet", True)

    # Long enough paragraphs (4+ words) → editable (likely bullet without formal numbering)
    if len(text_stripped.split()) >= 4:
        return ("paragraph", True)

    return ("other", False)


def parse_docx(docx_path: str) -> ParsedResume:
    """Unpack a .docx and extract all content blocks with XML references."""
    temp_dir = tempfile.mkdtemp(prefix="resume_tailor_")

    # Extract the ZIP
    with zipfile.ZipFile(docx_path, "r") as z:
        z.extractall(temp_dir)

    doc_xml_path = os.path.join(temp_dir, "word", "document.xml")
    tree = etree.parse(doc_xml_path)
    root = tree.getroot()
    body = root.find(f"{W}body")

    all_blocks = []
    editable_blocks = []
    current_section = "Unknown"
    idx = 0

    for para in body.iter(f"{W}p"):
        text = _get_para_text(para)
        text_stripped = text.strip()

        if not text_stripped:
            continue

        block_type, is_editable = _classify_block(para, text)

        # Track current section from headers and company lines
        if block_type == "header":
            current_section = text_stripped
        elif block_type == "company_line":
            # Extract company name (text before the date)
            match = re.match(r"^(.+?)\s*\d{2}/", text_stripped)
            if match:
                current_section = match.group(1).strip()

        block = ContentBlock(
            index=idx,
            section=current_section,
            block_type=block_type,
            text=text_stripped,
            is_editable=is_editable,
            text_elements=_get_text_elements(para),
            run_info=_get_run_info(para),
        )

        all_blocks.append(block)
        if is_editable:
            editable_blocks.append(block)
        idx += 1

    return ParsedResume(
        editable_blocks=editable_blocks,
        all_blocks=all_blocks,
        temp_dir=temp_dir,
        doc_xml_path=doc_xml_path,
    )
