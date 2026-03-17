"""Apply text-only replacements to document.xml <w:t> elements.

CRITICAL RULES (from fix prompt):
- Each <w:p> paragraph MUST be edited independently
- NEVER concatenate text from multiple paragraphs into one
- NEVER delete <w:r>, <w:rPr>, <w:pPr>, or <w:numPr> elements
- Only modify the .text property of <w:t> nodes
- For multi-run paragraphs: put new text in FIRST <w:t> of THAT paragraph only,
  empty remaining <w:t> elements ONLY within that same <w:p>
"""

import sys

from lxml import etree

from lib.ai_tailor import EditInstruction
from lib.docx_parser import ParsedResume

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"


def apply_edits(resume: ParsedResume, edits: list[EditInstruction]) -> None:
    """Apply text edits to the parsed resume's XML elements in-place.

    CORRECT approach — edit per paragraph:
    1. Get the specific <w:p> element for THIS bullet
    2. Get only the <w:t> elements WITHIN THIS <w:p>
    3. Put new text in first <w:t> of THIS paragraph only
    4. Empty remaining <w:t> elements ONLY within this same <w:p>
    """
    block_by_index = {b.index: b for b in resume.editable_blocks}

    applied_count = 0
    for edit in edits:
        block = block_by_index.get(edit.block_index)
        if block is None:
            print(f"  WARNING: Block index {edit.block_index} not found in editable blocks, skipping", file=sys.stderr)
            continue

        # Get <w:t> elements ONLY from this block's paragraph
        # text_elements are collected per-paragraph in docx_parser._get_text_elements(para)
        t_elems = block.text_elements
        if not t_elems:
            print(f"  WARNING: Block {edit.block_index} has no <w:t> elements, skipping", file=sys.stderr)
            continue

        # Verify all t_elems belong to the same <w:p> — safety check
        para_of_first = _get_parent_paragraph(t_elems[0])
        if para_of_first is not None:
            for t in t_elems[1:]:
                if _get_parent_paragraph(t) is not para_of_first:
                    print(f"  ERROR: Block {edit.block_index} has <w:t> elements spanning multiple <w:p>! Skipping to prevent paragraph collapse.", file=sys.stderr)
                    continue

        # Put all new text in the first <w:t>, empty the rest WITHIN THIS PARAGRAPH ONLY
        t_elems[0].text = edit.new_text
        t_elems[0].set(f"{{{W_NS}}}space", "preserve")

        for t in t_elems[1:]:
            t.text = ""

        applied_count += 1
        print(f"  Applied edit to block {edit.block_index}: {edit.reason}")

    print(f"  Total edits applied: {applied_count}/{len(edits)}")

    if applied_count == 0 and len(edits) > 0:
        print("  ERROR: Had edits but applied none! Something is wrong.", file=sys.stderr)
        sys.exit(1)


def _get_parent_paragraph(elem):
    """Walk up the tree to find the parent <w:p> element."""
    parent = elem.getparent()
    while parent is not None:
        if parent.tag == f"{W}p":
            return parent
        parent = parent.getparent()
    return None


def save_xml(resume: ParsedResume) -> None:
    """Write the modified XML tree back to the document.xml file.

    The text_elements in ContentBlock are live lxml element references,
    so their parent tree is already modified in-memory. We just serialize it.
    """
    # Get the tree from any element that was parsed
    tree = None
    for block in resume.all_blocks:
        if block.text_elements:
            tree = block.text_elements[0].getroottree()
            break

    if tree is None:
        print("  ERROR: Could not find XML tree to save — no blocks have text elements", file=sys.stderr)
        sys.exit(1)

    # Count <w:p> elements BEFORE saving (for assertion after repack)
    root = tree.getroot()
    para_count = len(list(root.iter(f"{W}p")))
    print(f"  Paragraph count in modified XML: {para_count}")

    tree.write(
        resume.doc_xml_path,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )

    # Verify the saved file by re-parsing and counting paragraphs
    verify_tree = etree.parse(resume.doc_xml_path)
    verify_count = len(list(verify_tree.getroot().iter(f"{W}p")))
    if verify_count != para_count:
        print(f"  ERROR: Paragraph count mismatch! In-memory: {para_count}, on-disk: {verify_count}", file=sys.stderr)
        sys.exit(1)

    print(f"  Saved document.xml ({verify_count} paragraphs verified)")
