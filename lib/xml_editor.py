"""Apply text-only replacements to document.xml <w:t> elements.

Per CLAUDE.md:
- ONLY modify <w:t> text content within <w:r> runs
- NEVER delete <w:r>, <w:rPr>, or <w:pPr> elements
- NEVER add new <w:r> elements
- For multi-run paragraphs: put all new text in first <w:r>'s <w:t>, empty remaining
- Preserve xml:space="preserve" attributes
"""

from lxml import etree

from lib.ai_tailor import EditInstruction
from lib.docx_parser import ParsedResume

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"


def apply_edits(resume: ParsedResume, edits: list[EditInstruction]) -> None:
    """Apply text edits to the parsed resume's XML elements in-place.

    Each edit replaces text in the corresponding ContentBlock's <w:t> elements:
    - First <w:t> gets the full new_text with xml:space="preserve"
    - Remaining <w:t> elements are emptied (text set to "")
    This preserves all run-level formatting (<w:rPr>).

    After calling this, save the modified XML via docx_packer.
    """
    # Build lookup of editable blocks by index
    block_by_index = {b.index: b for b in resume.editable_blocks}

    for edit in edits:
        block = block_by_index.get(edit.block_index)
        if block is None:
            continue

        t_elems = block.text_elements
        if not t_elems:
            continue

        # Put all new text in the first <w:t>, empty the rest
        t_elems[0].text = edit.new_text
        t_elems[0].set(f"{{{W_NS}}}space", "preserve")

        for t in t_elems[1:]:
            t.text = ""


def save_xml(resume: ParsedResume) -> None:
    """Write the modified XML tree back to the document.xml file."""
    # Parse the file to get the tree, then find root
    tree = etree.parse(resume.doc_xml_path)
    # Actually, the elements we modified are in-memory via lxml refs from parse_docx.
    # We need to serialize the tree that contains those elements.
    # The text_elements in ContentBlock are live lxml element references,
    # so their parent tree is already modified. We just need to find and save it.

    # Get the root from any text element's tree
    if resume.editable_blocks and resume.editable_blocks[0].text_elements:
        root = resume.editable_blocks[0].text_elements[0].getroottree()
        root.write(
            resume.doc_xml_path,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )
    elif resume.all_blocks and resume.all_blocks[0].text_elements:
        root = resume.all_blocks[0].text_elements[0].getroottree()
        root.write(
            resume.doc_xml_path,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )
