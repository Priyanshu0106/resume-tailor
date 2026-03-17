import json
import os
import re
import shutil
import zipfile
from io import BytesIO
from xml.etree import ElementTree as ET
from openai import OpenAI


# Word XML namespace
_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ET.register_namespace("w", _W_NS)


def _get_client():
    api_key = os.environ.get("OPENROUTER_API_KEY")
    if not api_key:
        raise ValueError("OPENROUTER_API_KEY environment variable is not set.")
    return OpenAI(base_url="https://openrouter.ai/api/v1", api_key=api_key)


def _is_modifiable_text(text: str) -> bool:
    """Identify text safe to reword: bullet points and descriptive lines.
    Skip: empty lines, short headers, contact info, ALL CAPS headers."""
    text = text.strip()
    if not text:
        return False
    if len(text.split()) <= 3:
        return False
    if re.search(r"@|linkedin\.com|github\.com|https?://|\+?\d[\d\s\-\(\)]{7,}", text, re.I):
        return False
    if text.isupper():
        return False
    return True


def _extract_para_text(para_elem) -> str:
    """Extract full text from a w:p element by reading all w:t elements."""
    texts = []
    for t in para_elem.iter(f"{{{_W_NS}}}t"):
        if t.text:
            texts.append(t.text)
    return "".join(texts)


def _set_para_text_xml(para_elem, new_text: str):
    """Replace text in a w:p element. Puts all text in the first w:r/w:t,
    blanks remaining w:t elements. Preserves all run formatting."""
    t_elems = list(para_elem.iter(f"{{{_W_NS}}}t"))
    if not t_elems:
        return
    t_elems[0].text = new_text
    t_elems[0].set(f"{{{_W_NS}}}space", "preserve")
    for t in t_elems[1:]:
        t.text = ""


def tailor_with_ai(paragraphs: dict[int, str], jd: str) -> dict[int, str]:
    """Send resume paragraphs + JD to AI, get back reworded versions."""
    para_list = "\n".join(
        f'[{idx}] {text}' for idx, text in paragraphs.items()
    )

    prompt = f"""You are a professional resume writer. Your job is to tailor resume bullet points and descriptions to better match a job description — WITHOUT changing the format, structure, or length significantly.

RULES:
- Rewrite each line to naturally incorporate relevant keywords/skills from the JD
- Keep roughly the same length and sentence structure
- Do NOT add new bullet points or remove existing ones
- Do NOT change section headers, names, dates, or contact info (they won't be sent)
- Keep quantifiable achievements (numbers, percentages) intact
- Sound natural and professional, not keyword-stuffed
- If a line is already a good match, keep it nearly the same

JOB DESCRIPTION:
{jd}

RESUME LINES TO TAILOR (format: [index] text):
{para_list}

Return a JSON object mapping each index to its tailored version. Example:
{{"0": "tailored line here", "5": "another tailored line"}}

Return ONLY valid JSON, no explanation."""

    message = _get_client().chat.completions.create(
        model="google/gemini-2.0-flash-001",
        max_tokens=4096,
        messages=[{"role": "user", "content": prompt}],
    )

    raw = message.choices[0].message.content.strip()
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)

    data = json.loads(raw)
    return {int(k): v for k, v in data.items()}


def tailor_resume_docx(input_path: str, output_path: str, jd: str):
    """Tailor a DOCX resume by directly editing XML inside the ZIP.
    This preserves all formatting, styles, and features that python-docx would strip."""

    # Read the docx as a ZIP and parse word/document.xml
    with zipfile.ZipFile(input_path, "r") as zin:
        doc_xml = zin.read("word/document.xml")
        all_parts = {name: zin.read(name) for name in zin.namelist()}

    # Parse the document XML, preserving all namespaces
    # Register common OOXML namespaces to avoid ns0/ns1 prefixes
    namespaces = {
        "wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
        "cx": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        "o": "urn:schemas-microsoft-com:office:office",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        "v": "urn:schemas-microsoft-com:vml",
        "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        "w10": "urn:schemas-microsoft-com:office:word",
        "w": _W_NS,
        "wne": "http://schemas.microsoft.com/office/word/2006/wordml",
        "sl": "http://schemas.openxmlformats.org/schemaLibrary/2006/main",
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
        "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        "lc": "http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas",
        "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
        "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
        "w16se": "http://schemas.microsoft.com/office/word/2015/wordml/symex",
        "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
        "wpi": "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
        "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    }
    for prefix, uri in namespaces.items():
        ET.register_namespace(prefix, uri)

    root = ET.fromstring(doc_xml)
    body = root.find(f"{{{_W_NS}}}body")

    # Extract paragraphs and identify modifiable ones
    paragraphs = list(body.iter(f"{{{_W_NS}}}p"))
    modifiable = {}
    for i, para in enumerate(paragraphs):
        text = _extract_para_text(para).strip()
        if _is_modifiable_text(text):
            modifiable[i] = text

    if not modifiable:
        shutil.copy(input_path, output_path)
        return

    # Get tailored text from AI
    changes = tailor_with_ai(modifiable, jd)

    # Apply changes directly to XML
    for idx, new_text in changes.items():
        if idx < len(paragraphs):
            _set_para_text_xml(paragraphs[idx], new_text)

    # Write modified XML back into the ZIP
    modified_xml = ET.tostring(root, encoding="UTF-8", xml_declaration=True)

    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in all_parts.items():
            if name == "word/document.xml":
                zout.writestr(name, modified_xml)
            else:
                zout.writestr(name, data)

    with open(output_path, "wb") as f:
        f.write(buf.getvalue())


def tailor_resume_pdf(input_path: str, output_path: str, jd: str):
    """Tailor a PDF resume while preserving the original format.
    Uses PyMuPDF to do in-place text replacement on the PDF."""
    import fitz

    doc = fitz.open(input_path)

    try:
        # Step 1: Extract text blocks with position and font metadata
        blocks_info = []
        for page_idx in range(len(doc)):
            page = doc[page_idx]
            page_dict = page.get_text("dict")
            for block in page_dict["blocks"]:
                if block["type"] != 0:  # skip image blocks
                    continue
                full_text = ""
                first_span = None
                for line in block["lines"]:
                    for span in line["spans"]:
                        if first_span is None:
                            first_span = span
                        full_text += span["text"]
                    full_text += " "
                full_text = full_text.strip()

                if not full_text or first_span is None:
                    continue

                blocks_info.append({
                    "page_idx": page_idx,
                    "text": full_text,
                    "bbox": block["bbox"],
                    "fontsize": first_span["size"],
                    "color": first_span["color"],
                })

        # Step 2: Identify modifiable blocks
        modifiable = {}
        for i, bi in enumerate(blocks_info):
            if _is_modifiable_text(bi["text"]):
                modifiable[i] = bi["text"]

        if not modifiable:
            doc.close()
            shutil.copy(input_path, output_path)
            return

        # Step 3: Get tailored text from AI
        changes = tailor_with_ai(modifiable, jd)

        # Step 4: Replace text in PDF
        # First pass: add redaction annotations to clear old text
        for idx in changes:
            bi = blocks_info[idx]
            page = doc[bi["page_idx"]]
            rect = fitz.Rect(bi["bbox"])
            page.add_redact_annot(rect, fill=(1, 1, 1))  # white fill

        # Apply all redactions
        for page in doc:
            page.apply_redactions()

        # Second pass: insert new text at same positions
        for idx, new_text in changes.items():
            bi = blocks_info[idx]
            page = doc[bi["page_idx"]]
            rect = fitz.Rect(bi["bbox"])
            fontsize = bi["fontsize"]
            color_int = bi["color"]

            # Skip tiny/invalid rectangles
            if rect.width < 5 or rect.height < 5:
                continue

            # Convert color int to RGB tuple
            r = ((color_int >> 16) & 0xFF) / 255.0
            g = ((color_int >> 8) & 0xFF) / 255.0
            b = (color_int & 0xFF) / 255.0

            try:
                tw = fitz.TextWriter(page.rect)
                font = fitz.Font("helv")
                tw.fill_textbox(rect, new_text, fontsize=fontsize, font=font, align=0)
                tw.write_text(page, color=(r, g, b))
            except ValueError:
                # If text doesn't fit in the rectangle, try with smaller font
                try:
                    tw = fitz.TextWriter(page.rect)
                    tw.fill_textbox(rect, new_text, fontsize=fontsize * 0.85, font=font, align=0)
                    tw.write_text(page, color=(r, g, b))
                except ValueError:
                    pass  # skip this block if it still doesn't fit

        doc.save(output_path)
    finally:
        doc.close()


def tailor_resume(input_path: str, output_path: str, jd: str):
    if input_path.lower().endswith(".pdf"):
        tailor_resume_pdf(input_path, output_path, jd)
    else:
        tailor_resume_docx(input_path, output_path, jd)
