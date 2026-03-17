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

# Common OOXML namespaces to preserve in output
_OOXML_NAMESPACES = {
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


def _get_client():
    api_key = os.environ.get("OPENROUTER_API_KEY")
    if not api_key:
        raise ValueError("OPENROUTER_API_KEY environment variable is not set.")
    return OpenAI(base_url="https://openrouter.ai/api/v1", api_key=api_key)


# ---------------------------------------------------------------------------
# Content detection: determine what is editable vs locked
# ---------------------------------------------------------------------------

def _is_modifiable_text(text: str) -> bool:
    """Identify text safe to reword: bullet points and descriptive lines.

    LOCKED (return False):
      - Empty lines, very short lines (headers/names)
      - Contact info (email, phone, URLs, LinkedIn, GitHub)
      - ALL CAPS lines (section headers)
      - Date patterns like MM/YYYY, MM/YY (company/date lines)
      - Education rows (institution | degree | percentage | date)
      - Certificate/achievement titles
      - Pipe-separated lines with very short segments (skills line handled separately)
    """
    text = text.strip()
    if not text:
        return False
    # Skip very short lines (likely headers/names/titles)
    if len(text.split()) <= 3:
        return False
    # Skip contact info
    if re.search(r"@|linkedin\.com|github\.com|https?://|\+?\d[\d\s\-\(\)]{7,}", text, re.I):
        return False
    # Skip ALL CAPS lines (section headers)
    if text.isupper():
        return False
    # Skip lines that are primarily date patterns (company lines like "Abbott India ... 04/2024 – 06/2024")
    if re.search(r"\d{2}/\d{2,4}\s*[–\-]\s*(\d{2}/\d{2,4}|Present)", text, re.I):
        return False
    # Skip education-style rows (contain percentage patterns like 79.68%)
    if re.search(r"\d{2,3}\.\d+%", text):
        return False
    return True


def _is_skills_line(text: str) -> bool:
    """Detect pipe-separated skills lines like 'SAP FICO | Excel | PowerPoint | ...'"""
    return text.count("|") >= 3


# ---------------------------------------------------------------------------
# XML text extraction and replacement (DOCX)
# ---------------------------------------------------------------------------

def _extract_para_text(para_elem) -> str:
    """Extract full text from a w:p element by reading all w:t elements."""
    texts = []
    for t in para_elem.iter(f"{{{_W_NS}}}t"):
        if t.text:
            texts.append(t.text)
    return "".join(texts)


def _set_para_text_xml(para_elem, new_text: str):
    """Replace text in a w:p element using Option B from CLAUDE.md:
    Put all new text in the FIRST w:r's w:t, empty remaining w:t elements.
    This preserves paragraph-level and run-level formatting tags."""
    t_elems = list(para_elem.iter(f"{{{_W_NS}}}t"))
    if not t_elems:
        return
    t_elems[0].text = new_text
    t_elems[0].set(f"{{{_W_NS}}}space", "preserve")
    for t in t_elems[1:]:
        t.text = ""


# ---------------------------------------------------------------------------
# AI tailoring with structured prompt (per CLAUDE.md spec)
# ---------------------------------------------------------------------------

_TAILOR_SYSTEM_PROMPT = """You are a resume tailoring expert. You will receive:
1. A list of resume content blocks (bullets, skills, etc.) with their index numbers
2. A job description

Your job: produce a JSON object mapping block indices to tailored replacement text.

HARD RULES — VIOLATING ANY MAKES THE OUTPUT USELESS:

FORMAT PRESERVATION:
- Each new_text MUST be approximately the same length as original_text (±15% character count)
- If original starts with an action verb ("Automated", "Streamlined", "Addressed"), new_text must also start with a strong action verb
- Preserve ALL numbers/metrics in the original. You may adjust context around them but keep quantified achievements intact
- Do NOT add line breaks, tabs, or special characters not in the original

CONTENT RULES:
- Only modify bullet points and skills lines. Never touch company names, dates, degrees, percentages, or section headers
- Maximum 60% of lines should change. Keep 40%+ unchanged to maintain authenticity
- Changes should be SUBTLE REWORDING, not wholesale rewriting. The resume should still sound like the same person
- Prioritize: (a) adding JD keywords naturally into existing bullets, (b) reordering skills to front-load relevant ones, (c) adjusting verb choices to match JD language
- Never fabricate experience, tools, or achievements not implied by the original content
- Never remove a line entirely — only reword it

SKILLS LINE (if present — identified by pipe | separators):
- Reorder skills so the most JD-relevant ones appear first
- You may add 1-2 skills from the JD IF they are reasonably implied by the existing experience
- Do NOT add skills that have zero basis in the resume content
- Keep the pipe | separator format

OUTPUT FORMAT:
Return ONLY a JSON object mapping each changed index to its new text. Example:
{"3": "tailored line here", "7": "another tailored line"}

If a block should NOT be changed, do NOT include it in the output.
Return ONLY valid JSON, no markdown fences, no explanation."""


def tailor_with_ai(paragraphs: dict[int, str], jd: str) -> dict[int, str]:
    """Send resume paragraphs + JD to AI, get back reworded versions."""
    para_list = "\n".join(
        f'[{idx}] {text}' for idx, text in paragraphs.items()
    )

    user_prompt = f"""JOB DESCRIPTION:
{jd}

RESUME LINES TO TAILOR (format: [index] text):
{para_list}"""

    message = _get_client().chat.completions.create(
        model="google/gemini-2.0-flash-001",
        max_tokens=4096,
        messages=[
            {"role": "system", "content": _TAILOR_SYSTEM_PROMPT},
            {"role": "user", "content": user_prompt},
        ],
    )

    raw = message.choices[0].message.content.strip()
    # Strip markdown code fences if present
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)

    data = json.loads(raw)
    return {int(k): v for k, v in data.items()}


# ---------------------------------------------------------------------------
# DOCX tailoring — direct XML/ZIP editing (never use python-docx)
# ---------------------------------------------------------------------------

def tailor_resume_docx(input_path: str, output_path: str, jd: str):
    """Tailor a DOCX resume by directly editing XML inside the ZIP.
    Per CLAUDE.md: ONLY edit word/document.xml, never touch styles/numbering/media."""

    # Step 1: Unpack the .docx ZIP
    with zipfile.ZipFile(input_path, "r") as zin:
        doc_xml = zin.read("word/document.xml")
        all_parts = {name: zin.read(name) for name in zin.namelist()}

    # Register all OOXML namespaces to preserve them in output
    for prefix, uri in _OOXML_NAMESPACES.items():
        ET.register_namespace(prefix, uri)

    root = ET.fromstring(doc_xml)
    body = root.find(f"{{{_W_NS}}}body")

    # Step 2: Extract paragraphs and identify modifiable ones
    paragraphs = list(body.iter(f"{{{_W_NS}}}p"))
    modifiable = {}
    for i, para in enumerate(paragraphs):
        text = _extract_para_text(para).strip()
        if _is_modifiable_text(text):
            modifiable[i] = text

    if not modifiable:
        shutil.copy(input_path, output_path)
        return

    # Step 3: Get tailored text from AI
    changes = tailor_with_ai(modifiable, jd)

    # Step 4: Apply edits — ONLY modify <w:t> text content
    for idx, new_text in changes.items():
        if idx < len(paragraphs):
            _set_para_text_xml(paragraphs[idx], new_text)

    # Step 5: Repack — write modified XML back into the ZIP
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


# ---------------------------------------------------------------------------
# PDF tailoring — PyMuPDF in-place text replacement
# ---------------------------------------------------------------------------

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

        # Second pass: insert new text at same positions with same font/color
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
                # If text doesn't fit, try with smaller font
                try:
                    tw = fitz.TextWriter(page.rect)
                    tw.fill_textbox(rect, new_text, fontsize=fontsize * 0.85, font=font, align=0)
                    tw.write_text(page, color=(r, g, b))
                except ValueError:
                    pass  # skip this block if it still doesn't fit

        doc.save(output_path)
    finally:
        doc.close()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def tailor_resume(input_path: str, output_path: str, jd: str):
    if input_path.lower().endswith(".pdf"):
        tailor_resume_pdf(input_path, output_path, jd)
    else:
        tailor_resume_docx(input_path, output_path, jd)
