"""Resume tailoring tool — CLI entry point and web API backend.

CLI usage:
    python tailor.py --resume resume.docx --jd jd.txt --output tailored.docx
    python tailor.py --resume resume.docx --jd-url https://example.com/job --dry-run

Web app (app.py) imports tailor_resume() which uses OpenRouter/Gemini.
CLI mode uses the modular lib/ pipeline with OpenRouter/Gemini.
"""

import argparse
import json
import os
import re
import shutil
import sys
import zipfile
from io import BytesIO
from pathlib import Path
from xml.etree import ElementTree as ET
from openai import OpenAI


# ===================================================================
# WEB APP BACKEND (used by app.py via `from tailor import tailor_resume`)
# Uses OpenRouter + Gemini — kept for Render deployment compatibility
# ===================================================================

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


def _is_modifiable_text(text: str) -> bool:
    text = text.strip()
    if not text:
        return False
    if len(text.split()) <= 3:
        return False
    if re.search(r"@|linkedin\.com|github\.com|https?://|\+?\d[\d\s\-\(\)]{7,}", text, re.I):
        return False
    if text.isupper():
        return False
    if re.search(r"\d{2}/\d{2,4}\s*[–\-]\s*(\d{2}/\d{2,4}|Present)", text, re.I):
        return False
    if re.search(r"\d{2,3}\.\d+%", text):
        return False
    return True


def _extract_para_text(para_elem) -> str:
    texts = []
    for t in para_elem.iter(f"{{{_W_NS}}}t"):
        if t.text:
            texts.append(t.text)
    return "".join(texts)


def _set_para_text_xml(para_elem, new_text: str):
    t_elems = list(para_elem.iter(f"{{{_W_NS}}}t"))
    if not t_elems:
        return
    t_elems[0].text = new_text
    t_elems[0].set(f"{{{_W_NS}}}space", "preserve")
    for t in t_elems[1:]:
        t.text = ""


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
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)

    data = json.loads(raw)
    return {int(k): v for k, v in data.items()}


def tailor_resume_docx(input_path: str, output_path: str, jd: str):
    with zipfile.ZipFile(input_path, "r") as zin:
        doc_xml = zin.read("word/document.xml")
        all_parts = {name: zin.read(name) for name in zin.namelist()}

    for prefix, uri in _OOXML_NAMESPACES.items():
        ET.register_namespace(prefix, uri)

    root = ET.fromstring(doc_xml)
    body = root.find(f"{{{_W_NS}}}body")

    paragraphs = list(body.iter(f"{{{_W_NS}}}p"))
    modifiable = {}
    for i, para in enumerate(paragraphs):
        text = _extract_para_text(para).strip()
        if _is_modifiable_text(text):
            modifiable[i] = text

    if not modifiable:
        shutil.copy(input_path, output_path)
        return

    changes = tailor_with_ai(modifiable, jd)

    for idx, new_text in changes.items():
        if idx < len(paragraphs):
            _set_para_text_xml(paragraphs[idx], new_text)

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
    import fitz

    doc = fitz.open(input_path)

    try:
        blocks_info = []
        for page_idx in range(len(doc)):
            page = doc[page_idx]
            page_dict = page.get_text("dict")
            for block in page_dict["blocks"]:
                if block["type"] != 0:
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

        modifiable = {}
        for i, bi in enumerate(blocks_info):
            if _is_modifiable_text(bi["text"]):
                modifiable[i] = bi["text"]

        if not modifiable:
            doc.close()
            shutil.copy(input_path, output_path)
            return

        changes = tailor_with_ai(modifiable, jd)

        for idx in changes:
            bi = blocks_info[idx]
            page = doc[bi["page_idx"]]
            rect = fitz.Rect(bi["bbox"])
            page.add_redact_annot(rect, fill=(1, 1, 1))

        for page in doc:
            page.apply_redactions()

        for idx, new_text in changes.items():
            bi = blocks_info[idx]
            page = doc[bi["page_idx"]]
            rect = fitz.Rect(bi["bbox"])
            fontsize = bi["fontsize"]
            color_int = bi["color"]

            if rect.width < 5 or rect.height < 5:
                continue

            r = ((color_int >> 16) & 0xFF) / 255.0
            g = ((color_int >> 8) & 0xFF) / 255.0
            b = (color_int & 0xFF) / 255.0

            try:
                tw = fitz.TextWriter(page.rect)
                font = fitz.Font("helv")
                tw.fill_textbox(rect, new_text, fontsize=fontsize, font=font, align=0)
                tw.write_text(page, color=(r, g, b))
            except ValueError:
                try:
                    tw = fitz.TextWriter(page.rect)
                    tw.fill_textbox(rect, new_text, fontsize=fontsize * 0.85, font=font, align=0)
                    tw.write_text(page, color=(r, g, b))
                except ValueError:
                    pass

        doc.save(output_path)
    finally:
        doc.close()


def tailor_resume(input_path: str, output_path: str, jd: str):
    """Web app entry point — called by app.py. Uses OpenRouter/Gemini."""
    if input_path.lower().endswith(".pdf"):
        tailor_resume_pdf(input_path, output_path, jd)
    else:
        tailor_resume_docx(input_path, output_path, jd)


# ===================================================================
# CLI MODE — modular pipeline using lib/ modules + OpenRouter/Gemini
# ===================================================================

def _cli_tailor(args):
    """Run the full CLI tailoring pipeline."""
    from lib.docx_parser import parse_docx
    from lib.jd_parser import parse_jd
    from lib.ai_tailor import generate_edits
    from lib.xml_editor import apply_edits, save_xml
    from lib.docx_packer import repack_docx, cleanup_temp

    if not os.environ.get("OPENROUTER_API_KEY"):
        print("Error: OPENROUTER_API_KEY environment variable is not set.", file=sys.stderr)
        sys.exit(1)

    resume_path = args.resume
    if not os.path.isfile(resume_path):
        print(f"Error: Resume file not found: {resume_path}", file=sys.stderr)
        sys.exit(1)

    if not resume_path.lower().endswith(".docx"):
        print("Error: CLI mode only supports .docx files. Use the web app for PDF.", file=sys.stderr)
        sys.exit(1)

    # Load JD text
    jd_text = None
    jd_url = None
    if args.jd:
        jd_path = args.jd
        if not os.path.isfile(jd_path):
            print(f"Error: JD file not found: {jd_path}", file=sys.stderr)
            sys.exit(1)
        jd_text = Path(jd_path).read_text(encoding="utf-8")
    elif args.jd_url:
        jd_url = args.jd_url
    else:
        print("Error: Provide --jd (file path) or --jd-url (URL).", file=sys.stderr)
        sys.exit(1)

    # Determine output path
    if args.output:
        output_path = args.output
    else:
        stem = Path(resume_path).stem
        output_path = str(Path(resume_path).parent / f"{stem}_tailored.docx")

    print(f"Parsing resume: {resume_path}")
    resume = parse_docx(resume_path)
    print(f"  Found {len(resume.editable_blocks)} editable blocks, {len(resume.all_blocks)} total blocks")

    print("Parsing job description...")
    signals = parse_jd(jd_text=jd_text, jd_url=jd_url)
    print(f"  Job: {signals.job_title} at {signals.company}")
    print(f"  Required skills: {', '.join(signals.required_skills[:5])}{'...' if len(signals.required_skills) > 5 else ''}")

    print("Generating tailored edits...")
    edits = generate_edits(resume, signals)
    print(f"  Generated {len(edits)} validated edits")

    if args.dry_run:
        print("\n--- DRY RUN (no file written) ---\n")
        print(f"{'Idx':<5} {'Section':<25} {'Change'}")
        print("-" * 80)
        for edit in edits:
            print(f"{edit.block_index:<5} {'':<25}")
            print(f"  BEFORE: {edit.original_text[:80]}{'...' if len(edit.original_text) > 80 else ''}")
            print(f"  AFTER:  {edit.new_text[:80]}{'...' if len(edit.new_text) > 80 else ''}")
            print(f"  REASON: {edit.reason}")
            print()
        return

    print("Applying edits to XML...")
    apply_edits(resume, edits)
    save_xml(resume)

    print(f"Repacking to: {output_path}")
    repack_docx(resume.temp_dir, output_path)

    cleanup_temp(resume.temp_dir)

    output_size = os.path.getsize(output_path)
    input_size = os.path.getsize(resume_path)
    print(f"Done! Output: {output_path} ({output_size:,} bytes, original was {input_size:,} bytes)")


def main():
    parser = argparse.ArgumentParser(
        description="Tailor a .docx resume to a job description using AI.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""Examples:
  python tailor.py --resume resume.docx --jd job.txt
  python tailor.py --resume resume.docx --jd job.txt --output tailored.docx
  python tailor.py --resume resume.docx --jd-url https://example.com/job --dry-run
""",
    )
    parser.add_argument("--resume", required=True, help="Path to .docx resume file")
    parser.add_argument("--jd", help="Path to job description text file")
    parser.add_argument("--jd-url", help="URL to fetch job description from")
    parser.add_argument("--output", help="Output path (default: <resume>_tailored.docx)")
    parser.add_argument("--dry-run", action="store_true", help="Show changes without writing file")

    args = parser.parse_args()
    _cli_tailor(args)


if __name__ == "__main__":
    # If called with CLI args, run CLI mode; otherwise start web server
    if len(sys.argv) > 1 and sys.argv[1].startswith("--"):
        main()
    else:
        import uvicorn
        port = int(os.environ.get("PORT", 8000))
        uvicorn.run("app:app", host="0.0.0.0", port=port)
