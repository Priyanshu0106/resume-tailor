import json
import os
import re
from docx import Document
from docx.shared import Pt
from openai import OpenAI

def _get_client():
    api_key = os.environ.get("OPENROUTER_API_KEY")
    if not api_key:
        raise ValueError("OPENROUTER_API_KEY environment variable is not set.")
    return OpenAI(base_url="https://openrouter.ai/api/v1", api_key=api_key)


def _para_text(para):
    return "".join(run.text for run in para.runs)


def _set_para_text(para, new_text: str):
    """Replace paragraph text while preserving all run-level formatting.
    Puts new text in first run, blanks out the rest."""
    if not para.runs:
        return
    para.runs[0].text = new_text
    for run in para.runs[1:]:
        run.text = ""


def _is_modifiable(para) -> bool:
    """Identify paragraphs safe to reword: bullet points and short descriptive lines.
    Skip: empty lines, single words (section headers), contact info lines."""
    text = _para_text(para).strip()
    if not text:
        return False
    # Skip very short lines (likely headers/names)
    if len(text.split()) <= 3:
        return False
    # Skip lines that look like contact info (email, phone, urls)
    if re.search(r"[@|•|linkedin|github|http|\+?\d[\d\s\-\(\)]{7,}]", text, re.I):
        return False
    # Skip lines that are ALL CAPS (section headers)
    if text.isupper():
        return False
    return True


def extract_modifiable(doc: Document) -> dict[int, str]:
    """Return {paragraph_index: text} for paragraphs that can be reworded."""
    result = {}
    for i, para in enumerate(doc.paragraphs):
        if _is_modifiable(para):
            result[i] = _para_text(para).strip()
    return result


def tailor_with_claude(paragraphs: dict[int, str], jd: str) -> dict[int, str]:
    """Send resume paragraphs + JD to Claude, get back reworded versions."""
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
    # Strip markdown code fences if present
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)

    data = json.loads(raw)
    return {int(k): v for k, v in data.items()}


def apply_tailoring(doc: Document, changes: dict[int, str]) -> Document:
    """Apply Claude's rewrites back to the document."""
    for idx, new_text in changes.items():
        if idx < len(doc.paragraphs):
            _set_para_text(doc.paragraphs[idx], new_text)
    return doc


def tailor_resume_docx(input_path: str, output_path: str, jd: str):
    doc = Document(input_path)
    modifiable = extract_modifiable(doc)
    if not modifiable:
        doc.save(output_path)
        return
    changes = tailor_with_claude(modifiable, jd)
    apply_tailoring(doc, changes)
    doc.save(output_path)


def tailor_resume_pdf(input_path: str, output_path: str, jd: str):
    import pdfplumber
    lines = []
    with pdfplumber.open(input_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for line in text.splitlines():
                lines.append(line)

    # Build index map of modifiable lines (same rules as docx)
    modifiable = {}
    for i, line in enumerate(lines):
        text = line.strip()
        if not text:
            continue
        if len(text.split()) <= 3:
            continue
        if re.search(r"[@|•|linkedin|github|http|\+?\d[\d\s\-\(\)]{7,}]", text, re.I):
            continue
        if text.isupper():
            continue
        modifiable[i] = text

    if modifiable:
        changes = tailor_with_claude(modifiable, jd)
        for idx, new_text in changes.items():
            lines[idx] = new_text

    # Write output as a plain .docx
    doc = Document()
    for line in lines:
        doc.add_paragraph(line)
    doc.save(output_path)


def tailor_resume(input_path: str, output_path: str, jd: str):
    if input_path.lower().endswith(".pdf"):
        tailor_resume_pdf(input_path, output_path, jd)
    else:
        tailor_resume_docx(input_path, output_path, jd)
