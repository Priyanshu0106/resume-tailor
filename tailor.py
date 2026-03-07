import json
import os
import re
from docx import Document
from openai import OpenAI

client = OpenAI(
    base_url="https://openrouter.ai/api/v1",
    api_key=os.environ.get("OPENROUTER_API_KEY"),
)


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

    message = client.chat.completions.create(
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


def tailor_resume(input_path: str, output_path: str, jd: str):
    doc = Document(input_path)
    modifiable = extract_modifiable(doc)
    if not modifiable:
        doc.save(output_path)
        return
    changes = tailor_with_claude(modifiable, jd)
    apply_tailoring(doc, changes)
    doc.save(output_path)
