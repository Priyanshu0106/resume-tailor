"""Call OpenRouter API (Gemini) to generate tailored text replacements for resume blocks."""

import json
import os
import re
from dataclasses import dataclass
from pathlib import Path

from openai import OpenAI

from lib.docx_parser import ContentBlock, ParsedResume
from lib.jd_parser import JDSignals

PROMPTS_DIR = Path(__file__).parent.parent / "prompts"


@dataclass
class EditInstruction:
    block_index: int
    original_text: str
    new_text: str
    reason: str


def _get_client():
    api_key = os.environ.get("OPENROUTER_API_KEY")
    if not api_key:
        raise ValueError("OPENROUTER_API_KEY environment variable is not set.")
    return OpenAI(base_url="https://openrouter.ai/api/v1", api_key=api_key)


def _load_system_prompt() -> str:
    """Load the tailoring system prompt from prompts/tailor_system.txt."""
    prompt_path = PROMPTS_DIR / "tailor_system.txt"
    return prompt_path.read_text(encoding="utf-8")


def _build_user_prompt(resume: ParsedResume, signals: JDSignals) -> str:
    """Build the user message with editable blocks, locked context, and JD signals."""
    editable_lines = []
    for block in resume.editable_blocks:
        editable_lines.append(
            f"[{block.index}] (type: {block.block_type}, section: {block.section}) {block.text}"
        )

    locked_lines = []
    for block in resume.all_blocks:
        if not block.is_editable:
            locked_lines.append(f"[{block.index}] ({block.block_type}) {block.text}")

    jd_section = f"""JOB TITLE: {signals.job_title}
COMPANY: {signals.company}
SENIORITY: {signals.seniority_level}
REQUIRED SKILLS: {', '.join(signals.required_skills)}
PREFERRED SKILLS: {', '.join(signals.preferred_skills)}
KEY ACTION VERBS: {', '.join(signals.key_action_verbs)}
DOMAIN KEYWORDS: {', '.join(signals.domain_keywords)}
FOCUS AREAS: {', '.join(signals.focus_areas)}
ATS KEYWORDS: {', '.join(signals.ats_keywords)}"""

    return f"""JOB DESCRIPTION SIGNALS:
{jd_section}

EDITABLE RESUME BLOCKS (you may modify these):
{chr(10).join(editable_lines)}

LOCKED RESUME BLOCKS (context only — do NOT modify or include in output):
{chr(10).join(locked_lines)}"""


def _validate_edit(edit: EditInstruction, resume: ParsedResume) -> bool:
    """Validate an edit instruction against constraints."""
    # Check block_index refers to an editable block
    editable_indices = {b.index for b in resume.editable_blocks}
    if edit.block_index not in editable_indices:
        return False

    # Check original_text matches
    matching = [b for b in resume.editable_blocks if b.index == edit.block_index]
    if not matching:
        return False
    actual_text = matching[0].text
    if actual_text != edit.original_text:
        return False

    # Check ±20% character count
    orig_len = len(edit.original_text)
    new_len = len(edit.new_text)
    if orig_len > 0:
        ratio = new_len / orig_len
        if ratio < 0.8 or ratio > 1.2:
            return False

    return True


def generate_edits(resume: ParsedResume, signals: JDSignals) -> list[EditInstruction]:
    """Generate validated edit instructions by calling OpenRouter API.

    Returns a list of EditInstruction objects that have passed validation.
    """
    system_prompt = _load_system_prompt()
    user_prompt = _build_user_prompt(resume, signals)

    message = _get_client().chat.completions.create(
        model="google/gemini-2.0-flash-001",
        max_tokens=4096,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
    )

    raw = message.choices[0].message.content.strip()
    # Strip markdown fences if present
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)

    data = json.loads(raw)

    # Build a lookup for original text by index
    block_by_index = {b.index: b for b in resume.editable_blocks}

    edits = []
    for idx_str, val in data.items():
        idx = int(idx_str)
        if idx not in block_by_index:
            continue

        new_text = val["new_text"] if isinstance(val, dict) else val
        reason = val.get("reason", "") if isinstance(val, dict) else ""

        edit = EditInstruction(
            block_index=idx,
            original_text=block_by_index[idx].text,
            new_text=new_text,
            reason=reason,
        )

        if _validate_edit(edit, resume):
            edits.append(edit)

    return edits
