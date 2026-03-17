"""Call OpenRouter API (Gemini) to generate tailored text replacements for resume blocks."""

import json
import os
import re
import sys
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
        print("ERROR: OPENROUTER_API_KEY environment variable is not set!", file=sys.stderr)
        sys.exit(1)
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


def _validate_edit(edit: EditInstruction, resume: ParsedResume) -> tuple[bool, str]:
    """Validate an edit instruction. Returns (is_valid, reason_if_invalid)."""
    editable_indices = {b.index for b in resume.editable_blocks}
    if edit.block_index not in editable_indices:
        return False, f"index {edit.block_index} is not editable"

    matching = [b for b in resume.editable_blocks if b.index == edit.block_index]
    if not matching:
        return False, f"index {edit.block_index} not found"

    actual_text = matching[0].text

    # Use fuzzy matching — normalize whitespace before comparing
    normalized_actual = " ".join(actual_text.split())
    normalized_original = " ".join(edit.original_text.split())
    if normalized_actual != normalized_original:
        # Still accept if they're close enough (AI might slightly alter quotes/spaces)
        if normalized_actual not in normalized_original and normalized_original not in normalized_actual:
            return False, f"original_text mismatch for index {edit.block_index}"

    # Check ±20% character count
    orig_len = len(edit.original_text)
    new_len = len(edit.new_text)
    if orig_len > 0:
        ratio = new_len / orig_len
        if ratio < 0.8 or ratio > 1.2:
            print(f"  WARNING: Edit {edit.block_index} length ratio {ratio:.2f} outside ±20% — accepting anyway", file=sys.stderr)
            # Accept anyway but warn — don't silently drop edits

    return True, ""


def generate_edits(resume: ParsedResume, signals: JDSignals) -> list[EditInstruction]:
    """Generate validated edit instructions by calling OpenRouter API.

    Returns a list of EditInstruction objects that have passed validation.
    """
    system_prompt = _load_system_prompt()
    user_prompt = _build_user_prompt(resume, signals)

    print(f"  Sending {len(resume.editable_blocks)} editable blocks to AI...")

    client = _get_client()
    message = client.chat.completions.create(
        model="google/gemini-2.0-flash-001",
        max_tokens=4096,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
    )

    raw_content = message.choices[0].message.content
    if not raw_content:
        print("  ERROR: AI returned empty response!", file=sys.stderr)
        sys.exit(1)

    raw = raw_content.strip()
    print(f"  AI response length: {len(raw)} chars")

    # Strip markdown fences if present
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)

    try:
        data = json.loads(raw)
    except json.JSONDecodeError as e:
        print(f"  ERROR: Failed to parse AI response as JSON: {e}", file=sys.stderr)
        print(f"  Raw response (first 500 chars): {raw[:500]}", file=sys.stderr)
        sys.exit(1)

    if not data:
        print("  ERROR: AI returned empty JSON object — no edits generated!", file=sys.stderr)
        sys.exit(1)

    print(f"  AI proposed {len(data)} edits")

    # Build a lookup for original text by index
    block_by_index = {b.index: b for b in resume.editable_blocks}

    edits = []
    skipped = []
    for idx_str, val in data.items():
        idx = int(idx_str)
        if idx not in block_by_index:
            skipped.append(f"index {idx} not in editable blocks")
            continue

        new_text = val["new_text"] if isinstance(val, dict) else val
        reason = val.get("reason", "") if isinstance(val, dict) else ""

        edit = EditInstruction(
            block_index=idx,
            original_text=block_by_index[idx].text,
            new_text=new_text,
            reason=reason,
        )

        is_valid, invalid_reason = _validate_edit(edit, resume)
        if is_valid:
            edits.append(edit)
        else:
            skipped.append(f"index {idx}: {invalid_reason}")

    if skipped:
        print(f"  Skipped {len(skipped)} edits: {'; '.join(skipped[:5])}", file=sys.stderr)

    if not edits:
        print("  ERROR: All AI edits were rejected by validation! 0 edits to apply.", file=sys.stderr)
        print("  This is a bug — the tailoring produced no usable output.", file=sys.stderr)
        sys.exit(1)

    return edits
