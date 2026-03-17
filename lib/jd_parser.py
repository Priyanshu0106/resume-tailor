"""Parse a job description into structured signals using OpenRouter API."""

import json
import os
import re
from dataclasses import dataclass, field

from openai import OpenAI


@dataclass
class JDSignals:
    job_title: str = ""
    company: str = ""
    required_skills: list = field(default_factory=list)
    preferred_skills: list = field(default_factory=list)
    key_action_verbs: list = field(default_factory=list)
    domain_keywords: list = field(default_factory=list)
    seniority_level: str = ""
    focus_areas: list = field(default_factory=list)
    ats_keywords: list = field(default_factory=list)


_JD_SYSTEM_PROMPT = """You are a job description analyst. Extract structured signals from the given job description.

Return ONLY a JSON object with these keys:
- job_title: string
- company: string
- required_skills: list of strings (must-have skills/technologies)
- preferred_skills: list of strings (nice-to-have skills)
- key_action_verbs: list of strings (action verbs used in the JD like "manage", "develop", "lead")
- domain_keywords: list of strings (industry/domain terms)
- seniority_level: string (e.g., "entry", "mid", "senior", "lead", "manager")
- focus_areas: list of strings (main areas of responsibility)
- ats_keywords: list of strings (keywords likely used by ATS screening)

Return ONLY valid JSON, no markdown fences, no explanation."""


def _get_client():
    api_key = os.environ.get("OPENROUTER_API_KEY")
    if not api_key:
        raise ValueError("OPENROUTER_API_KEY environment variable is not set.")
    return OpenAI(base_url="https://openrouter.ai/api/v1", api_key=api_key)


def _fetch_url_text(url: str) -> str:
    """Fetch URL and extract visible text from HTML."""
    import requests

    resp = requests.get(url, timeout=15, headers={"User-Agent": "Mozilla/5.0"})
    resp.raise_for_status()
    html = resp.text
    # Strip tags for a rough text extraction
    text = re.sub(r"<script[^>]*>.*?</script>", "", html, flags=re.S | re.I)
    text = re.sub(r"<style[^>]*>.*?</style>", "", text, flags=re.S | re.I)
    text = re.sub(r"<[^>]+>", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def parse_jd(jd_text: str | None = None, jd_url: str | None = None) -> JDSignals:
    """Parse job description text or URL into JDSignals.

    Provide either jd_text (raw text) or jd_url (fetch from URL).
    """
    if not jd_text and not jd_url:
        raise ValueError("Provide either jd_text or jd_url.")

    if jd_url and not jd_text:
        jd_text = _fetch_url_text(jd_url)

    message = _get_client().chat.completions.create(
        model="google/gemini-2.0-flash-001",
        max_tokens=1500,
        messages=[
            {"role": "system", "content": _JD_SYSTEM_PROMPT},
            {"role": "user", "content": f"Analyze this job description:\n\n{jd_text}"},
        ],
    )

    raw = message.choices[0].message.content.strip()
    # Strip markdown fences if present
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)

    data = json.loads(raw)

    return JDSignals(
        job_title=data.get("job_title", ""),
        company=data.get("company", ""),
        required_skills=data.get("required_skills", []),
        preferred_skills=data.get("preferred_skills", []),
        key_action_verbs=data.get("key_action_verbs", []),
        domain_keywords=data.get("domain_keywords", []),
        seniority_level=data.get("seniority_level", ""),
        focus_areas=data.get("focus_areas", []),
        ats_keywords=data.get("ats_keywords", []),
    )
