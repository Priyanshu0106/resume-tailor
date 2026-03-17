# Resume Tailoring Tool

## Mission
Build a tool that takes a .docx/.pdf resume + a job description and outputs a tailored resume that:
1. Preserves 100% of the original formatting (fonts, colors, spacing, section bars, tables, logos, bold patterns)
2. Only modifies text content — bullet points, skill keywords
3. Never adds/removes sections, changes layout structure, or alters visual design

## Critical Principle: EDIT, DON'T REGENERATE
The .docx is a ZIP of XML files. Unpack it, find the text nodes to change, swap the text while leaving all surrounding XML tags intact, then repack. NEVER use python-docx to open+save (it strips complex formatting).

## Architecture
```
INPUT: resume.docx/.pdf + job_description
  → 1. PARSE: Unpack .docx ZIP → extract word/document.xml → map editable text to XML locations
  → 2. PARSE JD: Extract required skills, preferred skills, key verbs, domain keywords
  → 3. AI TAILORING: Generate list of (original_text, new_text) replacements
  → 4. APPLY EDITS: Find-and-replace text in XML, preserve ALL <w:rPr> formatting
  → 5. REPACK: Rezip into .docx, validate
OUTPUT: tailored_resume.docx (identical formatting, tailored content)
```

## XML Editing Rules
- ONLY edit `word/document.xml` — never touch styles.xml, numbering.xml, media/, _rels/, [Content_Types].xml
- Only modify `<w:t>` text content within `<w:r>` runs
- NEVER delete `<w:r>`, `<w:rPr>`, or `<w:pPr>` elements
- NEVER add new `<w:r>` elements
- For multi-run paragraphs: put all new text in first `<w:r>`'s `<w:t>`, empty remaining `<w:t>` elements
- Preserve `xml:space="preserve"` attributes

## AI Tailoring Rules
- Each new_text must be approximately same length as original (±15%)
- If original starts with action verb, new_text must also start with strong action verb
- Preserve all numbers/metrics from original
- Maximum 60% of bullets should change; keep 40%+ unchanged
- Changes should be SUBTLE REWORDING, not wholesale rewriting
- Never fabricate experience, tools, or achievements not implied by original
- Never remove a bullet entirely — only reword it
- Skills line: reorder to front-load JD-relevant skills, may add 1-2 implied skills

## Editable vs Locked Content
| Content | Editable? |
|---|---|
| Bullet point text | YES |
| Skills line | YES |
| Project bullet text | YES |
| Company names | NO |
| Job titles | NO |
| Dates | NO |
| Section headers | NO |
| Education table | NO |
| Name/contact info | NO |
| Certificates | NO |

## PDF Handling
For PDF resumes, use PyMuPDF (fitz) for in-place text replacement:
1. Extract text blocks with bbox/font metadata
2. Identify modifiable blocks (same rules as docx)
3. Get AI-tailored text
4. Redact old text (white fill), insert new text at same position with same font/color

## Tech Stack
- FastAPI web framework
- OpenRouter API (via OpenAI SDK) using google/gemini-2.0-flash-001
- Direct XML editing via xml.etree.ElementTree for .docx
- PyMuPDF (fitz) for PDF editing
- Deployed on Render (https://resume-tailor-oz4o.onrender.com)
