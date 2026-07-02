"""Generation agent — produces content for ONE SC at a time from its plan item."""

import re
import streamlit as st
import anthropic


def get_client():
    return anthropic.Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])


VIDEO_SCRIPT_PROMPT = """You are writing a video production script for a capital markets training programme (CIBOP).

UOR ID: {uor_id}
UOR Title: {uor_title}
Sub-Competency ID: {sc_id}
Sub-Competency: {sc_text}
EAR Verb: {ear_verb}
Slide Range: {slide_start}–{slide_end}

SOURCE SLIDE TEXT (the ONLY content you may use):
{slide_excerpts}

KEY TERMS from this SC (must appear in the script):
{key_terms}

RULES — STRICTLY ENFORCED:
1. Every fact, term, example, and number in your script must come from the SOURCE SLIDE TEXT above.
2. Do NOT introduce any concept, person, company, regulation, or example not in the slide text.
3. Reference slide numbers in ascending order (e.g., slide 4 before slide 7, never reverse).
4. If the slides have no text about a concept, do not include that concept.

OUTPUT FORMAT — produce a markdown table with 8 rows and 5 columns:
| # | Character | Visual Cue | On-Screen Text | Voice Over |

Characters: RYO (male, navy suit, expert mentor, dry wit) or ARIA (female, teal blazer, curious analyst, comic timing)

Scene structure:
- Scene 1 (ARIA): Hook question based ONLY on what the slides say
- Scenes 2–6 (RYO + ARIA): Core content. RYO explains from slides, ARIA reacts/asks follow-ups. Include one moment of natural humor per UOR (relatable analogy or witty exchange — about a concept in the slides)
- Scene 7 (ARIA): Summary in her own words — paraphrasing only what RYO said
- Scene 8 (RYO): Tease next topic (keep vague if unknown)

After the table, on a new line write:
SLIDE_REFS_USED: [comma-separated slide numbers you referenced, in ascending order]

Keep the Voice Over for each scene to 2–3 sentences max. Keep On-Screen Text to 5 words max."""


ASSESSMENT_PROMPT = """You are writing a knowledge assessment question for CIBOP capital markets training.

UOR ID: {uor_id}
UOR Title: {uor_title}
Sub-Competency: {sc_text}
Question Type: {question_type}

SOURCE SLIDE TEXT (the ONLY content you may use):
{slide_excerpts}

KEY TERMS (all correct answers must come from these or the slide text):
{key_terms}

RULES — STRICTLY ENFORCED:
1. Every option, answer, and explanation must come directly from the SOURCE SLIDE TEXT.
2. Do NOT introduce any fact, example, regulation, or company not in the slide text.
3. Distractors must also be plausible but clearly wrong based on the slides.

Write ONE {question_type} question using this format:

MCQ_INLINE format:
Q: [question]
A) [option]
B) [option]
C) [option]
D) [option]
✅ ANSWER: [letter] — [1-sentence explanation citing slide text]

TRUE_FALSE_MULTI format:
Statement 1: [statement] → TRUE/FALSE
Statement 2: [statement] → TRUE/FALSE
Statement 3: [statement] → TRUE/FALSE
✅ ANSWERS: [list correct answers with brief explanation from slide]

GAP_FILL_DROPDOWN format:
[Sentence with ___BLANK1___ and ___BLANK2___]
BLANK1 options: [a, b, c]
BLANK2 options: [a, b, c]
✅ ANSWERS: BLANK1=[correct], BLANK2=[correct] — [explanation from slide]

MATCHING_PAIRS format:
Match the term to its definition:
Term 1: ___  |  Term 2: ___  |  Term 3: ___
A. [definition]  B. [definition]  C. [definition]
✅ ANSWERS: 1→[letter], 2→[letter], 3→[letter] — [explanation from slide]

TABLE_SELECTION format:
Select all that apply — [question]:
| Item | Correct? |
| [item 1] | |
| [item 2] | |
| [item 3] | |
| [item 4] | |
✅ ANSWERS: [items that apply] — [explanation from slide]

After the question, write:
SLIDE_REFS_USED: [slide numbers referenced, ascending order]"""


def generate_video_script(plan_item: dict) -> tuple[str, list[int]]:
    """Generate a video script for one SC. Returns (script_text, slide_refs)."""
    client = get_client()
    prompt = VIDEO_SCRIPT_PROMPT.format(
        uor_id=plan_item["uor_id"],
        uor_title=plan_item["uor_title"],
        sc_id=plan_item["sc_id"],
        sc_text=plan_item["sc_text"],
        ear_verb=plan_item["ear_verb"],
        slide_start=plan_item["slide_range_start"],
        slide_end=plan_item["slide_range_end"],
        slide_excerpts=plan_item["slide_excerpts"][:6000],
        key_terms=", ".join(plan_item.get("key_terms", []))
    )

    resp = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=2000,
        messages=[{"role": "user", "content": prompt}]
    )
    text = resp.content[0].text
    slide_refs = _extract_slide_refs(text)
    return text, slide_refs


def generate_assessment_question(plan_item: dict) -> tuple[str, list[int]]:
    """Generate one assessment question for one SC. Returns (question_text, slide_refs)."""
    client = get_client()
    prompt = ASSESSMENT_PROMPT.format(
        uor_id=plan_item["uor_id"],
        uor_title=plan_item["uor_title"],
        sc_text=plan_item["sc_text"],
        question_type=plan_item["question_type"],
        slide_excerpts=plan_item["slide_excerpts"][:5000],
        key_terms=", ".join(plan_item.get("key_terms", []))
    )

    resp = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1200,
        messages=[{"role": "user", "content": prompt}]
    )
    text = resp.content[0].text
    slide_refs = _extract_slide_refs(text)
    return text, slide_refs


def _extract_slide_refs(text: str) -> list[int]:
    """Extract the SLIDE_REFS_USED line from generated text."""
    match = re.search(r'SLIDE_REFS_USED:\s*\[?([\d,\s]+)\]?', text)
    if match:
        refs_str = match.group(1)
        try:
            return sorted(set(int(x.strip()) for x in refs_str.split(",") if x.strip().isdigit()))
        except Exception:
            pass
    return []
