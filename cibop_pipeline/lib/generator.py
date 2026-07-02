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

OUTPUT FORMAT — produce a markdown table with exactly 8 rows and 5 columns:
| # | Character | Visual Cue / Animation | On-Screen Text | Voice Over |

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
CHARACTER GUIDE — 4 options:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
• MOTION — Full-screen branded animation. Use for opening scene, complex concept diagrams, formulas, flow charts, infographics. No human on screen.
• RYO — Male, mid-30s, navy suit, calm and precise. Expert/Mentor. Dry wit. At whiteboard, desk, or camera.
• ARIA — Female, late-20s, teal blazer, animated and curious. Curious Analyst. Comic timing. Reacts, questions, connects.
• BOTH — Split screen (RYO left, ARIA right). Use for dialogue exchanges or summary scene.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SCENE STRUCTURE (8 scenes):
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Scene 1  (MOTION): Animated opening — introduce the topic visually with a graphic or diagram that sets the scene. On-Screen Text shows module ID, video title, EAR verb.
Scene 2  (RYO):    Core concept 1 from slides — whiteboard or desk. On-Screen Text: key term in CAPS + 2–3-line definition/formula from the slides.
Scene 3  (ARIA):   Question or reaction to Scene 2 — "Aria thinking / puzzled / excited". On-Screen Text: "Aria's Question:" then 1-line question.
Scene 4  (RYO):    Core concept 2 from slides — prop, diagram, or worked example. On-Screen Text: structured 3–4 line text with labels and data from slide.
Scene 5  (MOTION): Infographic or animated diagram for a list, process, or formula from the slides. On-Screen Text: structured label-value pairs or bullet list.
Scene 6  (ARIA):   Moment of natural humor OR clarifying question. On-Screen Text: "Aria's Reaction:" then 1-line witty or curious remark.
Scene 7  (BOTH):   Summary split-screen exchange — Aria summarises in her own words, Ryo confirms/adds nuance. On-Screen Text: "KEY PRINCIPLES:" then 3 bullet lines.
Scene 8  (RYO):    Tease next topic. Ryo straightens jacket, looks at camera. On-Screen Text: "COMING UP:" then 1-line tease.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
COLUMN INSTRUCTIONS:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Visual Cue / Animation — write a FULL production brief (2–4 sentences):
  • For MOTION: describe the animation (what graphics appear, what animates, what labels/arrows are shown, colours or icons if relevant)
  • For RYO/ARIA: describe character position (whiteboard / desk / camera), props (contract, diagram, board writing), and any key graphic behind them
  • For BOTH: describe split-screen layout and what each side shows

On-Screen Text — write MULTI-LINE structured text (use forward slash / to represent line breaks):
  • Scene 1: "{uor_id} | Video Title / {sc_id} / EAR: {ear_verb}"
  • RYO scenes: "TERM IN CAPS / definition line 1 / definition line 2" (3–5 lines, structured)
  • ARIA scenes: "Aria's Question: / [one-line question]" or "Aria's Reaction: / [one-line comment]"
  • MOTION scenes: structured label–value pairs or numbered list from the slide (4–6 lines)
  • BOTH scenes: "KEY PRINCIPLES: / • bullet 1 / • bullet 2 / • bullet 3"

Voice Over — 2–3 sentences max. Conversational but precise. Every word must be grounded in the SOURCE SLIDE TEXT.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

After the table, on a new line write:
SLIDE_REFS_USED: [comma-separated slide numbers you referenced, in ascending order]"""


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


def revise_content(plan_item: dict, original_text: str, reviewer_comments: str,
                   content_type: str = "video_script") -> tuple[str, list[int]]:
    """Revise generated content based on reviewer comments.

    The revision agent receives the original script + comments + slide text.
    It may only draw on slide content — no new hallucination allowed.
    """
    client = get_client()

    if content_type == "video_script":
        prompt = f"""You are revising a video production script for a capital markets training programme (CIBOP).

UOR ID: {plan_item["uor_id"]}
UOR Title: {plan_item["uor_title"]}
Sub-Competency: {plan_item["sc_text"]}
Slide Range: {plan_item["slide_range_start"]}–{plan_item["slide_range_end"]}

ORIGINAL SCRIPT:
{original_text}

REVIEWER COMMENTS / REQUESTED CHANGES:
{reviewer_comments}

SOURCE SLIDE TEXT (the ONLY content you may add or draw from):
{plan_item["slide_excerpts"][:6000]}

RULES — STRICTLY ENFORCED:
1. Apply the reviewer's requested changes faithfully.
2. Every fact, term, and example you add or keep must come from the SOURCE SLIDE TEXT.
3. Do NOT introduce any concept, person, company, or regulation not in the slide text.
4. Slide references must remain in ascending order.
5. Keep the same 8-scene markdown table format:
   | # | Character | Visual Cue | On-Screen Text | Voice Over |
6. Preserve RYO (dry wit, expert mentor) and ARIA (curious analyst, comic timing) voices.
7. Keep Voice Over to 2–3 sentences per scene. On-Screen Text to 5 words max.

Output ONLY the revised table, then on a new line:
SLIDE_REFS_USED: [comma-separated slide numbers in ascending order]"""

    else:  # assessment
        prompt = f"""You are revising a knowledge assessment question for CIBOP capital markets training.

UOR ID: {plan_item["uor_id"]}
Sub-Competency: {plan_item["sc_text"]}
Question Type: {plan_item.get("question_type", "MCQ_INLINE")}

ORIGINAL QUESTION:
{original_text}

REVIEWER COMMENTS / REQUESTED CHANGES:
{reviewer_comments}

SOURCE SLIDE TEXT (the ONLY content you may draw from):
{plan_item["slide_excerpts"][:5000]}

RULES — STRICTLY ENFORCED:
1. Apply the reviewer's requested changes faithfully.
2. Every option, answer, and explanation must come from the SOURCE SLIDE TEXT.
3. Keep the same question format as the original.
4. Do NOT introduce any fact or example not in the slide text.

Output ONLY the revised question, then:
SLIDE_REFS_USED: [slide numbers, ascending order]"""

    resp = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=2000,
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
