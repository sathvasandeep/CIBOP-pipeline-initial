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

SOURCE SLIDES — defines the mandatory curriculum for this video:
{slide_excerpts}

REQUIRED KEY TERMS — every term below MUST appear clearly in the script:
{key_terms}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
CONTENT RULES
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
1. MANDATORY COVERAGE — every topic, concept, and term in the SOURCE SLIDES must be addressed.
   If the slides mention "OTC Derivative", "Notional Amount", "Two Counterparties",
   "ISDA Master Agreement" — those must appear in the script. Never skip a slide concept.

2. DOMAIN ENRICHMENT ALLOWED — the slides are often brief bullet points. You MUST enrich them
   with accurate capital markets knowledge to produce a high-quality educational script.
   Example: If slides say "swap = bilateral contract", a good script says:
   "A swap is an OTC derivative — a bilateral contract negotiated directly between two
   counterparties, typically governed by an ISDA Master Agreement. Because swaps are
   OTC, the terms are customisable: notional size, tenor, payment frequency, and
   reference rate. This flexibility is both a feature and a risk."
   That depth is expected and required.

3. FACTUAL ACCURACY — every enrichment must be correct financial fact.
   Do not invent regulations, companies, or numbers not grounded in standard finance.

4. SC FOCUS ANGLE — this script covers ONE specific angle per the Sub-Competency above.
   Focus tightly on that angle. Do not repeat explanations that belong to adjacent SCs.

5. DEFINITIONS — define every key term precisely on first use.
   For example: "underlyers" means the underlying assets (e.g. interest rates, equities,
   credit reference entities) whose values drive the cash flow calculations in the swap —
   NOT the parties themselves.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
CHARACTER GUIDE — 4 options:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
• MOTION — Full-screen branded animation. Use for opening scene, complex diagrams, formulas, flows. No human on screen.
• RYO — Male, mid-30s, navy suit. Expert/Mentor. Measured, dry wit. At whiteboard, desk, or camera.
• ARIA — Female, late-20s, teal blazer. Curious Analyst. Energetic, great comic timing. Reacts, questions.
• BOTH — Split screen (RYO left, ARIA right). For dialogue exchanges or summary scene.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SCENE STRUCTURE (8 scenes):
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Scene 1  (MOTION): Animated opening — introduce the topic with a graphic/diagram. On-Screen Text: module ID, video title, EAR verb.
Scene 2  (RYO):    Core concept 1 — whiteboard or desk. On-Screen Text: key term in CAPS + 2–3 line definition.
Scene 3  (ARIA):   Question or reaction — "Aria thinking/puzzled/excited". On-Screen Text: "Aria's Question: / [question]"
Scene 4  (RYO):    Core concept 2 with worked example or diagram. On-Screen Text: structured 3–4 line text with labels.
Scene 5  (MOTION): Infographic for a list, process, or formula. On-Screen Text: structured label–value pairs or bullets.
Scene 6  (ARIA):   Natural humor moment OR clarifying question. On-Screen Text: "Aria's Reaction: / [remark]"
Scene 7  (BOTH):   Summary split-screen. Aria summarises, Ryo confirms. On-Screen Text: "KEY PRINCIPLES: / • ... / • ... / • ..."
Scene 8  (RYO):    Tease next topic. On-Screen Text: "COMING UP: / [1-line tease]"

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
COLUMN INSTRUCTIONS:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Visual Cue / Animation — FULL production brief (2–4 sentences):
  • MOTION: what graphics animate, what labels/arrows appear, colour scheme
  • RYO/ARIA: character position, props, board writing, graphic behind them
  • BOTH: split-screen layout, what each side shows

On-Screen Text — MULTI-LINE structured text (use / for line breaks):
  • Scene 1: "{uor_id} | [Video Title] / {sc_id} / EAR: {ear_verb}"
  • RYO scenes: "TERM / definition line 1 / definition line 2" (3–5 lines)
  • ARIA: "Aria's Question: / [question]" or "Aria's Reaction: / [remark]"
  • MOTION: structured label-value pairs (4–6 lines)
  • BOTH: "KEY PRINCIPLES: / • bullet 1 / • bullet 2 / • bullet 3"

Voice Over — 2–4 sentences. Precise, rich, professional. Use domain knowledge to explain well.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
OUTPUT FORMAT — produce a markdown table with exactly 8 rows and 5 columns:
| # | Character | Visual Cue / Animation | On-Screen Text | Voice Over |

After the table, on a new line write:
SLIDE_REFS_USED: [comma-separated slide numbers you referenced, in ascending order]"""


ASSESSMENT_PROMPT = """You are writing a knowledge assessment question for CIBOP capital markets training.

UOR ID: {uor_id}
UOR Title: {uor_title}
Sub-Competency: {sc_text}
Question Type: {question_type}

SOURCE SLIDES — mandatory curriculum (correct answers must align with this):
{slide_excerpts}

KEY TERMS (must appear in correct answer or question):
{key_terms}

RULES:
1. The correct answer and explanation must accurately reflect the SOURCE SLIDES and sound capital markets knowledge.
2. Distractors (wrong options) should be plausible real-world alternatives — you may draw on domain knowledge for these.
3. Do NOT contradict the slide content in the correct answer.
4. Every key term above should appear somewhere in the question or answer choices.

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

SOURCE SLIDES (mandatory curriculum — all concepts must be covered):
{plan_item["slide_excerpts"][:6000]}

RULES:
1. Apply the reviewer's requested changes faithfully.
2. All topics and key terms from the SOURCE SLIDES must still be present after revision.
3. You MAY use accurate capital markets domain knowledge to enrich explanations.
4. Do NOT introduce factual errors or fabricate regulations/companies/numbers.
5. Slide references must remain in ascending order.
6. Keep the same 8-scene markdown table format:
   | # | Character | Visual Cue / Animation | On-Screen Text | Voice Over |
7. Preserve RYO (expert/mentor, dry wit) and ARIA (curious analyst, comic timing) voices.
8. Visual Cue: 2–4 sentence production brief. On-Screen Text: multi-line with / separators.

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

SOURCE SLIDES (mandatory curriculum):
{plan_item["slide_excerpts"][:5000]}

RULES:
1. Apply the reviewer's requested changes faithfully.
2. Questions and answers must be grounded in the SOURCE SLIDES.
3. You may use domain knowledge for plausible distractors, clearly wrong but realistic.
4. Keep the same question format as the original.

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
