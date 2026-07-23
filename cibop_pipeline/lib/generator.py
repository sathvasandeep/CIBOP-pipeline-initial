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
   Read the [FOCUS ANGLE] section in the Sub-Competency field above and follow it precisely.

5. DEFINITIONS — define every key term precisely on first use.
   For example: "underlyers" means the underlying assets (e.g. interest rates, equities,
   credit reference entities) whose values drive the cash flow calculations in the swap —
   NOT the parties themselves.
   Every term in the REQUIRED KEY TERMS list must be defined or explained in a Voice Over.

6. THIS VIDEO IS STANDALONE — the video exists independently of any presentation deck.
   NEVER say "Slide X", "as shown in slide", "refer to the slide", or any slide reference
   in ANY column (Visual Cue, On-Screen Text, or Voice Over). The audience will not see
   slides. Write as if you are delivering a standalone educational video.

7. DEPTH OVER BREVITY — a Voice Over that says "this is a type of swap" is useless.
   Each RYO scene must deliver a crisp, accurate financial explanation. Example of the
   expected depth:
   "An Interest Rate Swap is an OTC derivative where two counterparties exchange cash
    flows — one paying a fixed rate, the other a floating rate benchmarked to SOFR or
    EURIBOR — calculated on the same notional principal. No principal is exchanged;
    only the net difference in interest obligations changes hands on each settlement date."
   That specificity is required.

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
Scene 8  (RYO):    Closing summary. Ryo faces camera directly, delivers a concise and memorable
               distillation of the ONE core insight from this video. No transition to next topic.
               Voice Over: 3–4 spoken sentences that crystallise the key concept — what it IS,
               why it MATTERS, and what a practitioner needs to remember.
               On-Screen Text: "KEY TAKEAWAY: / [1-sentence distillation of the core concept]"

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
COLUMN INSTRUCTIONS:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Visual Cue / Animation — FULL production brief (2–4 sentences):
  • MOTION: what graphics animate, what labels/arrows appear, colour scheme
  • RYO/ARIA: character position, props, board writing, graphic behind them
  • BOTH: split-screen layout, what each side shows

On-Screen Text — MULTI-LINE structured text (use / for line breaks ONLY):
  ⚠️ Use / (forward slash) for line breaks. NEVER use \ (backslash) as a separator.
  ⚠️ Keep OST to 4–6 short lines maximum. Do NOT put VO-style prose here.
  • Scene 1: "{uor_id} | [Video Title] / {sc_id} / EAR: {ear_verb}"
  • RYO scenes: "TERM / definition line 1 / definition line 2" (3–5 lines)
  • ARIA: "Aria's Question: / [question]" or "Aria's Reaction: / [remark]"
  • MOTION: structured label-value pairs (4–6 lines max — do NOT list every item)
  • BOTH: "KEY PRINCIPLES: / • bullet 1 / • bullet 2 / • bullet 3"
  • Scene 8 (closing): "KEY TAKEAWAY: / [1-sentence distillation of the core concept]"
    Do NOT write "COMING UP:" or any reference to the next video.

Voice Over — SPOKEN NARRATION. 3–5 complete sentences a narrator reads aloud.
  ⚠️ CRITICAL — Voice Over is AUDIO ONLY. It never appears on screen.
  ⚠️ NEVER put bullet points, labels, headers, or / separators in Voice Over.
  ⚠️ NEVER copy OST content into Voice Over. They serve DIFFERENT purposes.
  • MOTION scenes: VO narrates/explains the animation while the audience watches.
  • RYO/ARIA scenes: VO is exactly what the character says on camera.
  ❌ BAD VO (rejected): "Broker-Dealer Firms / Investment Bank Divisions"
  ✅ GOOD VO: "The STO acts as the operational engine of the investment bank's
     client franchise. Trading Operations handle every order from receipt through
     to settlement. Client Services manages the ongoing institutional relationships,
     while Market Insights provides real-time intelligence on prices and liquidity
     conditions — information that allows clients to time their transactions well."

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
        max_tokens=4000,
        messages=[{"role": "user", "content": prompt}]
    )
    if resp.stop_reason == "max_tokens":
        raise ValueError(
            "⚠️ The generated script was cut off — output exceeded the token limit. "
            "The script may be incomplete (missing scenes or empty cells). "
            "Try re-generating, or reduce the number of key terms in the plan."
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
    if resp.stop_reason == "max_tokens":
        raise ValueError(
            "⚠️ The assessment question was cut off — output exceeded the token limit. "
            "Try re-generating."
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
1. Apply the reviewer's requested changes faithfully and completely.
   — If the comment refers to a SPECIFIC SCENE (e.g. "add VO to scene 6"), modify ONLY that
     cell/column in that row. Copy every other row and cell EXACTLY from the original.
   — If the comment is general (e.g. "tighten the language"), revise all relevant scenes.
   — NEVER leave a cell empty if the reviewer explicitly asks you to add content to it.
   — NEVER delete existing content from scenes that were not mentioned in the comments.
2. All topics and key terms from the SOURCE SLIDES must still be present after revision.
3. You MAY use accurate capital markets domain knowledge to enrich explanations.
4. Do NOT introduce factual errors or fabricate regulations/companies/numbers.
5. Slide references must remain in ascending order.
6. Keep the same 8-scene markdown table format — output ALL 8 rows every time:
   | # | Character | Visual Cue / Animation | On-Screen Text | Voice Over |
   Scene 8 is ALWAYS a closing summary (not a "COMING UP" tease):
   — OST: "KEY TAKEAWAY: / [1-sentence distillation of the core concept]"
   — VO: 3–4 spoken sentences crystallising what the concept IS, why it MATTERS,
     and what a practitioner must remember. Rich, complete, standalone.
7. Preserve RYO (expert/mentor, dry wit) and ARIA (curious analyst, comic timing) voices.
8. Visual Cue: 2–4 sentence production brief.
   On-Screen Text: structured text with / (forward slash) line breaks — max 6 lines. Never use \.
9. Voice Over: SPOKEN NARRATION — 3–5 complete sentences read aloud by the narrator.
   ⚠️ Voice Over must be flowing prose. NO bullets, labels, headers, or / separators.
   ⚠️ Never copy On-Screen Text into Voice Over. They serve different purposes.
   ⚠️ Never leave Voice Over empty — every scene must have 3–5 spoken sentences.
   If reviewer says VO "looks like OST" or "is random", rewrite the VO as full spoken
   prose that explains the concept to the audience while they watch the scene.

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
        max_tokens=4000,
        messages=[{"role": "user", "content": prompt}]
    )
    if resp.stop_reason == "max_tokens":
        raise ValueError(
            "⚠️ The AI revision was cut off — output exceeded the token limit and the script "
            "is incomplete (likely missing scenes or empty Voice Over / On-Screen Text cells). "
            "Try breaking your comment into smaller targeted changes and re-running AI Edits."
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
