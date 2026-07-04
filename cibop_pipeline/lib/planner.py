"""Planning agent — produces a verified content plan before any generation."""

import re
import json
import streamlit as st
import anthropic
from lib.parser import parse_slide_range, get_slide_excerpt

QUESTION_TYPES = ["MCQ_INLINE", "TRUE_FALSE_MULTI", "GAP_FILL_DROPDOWN",
                  "MATCHING_PAIRS", "TABLE_SELECTION"]


def get_client():
    return anthropic.Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])


def build_plan(uors: list, slides: dict, extra_text: str = "") -> list:
    """
    For each UOR and each sub-competency, extract the exact slide excerpts
    and key terms. Returns a flat list of plan_items.

    This is the critical step that prevents hallucination:
    - Each plan item contains ONLY the slide text for that UOR's range
    - The generation step receives ONLY the plan item's excerpts
    """
    client = get_client()
    plan_items = []
    q_type_cycle = 0

    for uor in uors:
        slide_range = uor.get("slide_nos_source", "")
        start, end = parse_slide_range(slide_range)
        excerpt = get_slide_excerpt(slides, start, end)

        if not excerpt:
            st.warning(f"No slide text found for {uor['uor_id']} (slides {start}–{end})")

        sub_comps = uor.get("sub_competencies", [])
        extra_scs = uor.get("extra_scs", [])

        # Combine all SCs for this UOR
        all_scs = []
        for sc in sub_comps:
            all_scs.append({"sc_text": sc, "is_extra": False})
        for esc in extra_scs:
            all_scs.append({
                "sc_text": esc.get("sub_competency") or esc.get("title", ""),
                "is_extra": True,
                "extra_title": esc.get("title", ""),
                "extra_objective": esc.get("objective", "")
            })

        if not all_scs:
            # Treat the whole UOR as one SC
            all_scs = [{"sc_text": uor.get("objective", ""), "is_extra": False}]

        ears_raw = uor.get("ears", "")
        ears_list = [e.strip() for e in ears_raw.replace("·", ",").split(",") if e.strip()]

        # Extract required topics / key terms from the slides ONCE per UOR
        uor_key_terms = _extract_key_terms(client, uor.get("objective", ""), excerpt, uor["uor_id"])

        # Assign differentiated focus angles so SCs sharing the same slides cover different aspects
        sc_texts_raw = [sc["sc_text"].strip() for sc in all_scs if sc["sc_text"].strip()]
        focus_angles = _assign_focus_angles(client, sc_texts_raw, excerpt, uor["uor_id"],
                                            uor.get("title", ""))

        for sc_idx, sc in enumerate(all_scs):
            sc_text = sc["sc_text"].strip()
            if not sc_text:
                continue

            sc_id = f"SC{sc_idx + 1}"
            if sc_text.upper().startswith("SC"):
                parts = sc_text.split(" ", 1)
                if len(parts) > 0 and re.match(r'SC\d+\.\d+', parts[0]):
                    sc_id = parts[0]
                    sc_text = parts[1] if len(parts) > 1 else sc_text

            ear_verb = ears_list[sc_idx % len(ears_list)] if ears_list else "EXPLAIN"

            q_type = QUESTION_TYPES[q_type_cycle % len(QUESTION_TYPES)]
            q_type_cycle += 1

            # Attach the differentiated focus angle to the SC text so the generator uses it
            focus = focus_angles[sc_idx] if sc_idx < len(focus_angles) else sc_text
            enriched_sc_text = f"{sc_text}\n\n[FOCUS ANGLE] {focus}"

            plan_items.append({
                "uor_id": uor["uor_id"],
                "uor_title": uor.get("title", ""),
                "sc_id": sc_id,
                "sc_text": enriched_sc_text,
                "ear_verb": ear_verb,
                "slide_range_start": start,
                "slide_range_end": end,
                "slide_excerpts": excerpt,
                "key_terms": uor_key_terms,   # shared across SCs in this UOR
                "question_type": q_type,
                "plan_approved": False
            })

    return plan_items


def _assign_focus_angles(client: anthropic.Anthropic, sc_texts: list,
                         slide_excerpt: str, uor_id: str, uor_title: str) -> list:
    """
    Given multiple SCs that share the same slide range, assign each SC a DISTINCT
    focus angle so each video covers a different aspect of the slides.

    Returns a list of focus angle strings, one per SC.
    """
    if not sc_texts:
        return []
    if len(sc_texts) == 1:
        return [sc_texts[0]]

    try:
        sc_list = "\n".join(f"{i+1}. {s}" for i, s in enumerate(sc_texts))
        resp = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=600,
            messages=[{
                "role": "user",
                "content": (
                    f"UOR: {uor_id} — {uor_title}\n\n"
                    f"Slide content:\n{slide_excerpt[:3000]}\n\n"
                    f"These {len(sc_texts)} sub-competencies all map to the same slides:\n{sc_list}\n\n"
                    "Assign each SC a DISTINCT focus angle so each video script covers a "
                    "DIFFERENT aspect of the slides — no two SCs should produce similar scripts.\n"
                    "Each focus angle should be 1-2 sentences describing exactly which concepts "
                    "from the slides this SC should focus on, and what to emphasise.\n"
                    "Return ONLY a JSON array with one string per SC, in the same order.\n"
                    "Example: [\"Focus on defining the swap agreement structure: payment date and "
                    "accrual method. Emphasise the ISDA Master Agreement as the governing framework.\", "
                    "\"Focus on the parties to a swap: who First Party and Second Party are, and what "
                    "'underlyer' means (the underlying asset driving cash flows).\"]"
                )
            }]
        )
        text = resp.content[0].text.strip()
        match = re.search(r'\[.*\]', text, re.DOTALL)
        if match:
            angles = json.loads(match.group())
            # Ensure same length as input
            while len(angles) < len(sc_texts):
                angles.append(sc_texts[len(angles)])
            return angles[:len(sc_texts)]
    except Exception:
        pass
    return sc_texts  # fallback: use SC texts as-is


def _extract_key_terms(client: anthropic.Anthropic, sc_text: str,
                       slide_excerpt: str, uor_id: str) -> list:
    """Ask Claude to extract 5-8 key terms from the slide excerpt relevant to this SC."""
    if not slide_excerpt:
        return []
    try:
        resp = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=300,
            messages=[{
                "role": "user",
                "content": (
                    f"Sub-competency: {sc_text}\n\n"
                    f"Slide text:\n{slide_excerpt[:3000]}\n\n"
                    "List 5-8 key terms or phrases from the slide text that are directly "
                    "relevant to this sub-competency. Return ONLY a JSON array of strings. "
                    "Only include terms that actually appear in the slide text above."
                )
            }]
        )
        text = resp.content[0].text.strip()
        # Extract JSON array
        match = re.search(r'\[.*?\]', text, re.DOTALL)
        if match:
            return json.loads(match.group())
    except Exception:
        pass
    return []
