"""Planning agent — produces a verified content plan before any generation."""

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

        for sc_idx, sc in enumerate(all_scs):
            sc_text = sc["sc_text"].strip()
            if not sc_text:
                continue

            sc_id = f"SC{sc_idx + 1}"
            if sc_text.upper().startswith("SC"):
                # Extract existing SC ID like SC1.1
                parts = sc_text.split(" ", 1)
                if len(parts) > 0 and re.match(r'SC\d+\.\d+', parts[0]):
                    sc_id = parts[0]
                    sc_text = parts[1] if len(parts) > 1 else sc_text

            # Pick EAR verb: cycle through the list for variety
            ear_verb = ears_list[sc_idx % len(ears_list)] if ears_list else "EXPLAIN"

            # Assign question type cyclically
            q_type = QUESTION_TYPES[q_type_cycle % len(QUESTION_TYPES)]
            q_type_cycle += 1

            # Use AI to extract key terms from the relevant slide excerpt
            key_terms = _extract_key_terms(client, sc_text, excerpt, uor["uor_id"])

            plan_items.append({
                "uor_id": uor["uor_id"],
                "uor_title": uor.get("title", ""),
                "sc_id": sc_id,
                "sc_text": sc_text,
                "ear_verb": ear_verb,
                "slide_range_start": start,
                "slide_range_end": end,
                "slide_excerpts": excerpt,
                "key_terms": key_terms,
                "question_type": q_type,
                "plan_approved": False
            })

    return plan_items


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
        import re
        match = re.search(r'\[.*?\]', text, re.DOTALL)
        if match:
            return json.loads(match.group())
    except Exception:
        pass
    return []


import re
