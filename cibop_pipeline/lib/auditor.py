"""Audit agent — independently scores generated content against the content plan."""

import re
import streamlit as st
import anthropic


def get_client():
    return anthropic.Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])


AUDIT_PROMPT = """You are an independent content auditor for a capital markets training programme.

You will audit the following generated content against its plan. Be strict and objective.

=== CONTENT PLAN ===
UOR ID: {uor_id}
Sub-Competency: {sc_text}
Approved Slide Range: {slide_start}–{slide_end}
Key Terms that MUST appear: {key_terms}
EAR Verb: {ear_verb}

=== SLIDE SOURCE TEXT (ground truth) ===
{slide_excerpts}

=== GENERATED CONTENT TO AUDIT ===
{generated_content}

=== SLIDE REFERENCES CLAIMED ===
{slide_refs}

AUDIT CHECKS — score each 0 to 100:

1. COVERAGE SCORE (0–100):
   - Does the content address the sub-competency stated in the plan? (40 pts)
   - Do all key terms from the plan appear in the content? (30 pts)
   - Is the EAR verb ({ear_verb}) evident in the learning outcome? (30 pts)

2. ORDER SCORE (0–100):
   - Are slide references cited in ascending order (low to high slide numbers)? (50 pts)
   - Do all claimed slide numbers fall within the approved range {slide_start}–{slide_end}? (50 pts)
   Note: if no slide refs are cited, score is 0.

3. FIDELITY SCORE (0–100):
   - Is every fact, term, and example in the content traceable to the slide source text? (60 pts)
   - Are there any concepts, companies, regulations, or examples NOT in the slide text? (40 pts, deduct for each hallucination)

Return your audit as valid JSON only (no extra text):
{{
  "coverage_score": <number 0-100>,
  "order_score": <number 0-100>,
  "fidelity_score": <number 0-100>,
  "flags": [
    {{"type": "COVERAGE|ORDER|FIDELITY|HALLUCINATION", "detail": "<specific issue>"}}
  ],
  "summary": "<one sentence overall assessment>"
}}"""


def audit_content(plan_item: dict, generated_text: str, slide_refs: list) -> dict:
    """
    Audit generated content against its plan item.
    Returns dict with coverage_score, order_score, fidelity_score, flags.
    """
    client = get_client()

    prompt = AUDIT_PROMPT.format(
        uor_id=plan_item["uor_id"],
        sc_text=plan_item["sc_text"],
        slide_start=plan_item["slide_range_start"],
        slide_end=plan_item["slide_range_end"],
        key_terms=", ".join(plan_item.get("key_terms", [])),
        ear_verb=plan_item["ear_verb"],
        slide_excerpts=plan_item["slide_excerpts"][:4000],
        generated_content=generated_text[:4000],
        slide_refs=str(slide_refs)
    )

    try:
        resp = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=800,
            messages=[{"role": "user", "content": prompt}]
        )
        text = resp.content[0].text.strip()

        # Extract JSON
        match = re.search(r'\{.*\}', text, re.DOTALL)
        if match:
            import json
            result = json.loads(match.group())
            return {
                "coverage_score": float(result.get("coverage_score", 0)),
                "order_score": float(result.get("order_score", 0)),
                "fidelity_score": float(result.get("fidelity_score", 0)),
                "flags": result.get("flags", []),
                "summary": result.get("summary", "")
            }
    except Exception as e:
        return {
            "coverage_score": 0,
            "order_score": 0,
            "fidelity_score": 0,
            "flags": [{"type": "ERROR", "detail": str(e)}],
            "summary": "Audit failed due to error."
        }

    return {
        "coverage_score": 0,
        "order_score": 0,
        "fidelity_score": 0,
        "flags": [{"type": "ERROR", "detail": "Could not parse audit response."}],
        "summary": "Audit response could not be parsed."
    }


def score_color(score: float) -> str:
    """Return a colour string for a score."""
    if score >= 80:
        return "🟢"
    elif score >= 60:
        return "🟡"
    else:
        return "🔴"


def overall_pass(coverage: float, order: float, fidelity: float) -> bool:
    overall = (coverage + order + fidelity) / 3
    return overall >= 80.0
