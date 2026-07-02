"""Plan — generate and review the content plan before any content is created."""

import streamlit as st
import json
from lib.db import (get_files_for_topic, save_uors, get_uors,
                    save_plan_items, get_plan_items, approve_plan,
                    update_topic_status, update_plan_item)
from lib.parser import parse_slide_range, get_slide_excerpt
from lib.planner import build_plan
from lib.sidebar import render_topic_selector

st.set_page_config(page_title="Plan — CIBOP", layout="wide")
render_topic_selector()
st.title("📋 Content Plan")

if "selected_topic_id" not in st.session_state:
    st.warning("Please select a topic first on the **📂 Topics** page.")
    st.stop()

topic_id = st.session_state["selected_topic_id"]
topic_name = st.session_state.get("selected_topic_name", "")
st.subheader(f"Module: {topic_name}")

files = {f["file_type"]: f for f in get_files_for_topic(topic_id)}

if "ppt" not in files or "uor" not in files:
    st.warning("Please upload PPT and UOR Excel files in **⚙️ Setup** first.")
    st.stop()

# ── Build plan button ─────────────────────────────────────────────────────────
existing_plan = get_plan_items(topic_id)

col1, col2 = st.columns([2, 1])
with col1:
    if st.button("🔄 (Re)Build Content Plan", type="primary"):
        with st.spinner("Building plan — extracting slide ranges per sub-competency…"):
            # Parse sources
            slides_dict = {int(k): v for k, v in json.loads(files["ppt"]["extracted_text"]).items()}
            uors_raw = json.loads(files["uor"]["extracted_text"])

            # Add PDF text to last UOR's extra if PDF exists and any UOR references it
            pdf_text = ""
            if "pdf" in files:
                pdf_text = files["pdf"]["extracted_text"]

            # Save UOR records
            save_uors(topic_id, uors_raw)

            # Build the plan
            plan_items = build_plan(uors_raw, slides_dict, pdf_text)

            # For any UOR that references a PDF (has extra_scs with no slides),
            # append the PDF text to the slide excerpts
            if pdf_text:
                for item in plan_items:
                    if not item.get("slide_excerpts") and pdf_text:
                        item["slide_excerpts"] = pdf_text[:3000]

            save_plan_items(topic_id, plan_items)
            existing_plan = plan_items
            st.success(f"Plan built: {len(plan_items)} content items across all sub-competencies.")

with col2:
    if existing_plan:
        approved_count = sum(1 for i in existing_plan if i.get("plan_approved"))
        st.metric("Items", len(existing_plan))
        st.metric("Approved", approved_count)

# ── Display plan ──────────────────────────────────────────────────────────────
if not existing_plan:
    st.info("No plan yet. Click **Build Content Plan** above.")
    st.stop()

st.markdown("---")
st.markdown("""
**Review the plan below.** Each row = one video script that will be generated.
- ✅ Check that slide ranges match the UOR intent
- ✅ Check that key terms look right
- ✅ Edit slide ranges or question types if needed
- ✅ Then click **Approve Plan** to unlock generation
""")

# Group by UOR
from itertools import groupby
plan_by_uor = {}
for item in existing_plan:
    uid = item["uor_id"]
    plan_by_uor.setdefault(uid, []).append(item)

all_approved = all(i.get("plan_approved") for i in existing_plan)

for uor_id, items in plan_by_uor.items():
    uor_title = items[0].get("uor_title", "")
    with st.expander(f"**{uor_id}** — {uor_title}  ({len(items)} SCs)", expanded=True):
        for item in items:
            col_a, col_b, col_c, col_d = st.columns([2, 1, 1, 2])
            with col_a:
                st.markdown(f"**{item['sc_id']}** — {item['sc_text'][:80]}")
            with col_b:
                st.caption(f"Slides: {item['slide_range_start']}–{item['slide_range_end']}")
                st.caption(f"EAR: {item['ear_verb']}")
            with col_c:
                st.caption(f"Q Type: {item['question_type']}")
            with col_d:
                terms = item.get("key_terms", [])
                st.caption("Key terms: " + ", ".join(terms[:5]) + ("…" if len(terms) > 5 else ""))

            # Inline edit
            with st.expander(f"Edit {item['sc_id']}", expanded=False):
                new_start = st.number_input(
                    "Slide start", value=item["slide_range_start"],
                    key=f"start_{item.get('id', item['sc_id'])}"
                )
                new_end = st.number_input(
                    "Slide end", value=item["slide_range_end"],
                    key=f"end_{item.get('id', item['sc_id'])}"
                )
                new_qtype = st.selectbox(
                    "Question type",
                    ["MCQ_INLINE", "TRUE_FALSE_MULTI", "GAP_FILL_DROPDOWN",
                     "MATCHING_PAIRS", "TABLE_SELECTION"],
                    index=["MCQ_INLINE", "TRUE_FALSE_MULTI", "GAP_FILL_DROPDOWN",
                           "MATCHING_PAIRS", "TABLE_SELECTION"].index(
                               item.get("question_type", "MCQ_INLINE")),
                    key=f"qtype_{item.get('id', item['sc_id'])}"
                )
                if st.button("Save changes", key=f"save_{item.get('id', item['sc_id'])}"):
                    if "id" in item:
                        # Re-extract slide excerpts for new range
                        slides_dict = {int(k): v for k, v in
                                       json.loads(files["ppt"]["extracted_text"]).items()}
                        new_excerpt = get_slide_excerpt(slides_dict, int(new_start), int(new_end))
                        update_plan_item(item["id"], {
                            "slide_range_start": int(new_start),
                            "slide_range_end": int(new_end),
                            "question_type": new_qtype,
                            "slide_excerpts": new_excerpt
                        })
                        st.success("Saved.")
                        st.rerun()

st.markdown("---")

# ── Approve button ────────────────────────────────────────────────────────────
if not all_approved:
    if st.button("✅ Approve Plan — Unlock Generation", type="primary"):
        approve_plan(topic_id)
        update_topic_status(topic_id, "planned")
        st.success("Plan approved! Go to **🔨 Generate** to create content.")
        st.rerun()
else:
    st.success("✅ Plan is approved. Proceed to **🔨 Generate**.")
