"""Generate — create video scripts and assessment questions SC by SC."""

import streamlit as st
import time
from lib.db import (get_plan_items, get_files_for_topic,
                    save_generated, get_generated, update_topic_status)
from lib.generator import generate_video_script, generate_assessment_question

st.set_page_config(page_title="Generate — CIBOP", layout="wide")
st.title("🔨 Generate Content")

if "selected_topic_id" not in st.session_state:
    st.warning("Please select a topic first.")
    st.stop()

topic_id = st.session_state["selected_topic_id"]
topic_name = st.session_state.get("selected_topic_name", "")
st.subheader(f"Module: {topic_name}")

# ── Guard: plan must be approved ──────────────────────────────────────────────
plan_items = get_plan_items(topic_id)
if not plan_items:
    st.warning("No content plan found. Go to **📋 Plan** first.")
    st.stop()

unapproved = [i for i in plan_items if not i.get("plan_approved")]
if unapproved:
    st.error(f"Plan not yet approved ({len(unapproved)} items pending). Approve the plan in **📋 Plan** before generating.")
    st.stop()

# ── Status overview ───────────────────────────────────────────────────────────
existing_scripts = {f"{g['uor_id']}_{g['sc_id']}_{g['content_type']}": g
                    for g in get_generated(topic_id)}

total = len(plan_items)
scripts_done = sum(1 for i in plan_items
                   if f"{i['uor_id']}_{i['sc_id']}_video_script" in existing_scripts)
questions_done = sum(1 for i in plan_items
                     if f"{i['uor_id']}_{i['sc_id']}_assessment" in existing_scripts)

col1, col2, col3 = st.columns(3)
col1.metric("Total SCs", total)
col2.metric("Video Scripts", f"{scripts_done}/{total}")
col3.metric("Assessment Qs", f"{questions_done}/{total}")

st.markdown("---")

# ── Generation controls ───────────────────────────────────────────────────────
col_a, col_b = st.columns(2)
with col_a:
    gen_type = st.radio("Generate", ["Video Scripts", "Assessment Questions", "Both"],
                        horizontal=True)
with col_b:
    only_missing = st.checkbox("Skip already generated (only generate missing)", value=True)
    selected_uor = st.selectbox(
        "Generate for UOR (or All)",
        ["All"] + sorted(set(i["uor_id"] for i in plan_items))
    )

if st.button("▶️ Start Generation", type="primary"):
    items_to_process = plan_items
    if selected_uor != "All":
        items_to_process = [i for i in plan_items if i["uor_id"] == selected_uor]

    progress_bar = st.progress(0)
    status_area = st.empty()
    log_area = st.container()
    total_items = len(items_to_process)

    for idx, item in enumerate(items_to_process):
        key_v = f"{item['uor_id']}_{item['sc_id']}_video_script"
        key_a = f"{item['uor_id']}_{item['sc_id']}_assessment"

        label = f"{item['uor_id']} / {item['sc_id']} — {item['sc_text'][:50]}"
        status_area.info(f"Processing {idx+1}/{total_items}: {label}")

        # ── Video Script ──────────────────────────────────────────────────────
        if gen_type in ("Video Scripts", "Both"):
            if only_missing and key_v in existing_scripts:
                log_area.caption(f"⏭️ Skipped (exists): {item['uor_id']}/{item['sc_id']} video")
            else:
                try:
                    script_text, slide_refs = generate_video_script(item)
                    save_generated(
                        topic_id, item["uor_id"], item["sc_id"],
                        "video_script", script_text, slide_refs
                    )
                    existing_scripts[key_v] = True
                    log_area.success(f"✅ Video: {item['uor_id']}/{item['sc_id']} — refs: {slide_refs}")
                except Exception as e:
                    log_area.error(f"❌ Video failed {item['uor_id']}/{item['sc_id']}: {e}")

        # ── Assessment Question ───────────────────────────────────────────────
        if gen_type in ("Assessment Questions", "Both"):
            if only_missing and key_a in existing_scripts:
                log_area.caption(f"⏭️ Skipped (exists): {item['uor_id']}/{item['sc_id']} assessment")
            else:
                try:
                    q_text, slide_refs = generate_assessment_question(item)
                    save_generated(
                        topic_id, item["uor_id"], item["sc_id"],
                        "assessment", q_text, slide_refs
                    )
                    existing_scripts[key_a] = True
                    log_area.success(f"✅ Assessment: {item['uor_id']}/{item['sc_id']}")
                except Exception as e:
                    log_area.error(f"❌ Assessment failed {item['uor_id']}/{item['sc_id']}: {e}")

        progress_bar.progress((idx + 1) / total_items)
        time.sleep(0.3)  # Brief pause to avoid rate limiting

    update_topic_status(topic_id, "generated")
    status_area.success("✅ Generation complete! Go to **🔍 Audit** to score the content.")

# ── Preview generated content ─────────────────────────────────────────────────
st.markdown("---")
st.markdown("### Generated Content Preview")

generated = get_generated(topic_id)
if not generated:
    st.info("Nothing generated yet.")
else:
    view_type = st.radio("View", ["Video Scripts", "Assessment Questions"], horizontal=True)
    ct = "video_script" if view_type == "Video Scripts" else "assessment"
    items = [g for g in generated if g["content_type"] == ct]

    for g in items:
        with st.expander(f"{g['uor_id']} / {g['sc_id']}", expanded=False):
            st.markdown(g["content_text"])
            if g.get("slide_refs_used"):
                st.caption(f"Slide refs: {g['slide_refs_used']}")
