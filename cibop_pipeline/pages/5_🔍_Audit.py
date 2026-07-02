"""Audit — independent scoring of every generated piece of content."""

import streamlit as st
import time
from lib.db import (get_plan_items, get_generated, save_audit,
                    get_all_audits_for_topic, update_topic_status)
from lib.auditor import audit_content, score_color, overall_pass
from lib.sidebar import render_topic_selector

st.set_page_config(page_title="Audit — CIBOP", layout="wide")
render_topic_selector()
st.title("🔍 Content Audit")

if "selected_topic_id" not in st.session_state:
    st.warning("Please select a topic first.")
    st.stop()

topic_id = st.session_state["selected_topic_id"]
topic_name = st.session_state.get("selected_topic_name", "")
st.subheader(f"Module: {topic_name}")

plan_items = get_plan_items(topic_id)
generated = get_generated(topic_id)

if not generated:
    st.warning("No generated content yet. Go to **🔨 Generate** first.")
    st.stop()

# ── Build lookup maps ─────────────────────────────────────────────────────────
plan_map = {f"{i['uor_id']}_{i['sc_id']}": i for i in plan_items}
gen_map = {}
for g in generated:
    key = f"{g['uor_id']}_{g['sc_id']}_{g['content_type']}"
    gen_map[key] = g

# ── Summary metrics ───────────────────────────────────────────────────────────
all_audit_data = get_all_audits_for_topic(topic_id)
audited = [a for a in all_audit_data if a["audit"] is not None]
passed = [a for a in audited if a["audit"].get("passed")]

col1, col2, col3, col4 = st.columns(4)
col1.metric("Total Content Items", len(all_audit_data))
col2.metric("Audited", len(audited))
col3.metric("Passed (≥80%)", len(passed))
col4.metric("Failed", len(audited) - len(passed))

st.markdown("---")

# ── Run audit button ──────────────────────────────────────────────────────────
col_a, col_b = st.columns(2)
with col_a:
    audit_type = st.radio("Audit", ["Video Scripts", "Assessment Questions", "Both"],
                          horizontal=True)
with col_b:
    only_unaudited = st.checkbox("Only audit unaudited items", value=True)

if st.button("▶️ Run Audit", type="primary"):
    items_to_audit = []
    for item in plan_items:
        ct_list = []
        if audit_type in ("Video Scripts", "Both"):
            ct_list.append("video_script")
        if audit_type in ("Assessment Questions", "Both"):
            ct_list.append("assessment")

        for ct in ct_list:
            gkey = f"{item['uor_id']}_{item['sc_id']}_{ct}"
            if gkey in gen_map:
                items_to_audit.append((item, gen_map[gkey], ct))

    if not items_to_audit:
        st.warning("No generated content found to audit.")
        st.stop()

    progress_bar = st.progress(0)
    status_area = st.empty()

    for idx, (plan_item, gen_item, ct) in enumerate(items_to_audit):
        label = f"{gen_item['uor_id']}/{gen_item['sc_id']} [{ct}]"
        status_area.info(f"Auditing {idx+1}/{len(items_to_audit)}: {label}")

        result = audit_content(
            plan_item,
            gen_item["content_text"],
            gen_item.get("slide_refs_used", [])
        )

        save_audit(
            gen_item["id"],
            result["coverage_score"],
            result["order_score"],
            result["fidelity_score"],
            result.get("flags", [])
        )

        progress_bar.progress((idx + 1) / len(items_to_audit))
        time.sleep(0.2)

    update_topic_status(topic_id, "audited")
    status_area.success("✅ Audit complete! See results below.")
    st.rerun()

# ── Audit results display ─────────────────────────────────────────────────────
st.markdown("---")
st.markdown("### Audit Results")

all_audit_data = get_all_audits_for_topic(topic_id)

if not any(a["audit"] for a in all_audit_data):
    st.info("No audit results yet. Run the audit above.")
    st.stop()

# Group by UOR
by_uor = {}
for row in all_audit_data:
    if not row["audit"]:
        continue
    uid = row["content"]["uor_id"]
    by_uor.setdefault(uid, []).append(row)

for uor_id in sorted(by_uor.keys()):
    rows = by_uor[uor_id]
    uor_pass = all(r["audit"].get("passed") for r in rows if r["audit"])
    header_icon = "✅" if uor_pass else "❌"

    with st.expander(f"{header_icon} **{uor_id}** — {len(rows)} items", expanded=not uor_pass):
        for row in rows:
            audit = row["audit"]
            content = row["content"]
            ct_label = "📹 Video" if content["content_type"] == "video_script" else "📝 Assessment"

            cov = audit.get("coverage_score", 0)
            ord_ = audit.get("order_score", 0)
            fid = audit.get("fidelity_score", 0)
            overall = (cov + ord_ + fid) / 3
            pass_icon = "✅" if overall >= 80 else "❌"

            st.markdown(
                f"{pass_icon} **{content['sc_id']}** {ct_label} — "
                f"Overall: **{overall:.0f}%** | "
                f"{score_color(cov)} Coverage {cov:.0f}% | "
                f"{score_color(ord_)} Order {ord_:.0f}% | "
                f"{score_color(fid)} Fidelity {fid:.0f}%"
            )

            # Show flags
            flags = audit.get("flags", [])
            if flags:
                for flag in flags:
                    flag_type = flag.get("type", "")
                    detail = flag.get("detail", "")
                    icon = "🔴" if flag_type in ("HALLUCINATION", "FIDELITY") else "🟡"
                    st.caption(f"{icon} [{flag_type}] {detail}")

            st.divider()
