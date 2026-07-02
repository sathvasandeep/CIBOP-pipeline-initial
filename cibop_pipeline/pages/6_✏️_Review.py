"""Review — business owners read, edit inline, approve or reject content."""

import streamlit as st
from lib.db import (get_generated, get_plan_items, get_all_audits_for_topic,
                    update_generated_text, save_review, get_review,
                    update_topic_status, save_audit)
from lib.auditor import audit_content

st.set_page_config(page_title="Review — CIBOP", layout="wide")
st.title("✏️ Review & Approve")

if "selected_topic_id" not in st.session_state:
    st.warning("Please select a topic first.")
    st.stop()

topic_id = st.session_state["selected_topic_id"]
topic_name = st.session_state.get("selected_topic_name", "")
st.subheader(f"Module: {topic_name}")

st.info("""
**Instructions for reviewers:**
- Read each script carefully
- Edit the text box inline if changes are needed
- Click **Save Edits & Re-audit** to re-score after editing
- Click **Approve** to mark as final, or **Reject** to send back for regeneration
- Only approved content will appear in the Export
""")

reviewer_name = st.text_input("Your name (for review record)", placeholder="e.g. Priya Sharma")

# ── Load data ─────────────────────────────────────────────────────────────────
generated = get_generated(topic_id)
plan_items = get_plan_items(topic_id)
plan_map = {f"{i['uor_id']}_{i['sc_id']}": i for i in plan_items}
audit_data = get_all_audits_for_topic(topic_id)
audit_map = {a["content"]["id"]: a["audit"] for a in audit_data if a["audit"]}

if not generated:
    st.warning("No content generated yet. Go to **🔨 Generate** first.")
    st.stop()

# ── Filter controls ───────────────────────────────────────────────────────────
col1, col2, col3 = st.columns(3)
with col1:
    view_type = st.radio("Content type", ["Video Scripts", "Assessment Questions"], horizontal=True)
with col2:
    filter_status = st.selectbox("Show", ["All", "Needs Review (no audit)", "Failed audit", "Approved"])
with col3:
    filter_uor = st.selectbox("UOR", ["All"] + sorted(set(g["uor_id"] for g in generated)))

ct = "video_script" if view_type == "Video Scripts" else "assessment"
items = [g for g in generated if g["content_type"] == ct]

if filter_uor != "All":
    items = [g for g in items if g["uor_id"] == filter_uor]

if filter_status == "Needs Review (no audit)":
    items = [g for g in items if g["id"] not in audit_map]
elif filter_status == "Failed audit":
    items = [g for g in items if g["id"] in audit_map and not audit_map[g["id"]].get("passed")]
elif filter_status == "Approved":
    reviews = {g["id"]: get_review(g["id"]) for g in items}
    items = [g for g in items if reviews.get(g["id"]) and reviews[g["id"]].get("approved")]

st.markdown(f"**Showing {len(items)} items**")
st.markdown("---")

# ── Review cards ──────────────────────────────────────────────────────────────
for item in items:
    audit = audit_map.get(item["id"])
    review = get_review(item["id"])
    plan_item = plan_map.get(f"{item['uor_id']}_{item['sc_id']}")

    # Header
    if audit:
        cov = audit.get("coverage_score", 0)
        ord_ = audit.get("order_score", 0)
        fid = audit.get("fidelity_score", 0)
        overall = (cov + ord_ + fid) / 3
        score_badge = f"{'✅' if overall >= 80 else '❌'} {overall:.0f}%"
    else:
        score_badge = "⚪ Not audited"

    approved_badge = ""
    if review:
        approved_badge = " | ✅ Approved" if review.get("approved") else " | 🔄 Rejected"

    with st.expander(
        f"**{item['uor_id']} / {item['sc_id']}** — {score_badge}{approved_badge}",
        expanded=(not review)  # Auto-expand un-reviewed items
    ):
        # Show audit flags
        if audit:
            flags = audit.get("flags", [])
            if flags:
                st.markdown("**Audit flags:**")
                for flag in flags:
                    st.caption(f"🔴 [{flag.get('type','')}] {flag.get('detail','')}")
                st.divider()

        # Editable text area
        current_text = review["edited_content"] if review else item["content_text"]
        edited = st.text_area(
            "Script / Question content (edit directly):",
            value=current_text,
            height=400,
            key=f"edit_{item['id']}"
        )

        comments = st.text_input(
            "Comments / notes for this item:",
            value=review.get("comments", "") if review else "",
            key=f"comments_{item['id']}"
        )

        col_a, col_b, col_c = st.columns(3)

        with col_a:
            if st.button("💾 Save & Re-audit", key=f"save_{item['id']}"):
                if not reviewer_name:
                    st.error("Enter your name at the top.")
                else:
                    # Save edited text back to generated_content
                    update_generated_text(item["id"], edited)
                    # Re-run audit on edited content
                    if plan_item:
                        result = audit_content(
                            plan_item, edited,
                            item.get("slide_refs_used", [])
                        )
                        save_audit(
                            item["id"],
                            result["coverage_score"],
                            result["order_score"],
                            result["fidelity_score"],
                            result.get("flags", [])
                        )
                    save_review(item["id"], reviewer_name, edited, False, comments)
                    st.success("Saved and re-audited.")
                    st.rerun()

        with col_b:
            if st.button("✅ Approve", key=f"approve_{item['id']}", type="primary"):
                if not reviewer_name:
                    st.error("Enter your name at the top.")
                else:
                    update_generated_text(item["id"], edited)
                    save_review(item["id"], reviewer_name, edited, True, comments)
                    st.success("Approved!")
                    st.rerun()

        with col_c:
            if st.button("❌ Reject", key=f"reject_{item['id']}"):
                if not reviewer_name:
                    st.error("Enter your name at the top.")
                else:
                    save_review(item["id"], reviewer_name, edited, False,
                                f"REJECTED: {comments}")
                    st.warning("Marked as rejected.")
                    st.rerun()

# ── Mark module approved ──────────────────────────────────────────────────────
st.markdown("---")
all_reviews = [get_review(g["id"]) for g in generated]
all_approved = all(r and r.get("approved") for r in all_reviews)

if all_approved:
    st.success("🎉 All content approved! Go to **📥 Export** to download.")
    if st.button("Mark module as Approved", type="primary"):
        update_topic_status(topic_id, "approved")
else:
    pending = sum(1 for r in all_reviews if not r or not r.get("approved"))
    st.info(f"{pending} items still need review/approval.")
