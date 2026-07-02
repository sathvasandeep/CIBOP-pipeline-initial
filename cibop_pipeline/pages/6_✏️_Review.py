"""Review — business owners read, edit inline, approve or reject content."""

import re
import streamlit as st
import pandas as pd
from lib.db import (get_generated, get_plan_items, get_all_audits_for_topic,
                    update_generated_text, save_review, get_review,
                    update_topic_status, save_audit)
from lib.auditor import audit_content
from lib.generator import revise_content

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
- Read each script in the table view below
- Enter comments in the **Comments** box if changes are needed
- Click **🤖 Apply AI Edits** to have Claude revise the script based on your comments
- Click **💾 Save & Re-audit** after any changes to re-score the content
- Click **✅ Approve** to mark as final, or **❌ Reject** to send back
- Only approved content will appear in the Export
""")

reviewer_name = st.text_input("Your name (for review record)", placeholder="e.g. Priya Sharma")

# ── Helpers ───────────────────────────────────────────────────────────────────

def parse_script_table(script_text: str):
    """Parse the markdown table portion of a video script into a DataFrame.
    Returns (df, prefix_text, suffix_text) where prefix/suffix are non-table lines."""
    lines = script_text.split("\n")
    table_lines = []
    before = []
    after = []
    in_table = False

    for line in lines:
        stripped = line.strip()
        if stripped.startswith("|"):
            in_table = True
            table_lines.append(stripped)
        elif in_table:
            after.append(line)
        else:
            before.append(line)

    if not table_lines:
        return None, script_text, ""

    # Parse header
    header_raw = table_lines[0]
    headers = [c.strip() for c in header_raw.split("|")[1:-1]]

    rows = []
    for tline in table_lines[1:]:
        cells = [c.strip() for c in tline.split("|")[1:-1]]
        if len(cells) < 2:
            continue
        # Skip separator rows (all dashes)
        if all(re.fullmatch(r"[-: ]+", c) for c in cells):
            continue
        if len(cells) == len(headers):
            rows.append(cells)
        else:
            # Pad if needed
            while len(cells) < len(headers):
                cells.append("")
            rows.append(cells[:len(headers)])

    if not rows:
        return None, script_text, ""

    df = pd.DataFrame(rows, columns=headers)
    prefix = "\n".join(before).strip()
    suffix = "\n".join(after).strip()
    return df, prefix, suffix


def render_script_view(script_text: str):
    """Render a video script with the markdown table as a visual table."""
    df, prefix, suffix = parse_script_table(script_text)

    if prefix:
        st.caption(prefix)

    if df is not None:
        # Build column config with appropriate widths
        col_config = {}
        for col in df.columns:
            col_lower = col.lower()
            if "#" in col_lower:
                col_config[col] = st.column_config.TextColumn(col, width="small")
            elif "character" in col_lower:
                col_config[col] = st.column_config.TextColumn(col, width="small")
            elif "visual" in col_lower or "cue" in col_lower:
                col_config[col] = st.column_config.TextColumn(col, width="medium")
            elif "on-screen" in col_lower or "text" in col_lower:
                col_config[col] = st.column_config.TextColumn(col, width="medium")
            elif "voice" in col_lower or "over" in col_lower:
                col_config[col] = st.column_config.TextColumn(col, width="large")

        st.dataframe(
            df,
            use_container_width=True,
            hide_index=True,
            column_config=col_config,
        )
    else:
        st.text(script_text)

    if suffix:
        # Show SLIDE_REFS line cleanly
        refs_match = re.search(r'SLIDE_REFS_USED:\s*(.+)', suffix)
        if refs_match:
            st.caption(f"📎 Slides referenced: {refs_match.group(1)}")


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

    # Score badge
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
        expanded=(not review)
    ):
        # ── Audit flags ───────────────────────────────────────────────────────
        if audit:
            flags = audit.get("flags", [])
            if flags:
                st.markdown("**Audit flags:**")
                for flag in flags:
                    st.caption(f"🔴 [{flag.get('type','')}] {flag.get('detail','')}")
                st.divider()

        # Current content (prefer last saved review edit, else original)
        current_text = review["edited_content"] if review else item["content_text"]

        # ── TABLE VIEW (video scripts) / TEXT VIEW (assessments) ─────────────
        if ct == "video_script":
            st.markdown("##### 📋 Script Table")
            render_script_view(current_text)

            # Raw editor in expander
            with st.expander("✏️ Edit raw script (advanced)", expanded=False):
                edited = st.text_area(
                    "Raw markdown table — edit directly if needed:",
                    value=current_text,
                    height=400,
                    key=f"edit_{item['id']}"
                )
        else:
            # Assessment — plain text editor
            edited = st.text_area(
                "Assessment question content (edit directly):",
                value=current_text,
                height=300,
                key=f"edit_{item['id']}"
            )

        st.markdown("---")

        # ── Comments + AI revision ────────────────────────────────────────────
        comments = st.text_area(
            "💬 Comments / change requests for this item:",
            value=review.get("comments", "") if review else "",
            height=100,
            placeholder=(
                "e.g. Scene 4 voice-over is too long. Shorten to 2 sentences.\n"
                "Remove the analogy in scene 5 — use exact slide wording instead."
            ),
            key=f"comments_{item['id']}"
        )

        col_ai, col_save, col_approve, col_reject = st.columns([2, 2, 2, 1])

        with col_ai:
            ai_btn = st.button(
                "🤖 Apply AI Edits",
                key=f"ai_{item['id']}",
                help="Claude will revise the script based on your comments above, using only slide content."
            )
            if ai_btn:
                if not reviewer_name:
                    st.error("Enter your name at the top first.")
                elif not comments.strip():
                    st.warning("Enter some comments above describing what to change.")
                elif not plan_item:
                    st.error("Plan item not found — cannot run revision.")
                else:
                    with st.spinner("🤖 Claude is revising the script…"):
                        try:
                            # Use edited raw text if the user changed it, else current_text
                            source_text = st.session_state.get(f"edit_{item['id']}", current_text)
                            revised_text, new_refs = revise_content(
                                plan_item, source_text, comments, ct
                            )
                            # Persist to DB
                            update_generated_text(item["id"], revised_text)
                            # Save as a pending review (not yet approved)
                            save_review(item["id"], reviewer_name, revised_text, False,
                                        f"AI revision applied. Original comments: {comments}")
                            # Re-audit
                            result = audit_content(plan_item, revised_text, new_refs)
                            save_audit(
                                item["id"],
                                result["coverage_score"],
                                result["order_score"],
                                result["fidelity_score"],
                                result.get("flags", [])
                            )
                            st.success("✅ AI revision applied and re-audited. Review the updated table above.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"AI revision failed: {e}")

        with col_save:
            if st.button("💾 Save & Re-audit", key=f"save_{item['id']}"):
                if not reviewer_name:
                    st.error("Enter your name at the top.")
                else:
                    raw_edited = st.session_state.get(f"edit_{item['id']}", current_text)
                    update_generated_text(item["id"], raw_edited)
                    if plan_item:
                        result = audit_content(
                            plan_item, raw_edited,
                            item.get("slide_refs_used", [])
                        )
                        save_audit(
                            item["id"],
                            result["coverage_score"],
                            result["order_score"],
                            result["fidelity_score"],
                            result.get("flags", [])
                        )
                    save_review(item["id"], reviewer_name, raw_edited, False, comments)
                    st.success("Saved and re-audited.")
                    st.rerun()

        with col_approve:
            if st.button("✅ Approve", key=f"approve_{item['id']}", type="primary"):
                if not reviewer_name:
                    st.error("Enter your name at the top.")
                else:
                    raw_edited = st.session_state.get(f"edit_{item['id']}", current_text)
                    update_generated_text(item["id"], raw_edited)
                    save_review(item["id"], reviewer_name, raw_edited, True, comments)
                    st.success("Approved!")
                    st.rerun()

        with col_reject:
            if st.button("❌ Reject", key=f"reject_{item['id']}"):
                if not reviewer_name:
                    st.error("Enter your name at the top.")
                else:
                    raw_edited = st.session_state.get(f"edit_{item['id']}", current_text)
                    save_review(item["id"], reviewer_name, raw_edited, False,
                                f"REJECTED: {comments}")
                    st.warning("Marked as rejected.")
                    st.rerun()

# ── Module-level approval ─────────────────────────────────────────────────────
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
