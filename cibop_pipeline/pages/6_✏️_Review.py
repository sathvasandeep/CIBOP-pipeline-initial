"""Review — business owners read, edit inline, approve or reject content."""

import re
import streamlit as st
import pandas as pd


def _nkey(s):
    """Natural sort key: ETLC.2 < ETLC.10."""
    return [int(c) if c.isdigit() else c.lower() for c in re.split(r'(\d+)', str(s))]

from lib.db import (get_generated, get_plan_items, get_all_audits_for_topic,
                    update_generated_text, save_review, get_review,
                    update_topic_status, save_audit)
from lib.auditor import audit_content
from lib.generator import revise_content
from lib.sidebar import render_topic_selector

st.set_page_config(page_title="Review — CIBOP", layout="wide")
render_topic_selector()
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
- Use **✏️ Edit Script (Scene by Scene)** to edit individual scenes comfortably
- Click **🔄 Apply Edits to Table** after editing, then **💾 Save & Re-audit** to persist
- Click **✅ Approve** to mark as final — only Approved scripts are exported
- Click **❌ Reject** to exclude a script from the export
""")

# Persist reviewer name across reruns
_saved_name = st.session_state.get("_reviewer_name", "")
reviewer_name = st.text_input(
    "Your name (for review record)",
    value=_saved_name,
    placeholder="e.g. Priya Sharma",
    key="_reviewer_name_widget"
)
if reviewer_name:
    st.session_state["_reviewer_name"] = reviewer_name


# ══════════════════════════════════════════════════════════════════════════════
# TABLE PARSING
# ══════════════════════════════════════════════════════════════════════════════

def parse_script_table(script_text: str):
    """Parse the markdown table portion of a video script.
    Returns (df, prefix_text, suffix_text)."""
    lines = script_text.split("\n")
    table_lines, before, after = [], [], []
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

    headers = [c.strip() for c in table_lines[0].split("|")[1:-1]]
    rows = []
    for tline in table_lines[1:]:
        cells = [c.strip() for c in tline.split("|")[1:-1]]
        if len(cells) < 2:
            continue
        if all(re.fullmatch(r"[-: ]+", c) for c in cells):
            continue
        while len(cells) < len(headers):
            cells.append("")
        rows.append(cells[:len(headers)])

    if not rows:
        return None, script_text, ""

    df = pd.DataFrame(rows, columns=headers)
    return df, "\n".join(before).strip(), "\n".join(after).strip()


# ══════════════════════════════════════════════════════════════════════════════
# TABLE VIEWER
# ══════════════════════════════════════════════════════════════════════════════

def render_script_view(script_text: str):
    """Render a video script with the markdown table as a wrapped HTML table."""
    df, prefix, _ = parse_script_table(script_text)
    if prefix:
        st.caption(prefix)

    if df is not None:
        col_widths = {}
        for col in df.columns:
            cl = col.lower()
            if "#" in cl:                                         col_widths[col] = "3%"
            elif "character" in cl:                               col_widths[col] = "8%"
            elif "visual" in cl or "cue" in cl or "animation" in cl: col_widths[col] = "22%"
            elif "on-screen" in cl or "screen" in cl:            col_widths[col] = "20%"
            elif "voice" in cl or "over" in cl:                  col_widths[col] = "47%"
            else:                                                 col_widths[col] = "15%"

        char_colours = {
            "MOTION": ("#1E3A5F", "#E8F0FA"),
            "RYO":    ("#005F5F", "#E8FAF5"),
            "ARIA":   ("#6B3A8B", "#F5EAF5"),
            "BOTH":   ("#8B4513", "#FAF0E8"),
        }

        header_cells = "".join(
            f'<th style="width:{col_widths.get(c,"15%")};padding:8px 10px;'
            f'background:#1E3A5F;color:white;font-size:12px;text-align:left;'
            f'white-space:nowrap;">{c}</th>'
            for c in df.columns
        )
        rows_html = ""
        for _, row in df.iterrows():
            char_val = str(row.get("Character", "")).strip().upper()
            fg, bg = char_colours.get(char_val, ("#333", "#FAFAFA"))
            row_cells = ""
            for col in df.columns:
                cell_val = str(row[col]).replace("\n", "<br>").replace("/", "<br>")
                cl = col.lower()
                is_char = "character" in cl
                style = (
                    "padding:8px 10px;vertical-align:top;font-size:12px;"
                    "word-wrap:break-word;white-space:normal;border-bottom:1px solid #E0E0E0;"
                )
                if is_char:
                    style += f"font-weight:bold;color:{fg};background:{bg};"
                row_cells += f'<td style="{style}">{cell_val}</td>'
            rows_html += f"<tr>{row_cells}</tr>"

        html = f"""
<div style="overflow-x:auto;border-radius:8px;border:1px solid #DDD;margin-bottom:8px;">
<table style="width:100%;border-collapse:collapse;table-layout:fixed;font-family:Arial,sans-serif;">
  <thead><tr>{header_cells}</tr></thead>
  <tbody>{rows_html}</tbody>
</table>
</div>"""
        st.html(html)
    else:
        st.text(script_text)


# ══════════════════════════════════════════════════════════════════════════════
# SCENE EDITOR
# ══════════════════════════════════════════════════════════════════════════════

CHAR_OPTIONS = ["MOTION", "RYO", "ARIA", "BOTH"]
CHAR_ICONS   = {"MOTION": "🔵", "RYO": "🟢", "ARIA": "🟣", "BOTH": "🟠"}


def _col_role(col: str) -> str:
    cl = col.lower()
    if "#" in cl:                                          return "num"
    if "character" in cl:                                  return "char"
    if "visual" in cl or "cue" in cl or "animation" in cl: return "vc"
    if "on-screen" in cl or "screen" in cl:               return "ost"
    if "voice" in cl or "over" in cl:                     return "vo"
    return "other"


def render_scene_editor(item_id: str, df):
    """Show a scene-by-scene form. Edits are stored in scene_* session state keys."""
    if df is None:
        st.warning("Cannot parse table for scene editing.")
        return

    cols = list(df.columns)
    num_col = next((c for c in cols if _col_role(c) == "num"), None)

    for r_idx, row in df.iterrows():
        char_val = str(row.get("Character", "")).strip().upper()
        scene_n  = str(row[num_col]).strip() if num_col and num_col in row else str(r_idx + 1)
        icon = CHAR_ICONS.get(char_val, "⚪")

        with st.expander(f"{icon} Scene {scene_n} — {char_val}", expanded=False):
            left, right = st.columns([1, 2])

            for col in cols:
                role = _col_role(col)
                key  = f"scene_{item_id}_{r_idx}_{col}"
                raw_val = str(row[col])

                if role == "num":
                    continue

                elif role == "char":
                    idx = CHAR_OPTIONS.index(char_val) if char_val in CHAR_OPTIONS else 0
                    with left:
                        st.selectbox("Character", CHAR_OPTIONS, index=idx, key=key)

                elif role == "ost":
                    # Show / as newlines for comfortable editing
                    display = re.sub(r'\s*/\s*', '\n', raw_val)
                    with left:
                        st.text_area(col, value=display, height=160, key=key,
                                     help="Each line = one on-screen line (saved with / separators)")

                elif role == "vc":
                    with right:
                        st.text_area(col, value=raw_val, height=130, key=key,
                                     help="Full 2-4 sentence production brief for this scene")

                elif role == "vo":
                    with right:
                        st.text_area(col, value=raw_val, height=160, key=key,
                                     help="2-4 sentences. Rich, precise financial explanation.")

                else:
                    with left:
                        st.text_area(col, value=raw_val, height=80, key=key)


def _compile_from_scenes(item_id: str, df) -> str:
    """Compile scene_ session state back into a markdown table string."""
    if df is None:
        return ""
    cols = list(df.columns)
    header = "| " + " | ".join(cols) + " |"
    sep    = "| " + " | ".join(["---"] * len(cols)) + " |"
    rows   = []

    for r_idx, orig_row in df.iterrows():
        cells = []
        for col in cols:
            key = f"scene_{item_id}_{r_idx}_{col}"
            if key in st.session_state:
                val = str(st.session_state[key])
                if _col_role(col) == "ost":
                    # Convert newlines back to " / "
                    val = " / ".join(ln.strip() for ln in val.splitlines() if ln.strip())
            else:
                val = str(orig_row[col])
            val = val.replace("|", "\\|")  # escape pipes
            cells.append(val)
        rows.append("| " + " | ".join(cells) + " |")

    return "\n".join([header, sep] + rows)


def _clear_scene_editor(item_id: str, df):
    """Remove scene_ session state keys for this item."""
    if df is None:
        return
    for r_idx in range(len(df)):
        for col in df.columns:
            st.session_state.pop(f"scene_{item_id}_{r_idx}_{col}", None)


# ══════════════════════════════════════════════════════════════════════════════
# LOAD DATA
# ══════════════════════════════════════════════════════════════════════════════

generated  = get_generated(topic_id)
plan_items = get_plan_items(topic_id)
plan_map   = {f"{i['uor_id']}_{i['sc_id']}": i for i in plan_items}
audit_data = get_all_audits_for_topic(topic_id)
audit_map  = {a["content"]["id"]: a["audit"] for a in audit_data if a["audit"]}

if not generated:
    st.warning("No content generated yet. Go to **🔨 Generate** first.")
    st.stop()

# ── Filter controls ────────────────────────────────────────────────────────────
col1, col2, col3 = st.columns(3)
with col1:
    view_type = st.radio("Content type", ["Video Scripts", "Assessment Questions"], horizontal=True)
with col2:
    filter_status = st.selectbox("Show", ["All", "Needs Review (no audit)", "Failed audit", "Approved"])
with col3:
    filter_uor = st.selectbox("UOR", ["All"] + sorted(set(g["uor_id"] for g in generated), key=_nkey))

ct    = "video_script" if view_type == "Video Scripts" else "assessment"
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

# ══════════════════════════════════════════════════════════════════════════════
# REVIEW CARDS
# ══════════════════════════════════════════════════════════════════════════════

for item in sorted(items, key=lambda x: (_nkey(x["uor_id"]), _nkey(x["sc_id"]))):
    audit     = audit_map.get(item["id"])
    review    = get_review(item["id"])
    plan_item = plan_map.get(f"{item['uor_id']}_{item['sc_id']}")

    # Score badge
    if audit:
        cov  = audit.get("coverage_score", 0)
        ord_ = audit.get("order_score", 0)
        fid  = audit.get("fidelity_score", 0)
        overall = (cov + ord_ + fid) / 3
        score_badge = f"{'✅' if overall >= 80 else '❌'} {overall:.0f}%"
    else:
        score_badge = "⚪ Not audited"

    approved_badge = ""
    if review:
        approved_badge = " | ✅ Approved" if review.get("approved") else " | ❌ Rejected"

    # ── Versioned text-area key ────────────────────────────────────────────────
    # Each action (AI edit, Save, Apply Edits) increments _ver_{id}.
    # The new version number produces a brand-new widget key that has no
    # stale frontend state, so the text area always re-initialises from
    # value=current_text (which is always correct from the DB after any save).
    _ver     = st.session_state.get(f"_ver_{item['id']}", 0)
    _ta_key  = f"edit_{item['id']}_v{_ver}"
    # Clean up any leftover _pending_ keys from the old pattern
    st.session_state.pop(f"_pending_{item['id']}", None)

    with st.expander(
        f"**{item['uor_id']} / {item['sc_id']}** — {score_badge}{approved_badge}",
        expanded=(not review)
    ):
        # ── Audit flags ────────────────────────────────────────────────────────
        if audit:
            flags = audit.get("flags", [])
            if flags:
                st.markdown("**Audit flags:**")
                for flag in flags:
                    st.caption(f"🔴 [{flag.get('type','')}] {flag.get('detail','')}")
                st.divider()

        # Current content (prefer saved review edit)
        current_text = review["edited_content"] if review else item["content_text"]

        # ── SLIDE REFERENCES ──────────────────────────────────────────────────
        refs_db = item.get("slide_refs_used") or []
        if refs_db:
            st.caption(f"📎 Slides referenced: {', '.join(str(r) for r in refs_db)}")
        elif plan_item:
            s = plan_item.get("slide_range_start", "")
            e = plan_item.get("slide_range_end", "")
            if s and e:
                st.caption(f"📖 Slide range: {s}–{e}")

        # ── SCRIPT TABLE (video) or plain text (assessment) ───────────────────
        if ct == "video_script":
            st.markdown("##### 📋 Script Table")
            render_script_view(current_text)

            # Parse for scene editor
            df_edit, _, _ = parse_script_table(current_text)

            # ── SCENE EDITOR ──────────────────────────────────────────────────
            with st.expander("✏️ Edit Script — Scene by Scene", expanded=False):
                if df_edit is not None:
                    st.caption(
                        "Edit each scene below. Click **🔄 Apply Edits to Table** when done, "
                        "then **💾 Save & Re-audit** or **✅ Approve**."
                    )
                    render_scene_editor(item["id"], df_edit)
                    st.markdown("---")
                    btn_apply, btn_reset = st.columns([3, 1])
                    with btn_apply:
                        if st.button("🔄 Apply Edits to Table", key=f"apply_scenes_{item['id']}",
                                     type="primary"):
                            compiled = _compile_from_scenes(item["id"], df_edit)
                            # Increment version and pre-set the new text area key
                            # with the compiled markdown so it appears immediately.
                            _new_ver = _ver + 1
                            _new_ta_key = f"edit_{item['id']}_v{_new_ver}"
                            st.session_state[f"_ver_{item['id']}"] = _new_ver
                            st.session_state[_new_ta_key] = compiled
                            _clear_scene_editor(item["id"], df_edit)
                            st.success("✅ Scene edits applied — now Save or Approve below.")
                            st.rerun()
                    with btn_reset:
                        if st.button("↩️ Reset", key=f"reset_scenes_{item['id']}",
                                     help="Discard all scene edits"):
                            _clear_scene_editor(item["id"], df_edit)
                            st.rerun()
                else:
                    st.warning("Could not parse table for scene editing.")

            # ── RAW MARKDOWN EDITOR (Advanced) ────────────────────────────────
            with st.expander("🔧 Advanced — Raw Markdown Table", expanded=False):
                st.caption("Direct markdown edit — useful for precise formatting fixes.")
                st.text_area(
                    "Raw markdown table:",
                    value=current_text,
                    height=350,
                    key=_ta_key
                )

        else:
            # Assessment — plain text
            st.text_area(
                "Assessment question (edit directly):",
                value=current_text,
                height=300,
                key=_ta_key
            )

        st.markdown("---")

        # ── COMMENTS ──────────────────────────────────────────────────────────
        comments = st.text_area(
            "💬 Comments / change requests:",
            value=review.get("comments", "") if review else "",
            height=100,
            placeholder="e.g. Scene 4 Voice Over too long. Add IRS worked example in scene 4.",
            key=f"comments_{item['id']}"
        )

        col_ai, col_save, col_approve, col_reject = st.columns([2, 2, 2, 1])

        # ── AI EDITS ──────────────────────────────────────────────────────────
        with col_ai:
            if st.button("🤖 Apply AI Edits", key=f"ai_{item['id']}",
                         help="Claude revises the script based on your comments"):
                if not reviewer_name:
                    st.error("Enter your name at the top first.")
                elif not comments.strip():
                    st.warning("Enter comments above first.")
                elif not plan_item:
                    st.error("Plan item not found.")
                else:
                    with st.spinner("🤖 Claude is revising…"):
                        try:
                            source = st.session_state.get(_ta_key, current_text)
                            revised, new_refs = revise_content(plan_item, source, comments, ct)
                            update_generated_text(item["id"], revised)
                            save_review(item["id"], reviewer_name, revised, False,
                                        f"AI revision: {comments}")
                            result = audit_content(plan_item, revised, new_refs)
                            save_audit(item["id"], result["coverage_score"],
                                       result["order_score"], result["fidelity_score"],
                                       result.get("flags", []))
                            # Increment version so the next render creates a fresh text
                            # area widget that re-initialises from value=current_text
                            # (which will be the revised text, just saved to DB).
                            st.session_state[f"_ver_{item['id']}"] = _ver + 1
                            df_tmp, _, _ = parse_script_table(revised)
                            _clear_scene_editor(item["id"], df_tmp)
                            st.success("✅ AI revision applied and re-audited.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"AI revision failed: {e}")

        # ── SAVE & RE-AUDIT ───────────────────────────────────────────────────
        with col_save:
            if st.button("💾 Save & Re-audit", key=f"save_{item['id']}"):
                if not reviewer_name:
                    st.error("Enter your name at the top.")
                else:
                    raw = st.session_state.get(_ta_key, current_text)
                    update_generated_text(item["id"], raw)
                    if plan_item:
                        result = audit_content(plan_item, raw, item.get("slide_refs_used", []))
                        save_audit(item["id"], result["coverage_score"],
                                   result["order_score"], result["fidelity_score"],
                                   result.get("flags", []))
                    save_review(item["id"], reviewer_name, raw, False, comments)
                    df_tmp, _, _ = parse_script_table(raw)
                    _clear_scene_editor(item["id"], df_tmp)
                    # Increment version so next render re-initialises the text area
                    st.session_state[f"_ver_{item['id']}"] = _ver + 1
                    st.success("Saved and re-audited.")
                    st.rerun()

        # ── APPROVE ───────────────────────────────────────────────────────────
        with col_approve:
            if st.button("✅ Approve", key=f"approve_{item['id']}", type="primary"):
                if not reviewer_name:
                    st.error("Enter your name at the top.")
                else:
                    raw = st.session_state.get(_ta_key, current_text)
                    update_generated_text(item["id"], raw)
                    save_review(item["id"], reviewer_name, raw, True, comments)
                    df_tmp, _, _ = parse_script_table(raw)
                    _clear_scene_editor(item["id"], df_tmp)
                    # Increment version so next render re-initialises the text area
                    st.session_state[f"_ver_{item['id']}"] = _ver + 1
                    st.success("✅ Approved!")
                    st.rerun()

        # ── REJECT ────────────────────────────────────────────────────────────
        with col_reject:
            if st.button("❌ Reject", key=f"reject_{item['id']}"):
                if not reviewer_name:
                    st.error("Enter your name at the top.")
                else:
                    raw = st.session_state.get(_ta_key, current_text)
                    save_review(item["id"], reviewer_name, raw, False, f"REJECTED: {comments}")
                    df_tmp, _, _ = parse_script_table(raw)
                    _clear_scene_editor(item["id"], df_tmp)
                    st.warning("Marked as rejected.")
                    st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# MODULE-LEVEL SUMMARY
# ══════════════════════════════════════════════════════════════════════════════

st.markdown("---")
all_reviews   = [get_review(g["id"]) for g in generated]
approved_list = [r for r in all_reviews if r and r.get("approved")]
rejected_list = [r for r in all_reviews if r and not r.get("approved")]
pending_count = len(all_reviews) - len(approved_list) - len(rejected_list)

col_a, col_b, col_c = st.columns(3)
col_a.metric("✅ Approved", len(approved_list))
col_b.metric("❌ Rejected", len(rejected_list))
col_c.metric("⏳ Pending", pending_count)

if len(approved_list) == len(all_reviews):
    st.success("🎉 All content approved! Go to **📥 Export** to download.")
    if st.button("Mark module as Approved", type="primary"):
        update_topic_status(topic_id, "approved")
else:
    st.info(f"{len(approved_list)}/{len(all_reviews)} items approved. "
            f"Go to **📥 Export** to download approved content at any time.")
