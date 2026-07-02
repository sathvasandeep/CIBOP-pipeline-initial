"""Export — download approved content as formatted DOCX files."""

import streamlit as st
import io
from lib.db import get_generated, get_plan_items, get_review, get_all_audits_for_topic

st.set_page_config(page_title="Export — CIBOP", layout="wide")
st.title("📥 Export")

if "selected_topic_id" not in st.session_state:
    st.warning("Please select a topic first.")
    st.stop()

topic_id = st.session_state["selected_topic_id"]
topic_name = st.session_state.get("selected_topic_name", "")
module_code = st.session_state.get("selected_topic_code", "CIBOP")
st.subheader(f"Module: {topic_name}")

generated = get_generated(topic_id)
plan_items = get_plan_items(topic_id)
audit_data = {a["content"]["id"]: a["audit"] for a in get_all_audits_for_topic(topic_id) if a["audit"]}

if not generated:
    st.warning("No content generated yet.")
    st.stop()

# ── Export options ────────────────────────────────────────────────────────────
st.markdown("### Export Options")
include_unreviewed = st.checkbox(
    "Include content not yet reviewed (caution: may have issues)",
    value=False
)
include_failed = st.checkbox(
    "Include content that failed audit (<80%)",
    value=False
)

st.markdown("---")

# ── Build DOCX export ─────────────────────────────────────────────────────────
def build_video_docx(items, plan_map, audit_map, include_unreviewed, include_failed):
    from docx import Document
    from docx.shared import Pt, RGBColor, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    NAVY = RGBColor(0x1E, 0x3A, 0x5F)
    TEAL = RGBColor(0x00, 0x5F, 0x5F)
    GOLD = RGBColor(0x8B, 0x69, 0x14)

    doc = Document()
    # Page setup
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    for margin in ("top_margin", "bottom_margin", "left_margin", "right_margin"):
        setattr(section, margin, Cm(1.5))

    # Title
    title_para = doc.add_paragraph()
    title_run = title_para.add_run(f"{topic_name} — Video Production Document")
    title_run.font.size = Pt(18)
    title_run.font.color.rgb = NAVY
    title_run.font.bold = True
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    current_uor = None
    for item in sorted(items, key=lambda x: (x["uor_id"], x["sc_id"])):
        if item["content_type"] != "video_script":
            continue

        # Apply filters
        review = get_review(item["id"])
        audit = audit_map.get(item["id"])

        if not include_unreviewed and not review:
            continue
        if not include_failed and audit:
            overall = (audit.get("coverage_score", 0) + audit.get("order_score", 0) +
                       audit.get("fidelity_score", 0)) / 3
            if overall < 80:
                continue

        # UOR header
        if item["uor_id"] != current_uor:
            current_uor = item["uor_id"]
            doc.add_page_break()
            plan_item = plan_map.get(f"{item['uor_id']}_{item['sc_id']}")
            uor_title = plan_item.get("uor_title", "") if plan_item else ""

            h = doc.add_paragraph()
            r = h.add_run(f"UOR {item['uor_id']}: {uor_title}")
            r.font.size = Pt(14)
            r.font.bold = True
            r.font.color.rgb = NAVY

        # SC header
        sc_para = doc.add_paragraph()
        sc_run = sc_para.add_run(f"🎬 {item['sc_id']} — Video Script")
        sc_run.font.size = Pt(12)
        sc_run.font.bold = True
        sc_run.font.color.rgb = TEAL

        # Audit score note
        if audit:
            overall = (audit.get("coverage_score", 0) + audit.get("order_score", 0) +
                       audit.get("fidelity_score", 0)) / 3
            score_para = doc.add_paragraph()
            score_run = score_para.add_run(
                f"Audit: {overall:.0f}% | Coverage {audit.get('coverage_score',0):.0f}% | "
                f"Order {audit.get('order_score',0):.0f}% | "
                f"Fidelity {audit.get('fidelity_score',0):.0f}%"
            )
            score_run.font.size = Pt(9)
            score_run.font.italic = True

        # Content
        content = review["edited_content"] if review else item["content_text"]
        for line in content.split("\n"):
            if line.strip():
                p = doc.add_paragraph(line.strip())
                p.style.font.size = Pt(10)

        # Divider
        doc.add_paragraph("─" * 60)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


def build_assessment_docx(items, plan_map, audit_map, include_unreviewed, include_failed):
    from docx import Document
    from docx.shared import Pt, RGBColor, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    NAVY = RGBColor(0x1E, 0x3A, 0x5F)
    TEAL = RGBColor(0x00, 0x5F, 0x5F)

    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    for margin in ("top_margin", "bottom_margin", "left_margin", "right_margin"):
        setattr(section, margin, Cm(1.5))

    title_para = doc.add_paragraph()
    title_run = title_para.add_run(f"{topic_name} — Knowledge Assessment")
    title_run.font.size = Pt(18)
    title_run.font.color.rgb = NAVY
    title_run.font.bold = True
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    q_num = 1
    current_uor = None

    for item in sorted(items, key=lambda x: (x["uor_id"], x["sc_id"])):
        if item["content_type"] != "assessment":
            continue

        review = get_review(item["id"])
        audit = audit_map.get(item["id"])

        if not include_unreviewed and not review:
            continue
        if not include_failed and audit:
            overall = (audit.get("coverage_score", 0) + audit.get("order_score", 0) +
                       audit.get("fidelity_score", 0)) / 3
            if overall < 80:
                continue

        if item["uor_id"] != current_uor:
            current_uor = item["uor_id"]
            doc.add_paragraph()
            plan_item = plan_map.get(f"{item['uor_id']}_{item['sc_id']}")
            uor_title = plan_item.get("uor_title", "") if plan_item else ""
            h = doc.add_paragraph()
            r = h.add_run(f"UOR {item['uor_id']}: {uor_title}")
            r.font.size = Pt(12)
            r.font.bold = True
            r.font.color.rgb = NAVY

        # Question header
        qh = doc.add_paragraph()
        qr = qh.add_run(f"Q{q_num} | {item['uor_id']} / {item['sc_id']}")
        qr.font.size = Pt(11)
        qr.font.bold = True
        qr.font.color.rgb = TEAL
        q_num += 1

        content = review["edited_content"] if review else item["content_text"]
        for line in content.split("\n"):
            if line.strip():
                p = doc.add_paragraph(line.strip())
                p.style.font.size = Pt(10)

        doc.add_paragraph()

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# ── Export buttons ────────────────────────────────────────────────────────────
plan_map = {f"{i['uor_id']}_{i['sc_id']}": i for i in plan_items}

col1, col2 = st.columns(2)

with col1:
    st.markdown("#### 📹 Video Production Document")
    video_items = [g for g in generated if g["content_type"] == "video_script"]
    st.caption(f"{len(video_items)} video scripts")
    if st.button("Build Video DOCX", type="primary"):
        with st.spinner("Building document…"):
            try:
                buf = build_video_docx(
                    generated, plan_map, audit_data,
                    include_unreviewed, include_failed
                )
                st.download_button(
                    label="⬇️ Download Video Production DOCX",
                    data=buf,
                    file_name=f"{topic_name.replace(' ', '_')}_Video_Production.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"Build failed: {e}")

with col2:
    st.markdown("#### 📝 Knowledge Assessment")
    assess_items = [g for g in generated if g["content_type"] == "assessment"]
    st.caption(f"{len(assess_items)} assessment questions")
    if st.button("Build Assessment DOCX", type="primary"):
        with st.spinner("Building document…"):
            try:
                buf = build_assessment_docx(
                    generated, plan_map, audit_data,
                    include_unreviewed, include_failed
                )
                st.download_button(
                    label="⬇️ Download Assessment DOCX",
                    data=buf,
                    file_name=f"{topic_name.replace(' ', '_')}_Assessment.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"Build failed: {e}")

# ── Status summary ────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown("### Content Status Summary")

data = []
for g in generated:
    review = get_review(g["id"])
    audit = audit_data.get(g["id"])
    overall = None
    if audit:
        overall = (audit.get("coverage_score", 0) + audit.get("order_score", 0) +
                   audit.get("fidelity_score", 0)) / 3

    data.append({
        "UOR": g["uor_id"],
        "SC": g["sc_id"],
        "Type": "📹 Video" if g["content_type"] == "video_script" else "📝 Assessment",
        "Audit Score": f"{overall:.0f}%" if overall is not None else "—",
        "Pass": "✅" if overall and overall >= 80 else ("❌" if overall else "—"),
        "Review": "✅ Approved" if (review and review.get("approved")) else
                  ("❌ Rejected" if (review and not review.get("approved")) else "⏳ Pending")
    })

if data:
    import pandas as pd
    st.dataframe(pd.DataFrame(data), use_container_width=True, hide_index=True)
