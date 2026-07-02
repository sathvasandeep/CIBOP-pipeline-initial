"""Setup — upload PPT, UOR Excel, and optional supplementary PDF."""

import streamlit as st
from lib.db import save_file_record, get_files_for_topic, update_topic_status
from lib.parser import extract_slides_text, parse_uor_excel, extract_pdf_text, get_slide_excerpt
import json

st.set_page_config(page_title="Setup — CIBOP", layout="wide")
st.title("⚙️ Setup")

# ── Guard: topic must be selected ─────────────────────────────────────────────
if "selected_topic_id" not in st.session_state:
    st.warning("Please select a topic first on the **📂 Topics** page.")
    st.stop()

topic_id = st.session_state["selected_topic_id"]
topic_name = st.session_state.get("selected_topic_name", "")
st.subheader(f"Module: {topic_name}")

# ── Show existing files ───────────────────────────────────────────────────────
existing = {f["file_type"]: f for f in get_files_for_topic(topic_id)}

col_a, col_b = st.columns(2)
with col_a:
    if "ppt" in existing:
        st.success(f"✅ PPT uploaded: `{existing['ppt']['file_name']}`")
    if "uor" in existing:
        st.success(f"✅ UOR Excel uploaded: `{existing['uor']['file_name']}`")
    if "pdf" in existing:
        st.info(f"ℹ️ Supplementary PDF: `{existing['pdf']['file_name']}`")

st.markdown("---")

# ── Upload form ───────────────────────────────────────────────────────────────
with st.form("upload_form"):
    st.markdown("### Upload Source Files")
    ppt_file = st.file_uploader("Source PPT file (.pptx)", type=["pptx"])
    uor_file = st.file_uploader("Approved UOR Excel (.xlsx)", type=["xlsx"])
    pdf_file = st.file_uploader("Supplementary PDF (optional — e.g., SSI consequences)", type=["pdf"])
    submit = st.form_submit_button("Process & Save Files")

if submit:
    if not ppt_file and "ppt" not in existing:
        st.error("PPT file is required.")
        st.stop()
    if not uor_file and "uor" not in existing:
        st.error("UOR Excel file is required.")
        st.stop()

    progress = st.progress(0, text="Processing files…")

    # ── Process PPT ───────────────────────────────────────────────────────────
    if ppt_file:
        progress.progress(20, text="Extracting slide text…")
        ppt_bytes = ppt_file.read()
        slides_dict = extract_slides_text(ppt_bytes)
        # Store as JSON string for easy retrieval
        slides_json = json.dumps(slides_dict)
        save_file_record(topic_id, "ppt", ppt_file.name, slides_json)
        st.success(f"✅ PPT processed — {len(slides_dict)} slides extracted")

    # ── Process UOR ───────────────────────────────────────────────────────────
    if uor_file:
        progress.progress(50, text="Parsing UOR Excel…")
        uor_bytes = uor_file.read()
        uors = parse_uor_excel(uor_bytes)
        uors_json = json.dumps(uors)
        save_file_record(topic_id, "uor", uor_file.name, uors_json)
        st.success(f"✅ UOR Excel processed — {len(uors)} UORs found")

        # Show UOR preview
        with st.expander("Preview UORs", expanded=True):
            for u in uors:
                sc_count = len(u.get("sub_competencies", []))
                extra_count = len(u.get("extra_scs", []))
                st.markdown(f"**{u['uor_id']}** — {u['title']} | Slides: {u['slide_nos_source']} | {sc_count} SCs{f' + {extra_count} extras' if extra_count else ''}")

    # ── Process PDF ───────────────────────────────────────────────────────────
    if pdf_file:
        progress.progress(75, text="Extracting PDF text…")
        pdf_bytes = pdf_file.read()
        pdf_text = extract_pdf_text(pdf_bytes)
        save_file_record(topic_id, "pdf", pdf_file.name, pdf_text)
        st.success(f"✅ PDF processed — {len(pdf_text)} characters extracted")

    progress.progress(90, text="Saving…")
    update_topic_status(topic_id, "setup")
    progress.progress(100, text="Done!")
    st.balloons()
    st.success("All files processed. Go to **📋 Plan** to review the content plan.")

# ── Slide preview ─────────────────────────────────────────────────────────────
if "ppt" in existing:
    with st.expander("🔍 Preview Slide Content", expanded=False):
        slides_dict = json.loads(existing["ppt"]["extracted_text"])
        slide_nums = sorted(int(k) for k in slides_dict.keys())
        preview_range = st.select_slider(
            "Slide range to preview",
            options=slide_nums,
            value=(slide_nums[0], min(slide_nums[-1], slide_nums[0] + 4))
        )
        start_s, end_s = preview_range
        for n in range(start_s, end_s + 1):
            text = slides_dict.get(str(n), "")
            if text:
                st.markdown(f"**Slide {n}**")
                st.text(text[:500])
                st.divider()
