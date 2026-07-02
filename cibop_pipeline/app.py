"""CIBOP Content Pipeline — Main entry point."""

import streamlit as st

st.set_page_config(
    page_title="CIBOP Content Pipeline",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── Sidebar branding ──────────────────────────────────────────────────────────
with st.sidebar:
    st.image("https://via.placeholder.com/200x60/1E3A5F/FFFFFF?text=CIBOP", width=200)
    st.markdown("## Content Pipeline")
    st.markdown("---")
    st.markdown("""
**Workflow:**
1. 📂 Topics
2. ⚙️ Setup
3. 📋 Plan
4. 🔨 Generate
5. 🔍 Audit
6. ✏️ Review
7. 📥 Export
""")
    st.markdown("---")
    st.caption("Built for CIBOP Programme")

# ── Main landing page ─────────────────────────────────────────────────────────
st.title("📚 CIBOP Content Pipeline")
st.subheader("Plan → Generate → Audit → Review → Export")

st.markdown("""
Welcome to the CIBOP training content pipeline. This tool ensures every video script
and assessment question is **strictly grounded** in the approved UOR framework and source PPT slides.

---
### How it works

**Step 1 — Create a Topic**
Name the module (e.g., "Exchange Trade Lifecycle") and upload the source PPT and UOR Excel file.

**Step 2 — Review the Content Plan**
The planning agent reads the UOR file and PPT, maps each sub-competency to exact slide ranges and key terms. You approve the plan before any content is generated.

**Step 3 — Generate**
Content is generated *one sub-competency at a time*, with only that SC's slide text provided to the AI. This prevents cross-contamination and ensures slide order is maintained.

**Step 4 — Audit**
An independent audit agent scores every piece of content on three dimensions:
- **Coverage** — does it address the sub-competency and use the key terms?
- **Order** — are slides referenced in ascending order?
- **Fidelity** — is every fact traceable to the source slides? (80% threshold to pass)

**Step 5 — Review**
Business owners can read, edit inline, and approve/reject each script. Rejected items go back to generation with the reviewer's comments.

**Step 6 — Export**
Download the final approved Video Production DOCX and Assessment DOCX, ready for the vendor.

---
""")

col1, col2, col3 = st.columns(3)
with col1:
    st.info("👈 Use the sidebar to navigate between steps")
with col2:
    st.success("🔒 All content is locked to approved slide ranges")
with col3:
    st.warning("⚠️ Content below 80% audit score cannot be exported")

st.markdown("---")
st.markdown("**→ Start by going to [Topics](Topics) in the sidebar**")
