"""Topics — list all modules, create new ones, select to work on."""

import streamlit as st
from lib.db import list_topics, create_topic, delete_topic

st.set_page_config(page_title="Topics — CIBOP", layout="wide")
st.title("📂 Topics")
st.caption("Each topic = one training module with its own PPT and UOR file.")

# ── Create new topic ──────────────────────────────────────────────────────────
with st.expander("➕ Create New Topic", expanded=False):
    with st.form("new_topic"):
        col1, col2 = st.columns(2)
        with col1:
            name = st.text_input("Module Name", placeholder="Exchange Trade Lifecycle")
        with col2:
            code = st.text_input("Module Code", placeholder="ETLC")
        submitted = st.form_submit_button("Create Topic")
        if submitted:
            if not name or not code:
                st.error("Please fill in both fields.")
            else:
                topic = create_topic(name.strip(), code.strip().upper())
                st.success(f"Created: **{name}** ({code})")
                st.session_state["selected_topic_id"] = topic["id"]
                st.session_state["selected_topic_name"] = name
                st.rerun()

st.markdown("---")

# ── List existing topics ──────────────────────────────────────────────────────
topics = list_topics()

if not topics:
    st.info("No topics yet. Create your first one above.")
else:
    STATUS_ICONS = {
        "setup": "⚙️",
        "planned": "📋",
        "generated": "🔨",
        "audited": "🔍",
        "approved": "✅"
    }

    for topic in topics:
        icon = STATUS_ICONS.get(topic.get("status", "setup"), "⚙️")
        col1, col2, col3, col4 = st.columns([3, 1, 1, 1])
        with col1:
            st.markdown(f"**{icon} {topic['name']}** &nbsp; `{topic.get('module_code','')}`")
            st.caption(f"Status: {topic.get('status','setup').title()} | Created: {topic['created_at'][:10]}")
        with col2:
            if st.button("Select", key=f"sel_{topic['id']}"):
                st.session_state["selected_topic_id"] = topic["id"]
                st.session_state["selected_topic_name"] = topic["name"]
                st.success(f"Working on: {topic['name']}")
        with col3:
            pass
        with col4:
            if st.button("🗑️", key=f"del_{topic['id']}", help="Delete this topic"):
                delete_topic(topic["id"])
                st.rerun()
        st.divider()

# ── Show currently selected ───────────────────────────────────────────────────
if "selected_topic_id" in st.session_state:
    st.success(f"✅ Currently working on: **{st.session_state.get('selected_topic_name', '')}**")
    st.markdown("→ Go to **⚙️ Setup** to upload files, or jump to any step in the sidebar.")
