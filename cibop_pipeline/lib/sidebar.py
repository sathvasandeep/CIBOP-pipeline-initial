"""Shared sidebar components — call render_topic_selector() at the top of every page."""

import streamlit as st


def render_topic_selector():
    """Render a persistent topic selector in the sidebar.

    Loads the topic list, shows a selectbox pre-selected to the current topic,
    and updates session_state if the user switches. Safe to call on every page.
    """
    from lib.db import list_topics

    with st.sidebar:
        st.markdown("---")
        st.markdown("**📚 Active Topic**")

        try:
            topics = list_topics()
        except Exception:
            st.sidebar.caption("⚠️ Could not load topics")
            return

        if not topics:
            st.caption("No topics yet — create one on the **📂 Topics** page.")
            return

        # Build label → topic map
        labels = [f"{t['name']}  ({t.get('module_code', '')})" for t in topics]
        topic_by_label = dict(zip(labels, topics))

        # Find current selection index
        current_id = st.session_state.get("selected_topic_id")
        default_idx = 0
        for i, t in enumerate(topics):
            if t["id"] == current_id:
                default_idx = i
                break

        chosen_label = st.selectbox(
            "Switch topic",
            labels,
            index=default_idx,
            key="_global_topic_selector",
            label_visibility="collapsed",
        )

        chosen = topic_by_label[chosen_label]

        # Update session state if topic changed (or not set yet)
        if chosen["id"] != st.session_state.get("selected_topic_id"):
            st.session_state["selected_topic_id"]   = chosen["id"]
            st.session_state["selected_topic_name"] = chosen["name"]
            st.session_state["selected_topic_code"] = chosen.get("module_code", "")
            st.rerun()
        else:
            # Ensure all keys are populated even on first load
            st.session_state.setdefault("selected_topic_id",   chosen["id"])
            st.session_state.setdefault("selected_topic_name", chosen["name"])
            st.session_state.setdefault("selected_topic_code", chosen.get("module_code", ""))

        # Show status badge
        status = chosen.get("status", "setup")
        STATUS_ICONS = {
            "setup": "⚙️", "planned": "📋", "generated": "🔨",
            "audited": "🔍", "approved": "✅"
        }
        st.caption(f"{STATUS_ICONS.get(status, '⚙️')} {status.title()}")
        st.markdown("---")
