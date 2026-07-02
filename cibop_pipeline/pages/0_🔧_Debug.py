"""Temporary debug page — delete after fixing connection issues."""
import streamlit as st
import traceback

st.title("🔧 Connection Debug")

st.markdown("### 1. Secrets check")
try:
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    api = st.secrets["ANTHROPIC_API_KEY"]
    st.success(f"✅ SUPABASE_URL = `{url}`")
    st.success(f"✅ SUPABASE_KEY = `{key[:20]}...`")
    st.success(f"✅ ANTHROPIC_API_KEY = `{api[:15]}...`")
except Exception as e:
    st.error(f"❌ Secrets error: {e}")
    st.stop()

st.markdown("### 2. Raw HTTP test (bypass supabase client)")
import httpx
try:
    resp = httpx.get(f"{url}/rest/v1/topics?select=id&limit=1",
                     headers={"apikey": key, "Authorization": f"Bearer {key}"},
                     timeout=10)
    st.success(f"✅ HTTP response: {resp.status_code}")
    st.code(resp.text[:500])
except Exception as e:
    st.error(f"❌ Raw HTTP failed: {type(e).__name__}: {e}")
    st.code(traceback.format_exc())

st.markdown("### 3. Supabase client test")
try:
    from supabase import create_client
    sb = create_client(url, key)
    result = sb.table("topics").select("id").limit(1).execute()
    st.success(f"✅ Connected! Result: {result}")
except Exception as e:
    st.error(f"❌ Full error: {type(e).__name__}: {e}")
    st.code(traceback.format_exc())
