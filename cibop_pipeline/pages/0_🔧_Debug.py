"""Temporary debug page — delete after fixing connection issues."""
import streamlit as st

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
    st.error(f"❌ Secrets missing or wrong key name: {e}")
    st.stop()

st.markdown("### 2. Supabase connection test")
try:
    from supabase import create_client
    sb = create_client(url, key)
    st.success("✅ Supabase client created")
except Exception as e:
    st.error(f"❌ create_client failed: {e}")
    st.stop()

st.markdown("### 3. Table test")
try:
    result = sb.table("topics").select("id").limit(1).execute()
    st.success(f"✅ topics table exists — {result}")
except Exception as e:
    st.error(f"❌ Table query failed: {e}")
    st.markdown("""
**If you see 'relation topics does not exist':**
→ The SQL setup hasn't been run in Supabase yet.
Go to your Supabase project → SQL Editor → paste `supabase_setup.sql` → Run.

**If you see 'invalid API key' or 401:**
→ Wrong key. Use the **anon public** key from Supabase Settings → API, NOT the service role key.

**If you see a connection/DNS error:**
→ The SUPABASE_URL is wrong. It should look exactly like `https://abcdefgh.supabase.co` with no trailing slash.
    """)
