"""Supabase client and all database operations for CIBOP pipeline."""

import streamlit as st
from supabase import create_client, Client
import json
from datetime import datetime


@st.cache_resource
def get_supabase() -> Client:
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)


# ── Topics ──────────────────────────────────────────────────────────────────

def list_topics():
    sb = get_supabase()
    return sb.table("topics").select("*").order("created_at", desc=True).execute().data

def create_topic(name: str, module_code: str) -> dict:
    sb = get_supabase()
    return sb.table("topics").insert({
        "name": name,
        "module_code": module_code,
        "status": "setup"
    }).execute().data[0]

def get_topic(topic_id: str) -> dict:
    sb = get_supabase()
    return sb.table("topics").select("*").eq("id", topic_id).single().execute().data

def update_topic_status(topic_id: str, status: str):
    sb = get_supabase()
    sb.table("topics").update({"status": status}).eq("id", topic_id).execute()

def delete_topic(topic_id: str):
    sb = get_supabase()
    sb.table("topics").delete().eq("id", topic_id).execute()


# ── Files ────────────────────────────────────────────────────────────────────

def save_file_record(topic_id: str, file_type: str, file_name: str,
                     extracted_text: str, file_url: str = "") -> dict:
    sb = get_supabase()
    # Delete old record of same type for this topic
    sb.table("topic_files").delete().eq("topic_id", topic_id).eq("file_type", file_type).execute()
    return sb.table("topic_files").insert({
        "topic_id": topic_id,
        "file_type": file_type,
        "file_name": file_name,
        "file_url": file_url,
        "extracted_text": extracted_text
    }).execute().data[0]

def get_files_for_topic(topic_id: str) -> list:
    sb = get_supabase()
    return sb.table("topic_files").select("*").eq("topic_id", topic_id).execute().data

def get_file(topic_id: str, file_type: str) -> dict | None:
    sb = get_supabase()
    rows = sb.table("topic_files").select("*").eq("topic_id", topic_id).eq("file_type", file_type).execute().data
    return rows[0] if rows else None


# ── UORs ─────────────────────────────────────────────────────────────────────

def save_uors(topic_id: str, uors: list):
    sb = get_supabase()
    sb.table("uors").delete().eq("topic_id", topic_id).execute()
    for u in uors:
        sb.table("uors").insert({
            "topic_id": topic_id,
            "uor_id": u["uor_id"],
            "title": u["title"],
            "objective": u["objective"],
            "competency": u.get("competency", ""),
            "sub_competencies": u.get("sub_competencies", []),
            "ears": u.get("ears", ""),
            "slide_nos_source": u.get("slide_nos_source", ""),
            "extra": u.get("extra", {})
        }).execute()

def get_uors(topic_id: str) -> list:
    sb = get_supabase()
    return sb.table("uors").select("*").eq("topic_id", topic_id).order("uor_id").execute().data


# ── Content Plans ────────────────────────────────────────────────────────────

def save_plan_items(topic_id: str, items: list):
    sb = get_supabase()
    sb.table("content_plans").delete().eq("topic_id", topic_id).execute()
    for item in items:
        sb.table("content_plans").insert({
            "topic_id": topic_id,
            "uor_id": item["uor_id"],
            "uor_title": item["uor_title"],
            "sc_id": item["sc_id"],
            "sc_text": item["sc_text"],
            "ear_verb": item["ear_verb"],
            "slide_range_start": item["slide_range_start"],
            "slide_range_end": item["slide_range_end"],
            "slide_excerpts": item["slide_excerpts"],
            "key_terms": item.get("key_terms", []),
            "question_type": item.get("question_type", "MCQ_INLINE"),
            "plan_approved": False
        }).execute()

def get_plan_items(topic_id: str) -> list:
    sb = get_supabase()
    return sb.table("content_plans").select("*").eq("topic_id", topic_id).order("uor_id").execute().data

def approve_plan(topic_id: str):
    sb = get_supabase()
    sb.table("content_plans").update({"plan_approved": True}).eq("topic_id", topic_id).execute()

def update_plan_item(item_id: str, updates: dict):
    sb = get_supabase()
    sb.table("content_plans").update(updates).eq("id", item_id).execute()


# ── Generated Content ────────────────────────────────────────────────────────

def save_generated(topic_id: str, uor_id: str, sc_id: str,
                   content_type: str, content_text: str,
                   slide_refs_used: list, version: int = 1) -> dict:
    sb = get_supabase()
    # Upsert logic: delete existing same type+uor+sc
    sb.table("generated_content").delete()\
      .eq("topic_id", topic_id)\
      .eq("uor_id", uor_id)\
      .eq("sc_id", sc_id)\
      .eq("content_type", content_type).execute()
    return sb.table("generated_content").insert({
        "topic_id": topic_id,
        "uor_id": uor_id,
        "sc_id": sc_id,
        "content_type": content_type,
        "version": version,
        "content_text": content_text,
        "slide_refs_used": slide_refs_used,
        "updated_at": datetime.utcnow().isoformat()
    }).execute().data[0]

def get_generated(topic_id: str, content_type: str = None) -> list:
    sb = get_supabase()
    q = sb.table("generated_content").select("*").eq("topic_id", topic_id)
    if content_type:
        q = q.eq("content_type", content_type)
    return q.order("uor_id").execute().data

def update_generated_text(content_id: str, new_text: str):
    sb = get_supabase()
    sb.table("generated_content").update({
        "content_text": new_text,
        "updated_at": datetime.utcnow().isoformat()
    }).eq("id", content_id).execute()


# ── Audit Results ────────────────────────────────────────────────────────────

def save_audit(content_id: str, coverage: float, order_score: float,
               fidelity: float, flags: list) -> dict:
    sb = get_supabase()
    sb.table("audit_results").delete().eq("content_id", content_id).execute()
    overall = round((coverage + order_score + fidelity) / 3, 1)
    return sb.table("audit_results").insert({
        "content_id": content_id,
        "coverage_score": coverage,
        "order_score": order_score,
        "fidelity_score": fidelity,
        "overall_score": overall,
        "passed": overall >= 80.0,
        "flags": flags
    }).execute().data[0]

def get_audit_for_content(content_id: str) -> dict | None:
    sb = get_supabase()
    rows = sb.table("audit_results").select("*").eq("content_id", content_id).execute().data
    return rows[0] if rows else None

def get_all_audits_for_topic(topic_id: str) -> list:
    sb = get_supabase()
    # Join via generated_content
    content_rows = get_generated(topic_id)
    results = []
    for c in content_rows:
        audit = get_audit_for_content(c["id"])
        results.append({"content": c, "audit": audit})
    return results


# ── Reviews ─────────────────────────────────────────────────────────────────

def save_review(content_id: str, reviewer_name: str, edited_text: str,
                approved: bool, comments: str = "") -> dict:
    sb = get_supabase()
    sb.table("reviews").delete().eq("content_id", content_id).execute()
    return sb.table("reviews").insert({
        "content_id": content_id,
        "reviewer_name": reviewer_name,
        "edited_content": edited_text,
        "approved": approved,
        "comments": comments
    }).execute().data[0]

def get_review(content_id: str) -> dict | None:
    sb = get_supabase()
    rows = sb.table("reviews").select("*").eq("content_id", content_id).execute().data
    return rows[0] if rows else None
