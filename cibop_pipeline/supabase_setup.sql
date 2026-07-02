-- CIBOP Content Pipeline — Supabase Database Setup
-- Run this entire file in your Supabase SQL Editor

-- Enable UUID extension
CREATE EXTENSION IF NOT EXISTS "pgcrypto";

-- ── Topics ──────────────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS topics (
  id uuid DEFAULT gen_random_uuid() PRIMARY KEY,
  name text NOT NULL,
  module_code text,
  status text DEFAULT 'setup',
  created_at timestamptz DEFAULT now()
);

-- ── Topic Files ──────────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS topic_files (
  id uuid DEFAULT gen_random_uuid() PRIMARY KEY,
  topic_id uuid REFERENCES topics(id) ON DELETE CASCADE,
  file_type text,          -- 'ppt', 'uor', 'pdf'
  file_name text,
  file_url text DEFAULT '',
  extracted_text text,     -- JSON string for PPT/UOR, plain text for PDF
  created_at timestamptz DEFAULT now()
);

-- ── UORs ────────────────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS uors (
  id uuid DEFAULT gen_random_uuid() PRIMARY KEY,
  topic_id uuid REFERENCES topics(id) ON DELETE CASCADE,
  uor_id text,
  title text,
  objective text,
  competency text,
  sub_competencies jsonb DEFAULT '[]',
  ears text,
  slide_nos_source text,
  extra jsonb DEFAULT '{}',
  created_at timestamptz DEFAULT now()
);

-- ── Content Plans ────────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS content_plans (
  id uuid DEFAULT gen_random_uuid() PRIMARY KEY,
  topic_id uuid REFERENCES topics(id) ON DELETE CASCADE,
  uor_id text,
  uor_title text,
  sc_id text,
  sc_text text,
  ear_verb text,
  slide_range_start int DEFAULT 0,
  slide_range_end int DEFAULT 0,
  slide_excerpts text DEFAULT '',
  key_terms jsonb DEFAULT '[]',
  question_type text DEFAULT 'MCQ_INLINE',
  plan_approved boolean DEFAULT false,
  created_at timestamptz DEFAULT now()
);

-- ── Generated Content ────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS generated_content (
  id uuid DEFAULT gen_random_uuid() PRIMARY KEY,
  topic_id uuid REFERENCES topics(id) ON DELETE CASCADE,
  uor_id text,
  sc_id text,
  content_type text,       -- 'video_script', 'assessment'
  version int DEFAULT 1,
  content_text text,
  slide_refs_used jsonb DEFAULT '[]',
  created_at timestamptz DEFAULT now(),
  updated_at timestamptz DEFAULT now()
);

-- ── Audit Results ────────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS audit_results (
  id uuid DEFAULT gen_random_uuid() PRIMARY KEY,
  content_id uuid REFERENCES generated_content(id) ON DELETE CASCADE,
  coverage_score float DEFAULT 0,
  order_score float DEFAULT 0,
  fidelity_score float DEFAULT 0,
  overall_score float DEFAULT 0,
  passed boolean DEFAULT false,
  flags jsonb DEFAULT '[]',
  created_at timestamptz DEFAULT now()
);

-- ── Reviews ──────────────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS reviews (
  id uuid DEFAULT gen_random_uuid() PRIMARY KEY,
  content_id uuid REFERENCES generated_content(id) ON DELETE CASCADE,
  reviewer_name text,
  edited_content text,
  approved boolean DEFAULT false,
  comments text DEFAULT '',
  reviewed_at timestamptz DEFAULT now()
);

-- ── Row Level Security (disable for private/internal tool) ───────────────────
-- If you want public access without auth, run these:
ALTER TABLE topics ENABLE ROW LEVEL SECURITY;
ALTER TABLE topic_files ENABLE ROW LEVEL SECURITY;
ALTER TABLE uors ENABLE ROW LEVEL SECURITY;
ALTER TABLE content_plans ENABLE ROW LEVEL SECURITY;
ALTER TABLE generated_content ENABLE ROW LEVEL SECURITY;
ALTER TABLE audit_results ENABLE ROW LEVEL SECURITY;
ALTER TABLE reviews ENABLE ROW LEVEL SECURITY;

-- Allow all operations with the service role key (used by the app)
CREATE POLICY "allow_all_topics" ON topics FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "allow_all_files" ON topic_files FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "allow_all_uors" ON uors FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "allow_all_plans" ON content_plans FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "allow_all_content" ON generated_content FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "allow_all_audits" ON audit_results FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "allow_all_reviews" ON reviews FOR ALL USING (true) WITH CHECK (true);
