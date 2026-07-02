# CIBOP Content Pipeline — Deployment Guide

## Step 1: Supabase Setup (5 minutes)

1. Go to https://supabase.com and create a free account
2. Create a new project (choose a region close to you)
3. Once created, go to **SQL Editor** in the left sidebar
4. Paste the entire contents of `supabase_setup.sql` and click **Run**
5. Go to **Settings → API** and copy:
   - **Project URL** (looks like `https://abcdef.supabase.co`)
   - **anon public** key (starts with `eyJ...`)

## Step 2: Get your Anthropic API Key

1. Go to https://console.anthropic.com
2. Click **API Keys** → **Create Key**
3. Copy the key (starts with `sk-ant-...`)

## Step 3: Deploy to Streamlit Cloud

1. Push this folder to a **private GitHub repository**
   ```
   cd cibop_pipeline
   git init
   git add .
   git commit -m "CIBOP pipeline initial"
   git remote add origin https://github.com/YOUR_USERNAME/cibop-pipeline.git
   git push -u origin main
   ```

2. Go to https://share.streamlit.io
3. Click **New app** → connect your GitHub → select the repo
4. Set **Main file path** to `app.py`
5. Click **Advanced settings** → **Secrets** and paste:
   ```toml
   SUPABASE_URL = "https://YOUR_PROJECT.supabase.co"
   SUPABASE_KEY = "eyJ..."
   ANTHROPIC_API_KEY = "sk-ant-..."
   ```
6. Click **Deploy** — your app will be live in ~2 minutes

## Step 4: Share with Business Owners

Copy the Streamlit URL (e.g., `https://your-app.streamlit.app`) and share it.
All data is stored in Supabase — everyone sees the same topics and content.

---

## How to Use the Pipeline

| Step | Page | What happens |
|------|------|-------------|
| 1 | Topics | Create a new module (e.g. "Exchange TLC") |
| 2 | Setup | Upload the PPT and UOR Excel file |
| 3 | Plan | Review the auto-generated content plan (SC → slides → key terms). **Approve before proceeding.** |
| 4 | Generate | Click Run — content is generated one SC at a time |
| 5 | Audit | Run the audit agent — each item scored on Coverage, Order, Fidelity (80% = pass) |
| 6 | Review | Business owners read, edit inline, approve or reject |
| 7 | Export | Download final approved Video DOCX and Assessment DOCX |

## What prevents hallucination

- The **planning step** maps each sub-competency to exact slide numbers
- The **generation step** receives ONLY the text from those slides — nothing else
- The **audit step** uses a separate AI call to verify that every claim traces back to the slides
- Content below 80% audit score is blocked from export by default
