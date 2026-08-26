# 🦅 Harbinger Land Title Plant & Automation Suite — Agent Onboarding & System Handbook

> **For the incoming Antigravity Agent:**
> This document provides a complete end-to-end overview, architectural blueprint, environment setup, and operational guidelines for the **Belmont County Title Automation Plant**. Read this carefully to get 100% up to speed on the codebase, workflows, design decisions, and tools.

---

## 🎯 1. System Mission & Core Goals

This application is an end-to-end **Automated Landman & Title Plant Suite** designed for real estate and oil & gas title examination across **Belmont County, Ohio**.

### Key Workflows:
1. **1-Click Parcel Setup**: Given a Parcel ID (e.g. `42-01139.000`), the app extracts GIS parcel metadata (acreage, section, township, initial deed volume/page) from local shapefiles, creates standard subfolder trees (`DOCS/`, `MAPS/`, `TAX/`, `WELL INFO/`, `BACKUPS/`), and initializes standardized Excel templates (`PID OR (DATE).xlsx` and `PID RS (DATE).xlsx`).
2. **Kofile CountyFusion Name Search & Streaming AI Triage**:
   - Searches Belmont County public records online via Playwright.
   - Streams downloads with **5 concurrent worker scrapers** using strict party last-name filtering.
   - Streams multimodal PDF inspection into **Gemini 3.6 Flash** (using direct byte inlining) with tabular Exhibit A/B parsing.
   - Evaluates deterministic relational links (releases $\leftrightarrow$ mortgages, addendums/ratifications $\leftrightarrow$ leases).
   - Isolates confirmed hits into `DOCS/<Clean_Party_Name>/Hits/` and writes `Title_AI_Hits_Report.md`.
3. **Runsheet & Ownership Report (OR) Compiler & Sync Engine**:
   - Compiles Runsheets (`PID RS ...xlsx`) into complete, client-ready Ownership Reports (`PID OR ...xlsx`).
   - Automatically populates **Schedule A** (Vesting deeds, surface owners, mineral owners, active leaseholds, notice instruments).
   - Syncs Encumbrances (Easements, Rights of Way, Leases, Mortgages/Deeds of Trust).
   - Supports 2-way GUI sync dialogs (`or_sync_dialog.py`) with dropdown row overrides for every schedule.
4. **Interactive Document Viewer & PDF Tools**:
   - Multi-folder document explorer with integrated thumbnailing, full-screen Preview on macOS, multi-file combiner, and redaction staging.
5. **GIS & Well Map Integration**:
   - Generates high-res Belmont County GIS maps and ODNR interactive oil & gas well production maps.

---

## 📁 2. Codebase Architecture & File Map

```
/Volumes/davidlls/assignments/app/
├── title_work_automator.py       # Main Tkinter Application Entrypoint & GUI Orchestrator
├── or_compiler_engine.py         # Ownership Report compilation & Schedule A sync engine
├── or_sync_dialog.py             # 2-Way GUI Synchronizer between Runsheet and OR
├── runsheet_editor.py            # Standard Runsheet editor with date cleanups & Excel bolding
├── gemini_runsheet_editor.py     # AI-assisted Runsheet Editor grid
├── ai_parser.py                  # Gemini AI prompts, Exhibit A/B parser, abstracting chains
├── og_checker.py                 # ODNR Ohio oil & gas well data scraper
├── court_checker.py              # CourtView civil/criminal docket searcher
├── status_tracker.py             # Project-wide parcel status & assignment tracker
├── submission_email_dialog.py    # Automated delivery email generator
├── gis_map_generator.py          # Belmont GIS parcel mapping tool
├── odnr_map_generator.py         # ODNR GIS well mapping tool
└── config.json                   # Local secrets (GEMINI_API_KEY)
```

---

## ⚙️ 3. Setting Up on a New Computer

### Step 1: Clone or Sync the Git Repository
```bash
git clone git@github.com:davidmichalove/title-plant-app.git
cd title-plant-app
git checkout main
```

### Step 2: System & Python Dependencies
Ensure Python 3.11+ is installed. Install required packages:
```bash
pip3 install playwright google-genai pymupdf openpyxl pandas geopandas \
             tkinterdnd2 requests beautifulsoup4 send2trash Pillow
```

### Step 3: Initialize Playwright Browsers
```bash
playwright install chromium
```

### Step 4: Configure API Keys
Create `config.json` in the application directory:
```json
{
  "GEMINI_API_KEY": "YOUR_GEMINI_API_KEY_HERE"
}
```

### Step 5: External Data & Templates
The application expects the following templates and reference assets in the parent directory (`/Volumes/davidlls/assignments/`):
* `PID OR (DATE)_TEMPLATE (2).xlsx`
* `PID RS (DATE)_TEMPLATE.xlsx`
* `Polygon_Belmont_County_Web_Parcels_20260501085529 (1).zip` (Belmont shapefiles)
* *(Optional)* Local drive archive under `/Volumes/davidlls/drive/` (`DEEDS/`, `MTGS/`) for instant local deed resolution.

---

## 🧠 4. Critical Technical Patterns & Design Rules

### A. Kofile Direct Playwright Scraping
* **Target Portal**: `https://countyfusion13.kofiletech.us/countyweb/loginDisplay.action?countyname=BelmontOH`
* **Session Handling**: Guest login with resilient try/except wait on disclaimer button `input#accept`.
* **Book/Page Search**: Enters Book & Page into criteria frame, waits for result rows.
* **Party Matching**: Always match by party last name (`last_name.upper() in txt.upper()`) to prevent downloading unrelated instruments sharing a volume/page.
* **Document Viewer**: Wait dynamically for `documentFrame` `getNumPages() > 0`.
* **Direct Print/Image Stream**: Evaluate `continueDownloadDocImage(instrId, true, numPages, "printall", false)` to download the official TIFF/PDF bundle.
* **PDF/A Sanitization**: Always strip PDF/A OutputIntents via PyMuPDF (`fitz`) after download:
  ```python
  doc = fitz.open(target_path)
  catalog = doc.pdf_catalog()
  doc.xref_set_key(catalog, "OutputIntents", "null")
  doc.xref_set_key(catalog, "Metadata", "null")
  doc.save(target_path + ".tmp", incremental=False, deflate=True)
  doc.close()
  shutil.move(target_path + ".tmp", target_path)
  ```

### B. Gemini 3.6 Flash Multimodal Streaming
* **Model**: Use `gemini-3.6-flash`.
* **Direct Byte Inlining**: Pass PDF bytes directly in memory with `types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf")`. Do not upload temporary files to the File API unless processing >20MB batches.
* **Structured Response**: Always request `response_mime_type="application/json"` with explicit schema for `is_direct_hit`, `lots_found`, `parcels_found`, `referenced_prior_vol_pgs`, and `legal_summary`.

### C. GUI Thread Safety on macOS
* **Never** modify Tkinter UI states directly from background daemon threads.
* Always use `self.root.after(0, lambda: ...)` or the universal `self.set_buttons_state(state)` helper to prevent buttons from getting stuck in grayed-out disabled states.
* Ensure top-level classes (`AutomatorApp`, `KofileStreamingProgressWindow`, `SOPChatWindow`) remain strictly at module level.

---

## 🚀 5. How to Launch the Application
```bash
python3 app/title_work_automator.py
```
Or via the macOS Desktop launcher:
```bash
/Users/davidmichalove/Desktop/Run_Automate_App.command
```
