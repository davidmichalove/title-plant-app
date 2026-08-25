#!/usr/bin/env python3
"""
Advanced Title Streaming Engine:
- 5 Thread-Isolated Playwright Scraper Workers (Producers)
- 4 Parallel Gemini 3.6 Flash Direct Byte Inlining Workers (Consumers)
- Real-Time Producer-Consumer Streaming Queue
- Multi-page Kofile search with virtual scroll hydration
- Specialized Tabular Exhibit A/B Parsing & Bi-directional Relational Linking
"""

import os
import sys
import time
import json
import re
import shutil
import queue
import threading
from playwright.sync_api import sync_playwright
from google import genai
from google.genai import types

SEARCH_NAME = "Debolt Charles"
START_DATE = "01/01/1993"
END_DATE = "08/24/2026"
TARGET_LOTS = ["142", "Lot 142", "In Lot 142", "In-Lot 142", "Out Lot 142"]
TARGET_PARCELS = ["42-00124.000", "42-0124.000", "42-000124.000", "42-00124", "42-0124"]

DEST_DIR = "/Volumes/davidlls/assignments/test_debolt_streaming"
HITS_DIR = os.path.join(DEST_DIR, "Hits")
CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

os.makedirs(DEST_DIR, exist_ok=True)
os.makedirs(HITS_DIR, exist_ok=True)

with open(CONFIG_PATH) as f:
    api_key = json.load(f)["GEMINI_API_KEY"]

client = genai.Client(api_key=api_key)

print("=" * 85)
print(f"⚡ ADVANCED TITLE STREAMING & AI TRIAGE ENGINE")
print(f"👤 Search Target: {SEARCH_NAME} ({START_DATE} to {END_DATE})")
print(f"🎯 Target Criteria: Lots: {TARGET_LOTS[0]} | Parcels: {TARGET_PARCELS[0]}")
print(f"📁 Destination: {DEST_DIR}")
print("=" * 85)


def search_kofile_all_pages(name, start_date, end_date):
    """Searches Kofile and navigates all result pages with virtual scroll hydration."""
    all_parsed_records = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        try:
            print("🌐 Connecting to Belmont County Recorder...")
            page.goto("https://countyfusion13.kofiletech.us/countyweb/loginDisplay.action?countyname=BelmontOH", timeout=45000)
            page.locator("input[value='Login as Guest']").click(no_wait_after=True)
            page.wait_for_load_state('domcontentloaded')
            page.wait_for_timeout(2000)

            try:
                accept_btn = page.frame_locator("iframe[name='bodyframe']").locator("input#accept")
                accept_btn.wait_for(state="visible", timeout=15000)
                accept_btn.click()
                page.wait_for_load_state('domcontentloaded')
                page.wait_for_timeout(2000)
            except Exception:
                pass

            print("📋 Navigating to Name Search...")
            try:
                page.frame_locator("iframe[name='bodyframe']").locator("text='Search Public Records'").first.click(timeout=8000)
            except:
                pass
            page.wait_for_timeout(2000)

            page.frame_locator("iframe[name='bodyframe']").frame_locator("iframe[name='dynSearchFrame']").get_by_role("tab", name="Name").click()
            page.wait_for_timeout(2000)

            crit_frame = page.frame_locator("iframe[name='bodyframe']").frame_locator("iframe[name='dynSearchFrame']").frame_locator("iframe[name='criteriaframe']")
            try:
                crit_frame.locator("img#clearIcon").click(timeout=1000)
            except:
                pass

            print(f"✍️ Entering Name: {name} ({start_date} - {end_date})...")
            crit_frame.get_by_label("Name", exact=True).fill(name)

            if start_date:
                try: crit_frame.locator("input[aria-label='Recorded Date From'].textbox-text, input[name*='From']").first.fill(start_date)
                except: pass
            if end_date:
                try: crit_frame.locator("input[aria-label='Recorded Date To'].textbox-text, input[name*='To']").first.fill(end_date)
                except: pass

            print("🔎 Clicking Search...")
            page.frame_locator("iframe[name='bodyframe']").frame_locator("iframe[name='dynSearchFrame']").locator("img#imgSearch").click()
            page.wait_for_timeout(4000)

            reslist = page.frame_locator("iframe[name='bodyframe']").frame_locator("iframe[name='resultFrame']").frame_locator("iframe[name='resultListFrame']")
            try:
                reslist.locator("tr").first.wait_for(state="visible", timeout=30000)
            except Exception:
                print("⚠️ No results found.")
                browser.close()
                return []

            page_num = 1
            while True:
                # 1. Virtual Scroll Hydration on current page
                for _ in range(3):
                    reslist.locator("body").evaluate("() => window.scrollTo(0, document.body.scrollHeight)")
                    page.wait_for_timeout(1500)

                rows_data = reslist.locator("tr").evaluate_all("""
                    trs => trs.map(tr => {
                        let cells = Array.from(tr.querySelectorAll('td, th'));
                        return cells.map(c => c.innerText.split('\\n').join(' ').trim());
                    })
                """)

                print(f"📄 Result Page {page_num}: Hydrated {len(rows_data)} table rows.")
                for row_cells in rows_data:
                    if not row_cells or len(row_cells) < 4:
                        continue
                    date_idx = -1
                    for idx, cell in enumerate(row_cells):
                        if re.match(r'^\d{2}/\d{2}/\d{4}$', cell):
                            date_idx = idx
                            break
                    if date_idx != -1 and date_idx >= 3:
                        inst = row_cells[date_idx - 3]
                        book = row_cells[date_idx - 2]
                        pg = row_cells[date_idx - 1]
                        rec_date = row_cells[date_idx]
                        doc_type = row_cells[date_idx + 1] if len(row_cells) > date_idx + 1 else "DOC"
                        direct_role = row_cells[date_idx + 2] if len(row_cells) > date_idx + 2 else ""
                        direct_name = row_cells[date_idx + 3] if len(row_cells) > date_idx + 3 else ""
                        other_role = row_cells[date_idx + 4] if len(row_cells) > date_idx + 4 else ""
                        other_name = row_cells[date_idx + 5] if len(row_cells) > date_idx + 5 else ""
                        legal_desc = row_cells[date_idx + 6] if len(row_cells) > date_idx + 6 else ""

                        volpg = f"{book}/{pg}" if book and pg else (inst or "N/A")
                        grantor = f"{direct_name} ({direct_role})" if direct_role else direct_name
                        grantee = f"{other_name} ({other_role})" if other_role else other_name

                        all_parsed_records.append({
                            "inst": inst, "type": doc_type, "vol": book, "pg": pg,
                            "volpg": volpg, "date": rec_date, "grantor": grantor,
                            "grantee": grantee, "legal": legal_desc
                        })

                # Check if multi-page navigation button is present
                nav_frame = page.frame_locator("iframe[name='bodyframe']").frame_locator("iframe[name='resultFrame']").frame_locator("iframe[name='resultNavFrame']")
                try:
                    next_btn = nav_frame.locator("input[value='>'], a:has-text('>'), img[title*='Next'], img[alt*='Next']").first
                    if next_btn.is_visible() and next_btn.is_enabled():
                        next_btn.click()
                        page.wait_for_timeout(3000)
                        page_num += 1
                        reslist = page.frame_locator("iframe[name='bodyframe']").frame_locator("iframe[name='resultFrame']").frame_locator("iframe[name='resultListFrame']")
                    else:
                        break
                except Exception:
                    break

            browser.close()
            return all_parsed_records

        except Exception as e:
            print(f"❌ Error during search: {e}")
            try: browser.close()
            except: pass
            return []


def download_single_doc_worker(vol, pg, doc_type, dest_dir, worker_id=1, stagger_sec=0.0):
    """Downloads a single document from Kofile using Playwright sync_api in its own thread."""
    if stagger_sec > 0:
        time.sleep(stagger_sec)

    clean_type = "".join(c for c in doc_type if c.isalnum() or c in " _-").strip() or "DOC"
    out_file = os.path.join(dest_dir, f"{vol}-{pg} {clean_type}.pdf")
    if os.path.exists(out_file) and os.path.getsize(out_file) > 1000:
        return {"vol": vol, "pg": pg, "dtype": doc_type, "file": out_file, "success": True, "cached": True}

    t0 = time.time()
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        try:
            page.goto("https://countyfusion13.kofiletech.us/countyweb/loginDisplay.action?countyname=BelmontOH", timeout=45000)
            page.locator("input[value='Login as Guest']").click(no_wait_after=True)
            page.wait_for_load_state('domcontentloaded')
            page.wait_for_timeout(1000)

            try:
                accept_btn = page.frame_locator("iframe[name='bodyframe']").locator("input#accept")
                accept_btn.wait_for(state="visible", timeout=12000)
                accept_btn.click()
                page.wait_for_timeout(1000)
            except Exception:
                pass

            search_pub = page.frame_locator("iframe[name='bodyframe']").locator("text='Search Public Records'").first
            search_pub.wait_for(state="visible", timeout=12000)
            search_pub.click()
            page.wait_for_timeout(1000)

            bp_tab = page.frame_locator("iframe[name='bodyframe']").frame_locator("iframe[name='dynSearchFrame']").get_by_role("tab", name="Book / Page")
            bp_tab.wait_for(state="visible", timeout=12000)
            bp_tab.click()
            page.wait_for_timeout(500)

            crit = page.frame_locator("iframe[name='bodyframe']").frame_locator("iframe[name='dynSearchFrame']").frame_locator("iframe[name='criteriaframe']")
            book_input = crit.get_by_role("textbox", name="Book")
            book_input.wait_for(state="visible", timeout=12000)
            book_input.fill(str(vol))
            crit.get_by_role("textbox", name="Page").fill(str(pg))

            page.frame_locator("iframe[name='bodyframe']").frame_locator("iframe[name='dynSearchFrame']").locator("img#imgSearch").click()
            page.wait_for_timeout(1500)

            reslist = page.frame_locator("iframe[name='bodyframe']").frame_locator("iframe[name='resultFrame']").frame_locator("iframe[name='resultListFrame']")
            reslist.locator("tr").first.wait_for(state="visible", timeout=25000)

            rows = reslist.locator("tr")
            target_row = None
            for i in range(rows.count()):
                txt = rows.nth(i).inner_text().strip()
                if not txt or "Instrument" in txt or "Book/Page" in txt or "Type" in txt:
                    continue
                target_row = rows.nth(i)
                break

            if not target_row:
                browser.close()
                return {"vol": vol, "pg": pg, "dtype": doc_type, "success": False, "error": "No row found"}

            target_row.dblclick()
            page.on("dialog", lambda d: d.accept())

            page.frame(name='bodyframe').wait_for_function("""
                () => {
                    try {
                        var docFrame = document.getElementById("documentFrame");
                        return docFrame && docFrame.contentWindow && typeof docFrame.contentWindow.getNumPages === 'function' && docFrame.contentWindow.getNumPages() > 0;
                    } catch (e) { return false; }
                }
            """, timeout=45000)

            with page.expect_download(timeout=60000) as dl_info:
                page.frame(name='bodyframe').evaluate("""
                    var instrId = document.getElementById("documentFrame").contentWindow.getInstrumentId();
                    var numPages = document.getElementById("documentFrame").contentWindow.getNumPages();
                    continueDownloadDocImage(instrId, true, numPages, "printall", false);
                """)

            download = dl_info.value
            download.save_as(out_file)
            elapsed = round(time.time() - t0, 2)
            size_kb = round(os.path.getsize(out_file) / 1024, 1)
            print(f"  [Worker-{worker_id}] ⬇️ Downloaded: {os.path.basename(out_file)} ({size_kb} KB in {elapsed}s)")
            browser.close()
            return {"vol": vol, "pg": pg, "dtype": doc_type, "file": out_file, "success": True, "cached": False}

        except Exception as e:
            try: browser.close()
            except: pass
            return {"vol": vol, "pg": pg, "dtype": doc_type, "success": False, "error": str(e)}


def screen_document_with_gemini(pdf_path, doc_type="DOC"):
    """Evaluates PDF with Gemini 3.6 Flash using direct byte inlining and specialized Exhibit/Table parsing."""
    filename = os.path.basename(pdf_path)
    with open(pdf_path, "rb") as f:
        pdf_bytes = f.read()

    prompt = f"""
    You are an expert real estate, oil & gas title attorney examining title records from Belmont County, Ohio.
    Examine this entire document carefully, including Page 1, Page 2, Page 3, and any attached EXHIBIT "A", EXHIBIT "B", Legal Schedules, or Parcel Tables.

    TARGET SEARCH OBJECTIVES:
    1. Check if this document conveys, leases, encumbers, ratifies, assigns, or references:
       - Lot 142 (including "In Lot 142", "In-Lot 142", "Lot Number 142", "Shoe Factory Addition Lot 142", or "In Lot 142 2F")
       - Tax Parcel ID "42-00124.000" (including "42-0124.000", "42-000124.000", "42-00124", "00124.000")
    2. OIL & GAS LEASE SPECIAL INSTRUCTION:
       - Carefully read all tabular columns in Exhibit "A" / Exhibit "B" schedules. Look for Lot 142, Parcel 42-00124.000, or Prior Deed references (e.g., Vol. 794 Page 457).
    3. MORTGAGES & LIENS:
       - Does this mortgage encumber Lot 142 or Parcel 42-00124.000?
    4. RELEASES / SATISFACTIONS / ASSIGNMENTS:
       - What original Mortgage or Lease Book/Volume and Page numbers does this document release, assign, modify, or satisfy? (e.g. Vol 865 Pg 633, Vol 346 Pg 405, etc.)

    Return a strict JSON object:
    {{
      "is_direct_hit": true or false,
      "hit_reasons": ["List matching reasons, e.g. 'Exhibit A lists In Lot 142 (Parcel 42-00124.000)'"],
      "lots_found": ["List all lot numbers"],
      "parcels_found": ["List all parcel IDs"],
      "document_type": "{doc_type}",
      "recorded_date": "MM/DD/YYYY if visible",
      "grantor": "Grantor / Direct Party",
      "grantee": "Grantee / Reverse Party",
      "is_mortgage": true or false,
      "referenced_prior_vol_pgs": ["List any referenced prior deed, mortgage, or lease Vol/Pg numbers (e.g. '865/633', '794/457')"],
      "legal_summary": "1-2 sentence concise summary of the land or tract description",
      "exact_excerpt": "Exact text or table line from document"
    }}
    """

    try:
        res = client.models.generate_content(
            model='gemini-3.6-flash',
            contents=[types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"), prompt],
            config={"response_mime_type": "application/json"}
        )
        data = json.loads(res.text.strip())
        data["filename"] = filename
        data["filepath"] = pdf_path
        return data
    except Exception as e:
        return {
            "filename": filename, "filepath": pdf_path, "is_direct_hit": False,
            "hit_reasons": [f"AI Error: {e}"], "lots_found": [], "parcels_found": [],
            "document_type": doc_type, "recorded_date": "", "grantor": "", "grantee": "",
            "is_mortgage": False, "referenced_prior_vol_pgs": [],
            "legal_summary": f"Error: {e}", "exact_excerpt": ""
        }


def main():
    pipe_start = time.time()

    # 1. Search Kofile
    records = search_kofile_all_pages(SEARCH_NAME, START_DATE, END_DATE)
    print(f"\n🎉 Search complete! Total extracted records: {len(records)}")
    if not records:
        print("No documents found.")
        sys.exit(0)

    # Save search manifest
    with open(os.path.join(DEST_DIR, "Search_Results_Manifest.txt"), "w") as f:
        f.write(f"Kofile Search Results for: {SEARCH_NAME} ({START_DATE} - {END_DATE})\nTotal: {len(records)}\n\n")
        for r in records:
            f.write(f"[{r['type']}] Vol/Pg: {r['volpg']} | Date: {r['date']} | Grantor: {r['grantor']} | Grantee: {r['grantee']} | Legal: {r['legal']}\n")

    # Deduplicate Book/Page
    valid_items = []
    seen = set()
    for r in records:
        v = str(r.get("vol", "")).strip()
        p = str(r.get("pg", "")).strip()
        if v and p and (v, p) not in seen:
            seen.add((v, p))
            valid_items.append((v, p, r["type"]))

    print(f"\n📦 Found {len(valid_items)} unique Book/Page documents to process in streaming pipeline.")

    # 2. Setup Real-time Producer-Consumer Pipeline
    download_queue = queue.Queue()
    for item in valid_items:
        download_queue.put(item)

    ai_queue = queue.Queue()
    all_ai_results = []
    results_lock = threading.Lock()
    download_complete_event = threading.Event()

    # --- GEMINI CONSUMER WORKERS (4 parallel threads) ---
    def gemini_consumer_worker(worker_id):
        while not download_complete_event.is_set() or not ai_queue.empty():
            try:
                item = ai_queue.get(timeout=1.5)
            except queue.Empty:
                continue

            pdf_path, doc_type = item
            fn = os.path.basename(pdf_path)
            t_ai0 = time.time()
            ai_data = screen_document_with_gemini(pdf_path, doc_type)
            elapsed = round(time.time() - t_ai0, 2)

            is_hit = ai_data.get("is_direct_hit", False)
            status_tag = "⭐ LIVE HIT" if is_hit else "⚪ Non-Hit"
            print(f"  [AI-Worker-{worker_id}] 🧠 {status_tag} in {elapsed}s: {fn}")

            with results_lock:
                all_ai_results.append(ai_data)
            ai_queue.task_done()

    gemini_threads = []
    for i in range(4):
        t = threading.Thread(target=gemini_consumer_worker, args=(i+1,), daemon=True)
        t.start()
        gemini_threads.append(t)

    # --- SCRAPER PRODUCER WORKERS (5 parallel workers with 400ms micro-stagger) ---
    print("\n🚀 LAUNCHING 5 PARALLEL SCRAPER WORKERS WITH 400ms STAGGER...")
    
    def scraper_producer_worker(worker_id):
        time.sleep(worker_id * 0.4) # Stagger initial connection
        while True:
            try:
                vol, pg, dtype = download_queue.get_nowait()
            except queue.Empty:
                break

            res = download_single_doc_worker(vol, pg, dtype, DEST_DIR, worker_id=worker_id)
            if res.get("success") and os.path.exists(res.get("file", "")):
                # Stream immediately to Gemini Consumer Queue!
                ai_queue.put((res["file"], dtype))
            download_queue.task_done()

    scraper_threads = []
    for i in range(5):
        t = threading.Thread(target=scraper_producer_worker, args=(i+1,))
        t.start()
        scraper_threads.append(t)

    for st in scraper_threads:
        st.join()

    print("\n✅ All downloads complete! Finishing remaining streaming AI evaluations...")
    download_complete_event.set()
    for gt in gemini_threads:
        gt.join()

    # 3. Bi-Directional Relational Linking
    print("\n" + "=" * 85)
    print("🔗 BI-DIRECTIONAL RELATIONAL LINKING (Mortgages, Releases, Leases, Addendums)")
    print("=" * 85)

    hit_volpgs = set()
    for r in all_ai_results:
        if r.get("is_direct_hit"):
            m = re.search(r'(\d+)-(\d+)', r["filename"])
            if m:
                hit_volpgs.add(f"{m.group(1)}/{m.group(2)}")
                hit_volpgs.add(f"{int(m.group(1))}/{int(m.group(2))}")

    # Check for relational hits (Releases satisfying Hit Mortgages, or Addendums modifying Hit Leases)
    for r in all_ai_results:
        if not r.get("is_direct_hit"):
            refs = r.get("referenced_prior_vol_pgs", [])
            for ref in refs:
                clean_ref = ref.replace(" ", "").replace("-", "/")
                if clean_ref in hit_volpgs:
                    r["is_direct_hit"] = True
                    r["hit_reasons"].append(f"References Hit Instrument Vol/Pg {clean_ref}")
                    print(f"⭐ PROMOTED RELATIONAL HIT: {r['filename']} (Satisfies/Modifies {clean_ref})")
                    break

    # 4. Populate Hits/ Folder & Generate Report
    confirmed_hits = [r for r in all_ai_results if r.get("is_direct_hit")]
    non_hits = [r for r in all_ai_results if not r.get("is_direct_hit")]

    for h in confirmed_hits:
        shutil.copy2(h["filepath"], os.path.join(HITS_DIR, h["filename"]))

    report_path = os.path.join(DEST_DIR, "Title_AI_Hits_Report.md")
    with open(report_path, "w") as f:
        f.write("# 🏆 Advanced Title Streaming & AI Triage Report\n\n")
        f.write(f"**Search Target**: {SEARCH_NAME} ({START_DATE} - {END_DATE})\n")
        f.write(f"**Target Tract Criteria**: Lot {TARGET_LOTS[0]} / Parcel `{TARGET_PARCELS[0]}`\n")
        f.write(f"**Total Documents Examined**: {len(all_ai_results)}\n")
        f.write(f"**Confirmed Hits Isolated**: **{len(confirmed_hits)}**\n")
        f.write(f"**Hits Destination**: `Hits/`\n\n")
        f.write("---\n\n")
        f.write("## ⭐ Confirmed Hits (Deeds, Leases, Mortgages, Releases)\n\n")

        for idx, h in enumerate(sorted(confirmed_hits, key=lambda x: x["filename"]), 1):
            fn = h["filename"]
            dtype = h.get("document_type", "N/A")
            rdate = h.get("recorded_date", "N/A")
            grantor = h.get("grantor", "N/A")
            grantee = h.get("grantee", "N/A")
            lots = ", ".join(h.get("lots_found", [])) or "None"
            parcels = ", ".join(h.get("parcels_found", [])) or "None"
            summary = h.get("legal_summary", "N/A")
            reasons = " | ".join(h.get("hit_reasons", []))
            excerpt = h.get("exact_excerpt", "").strip()

            f.write(f"### {idx}. 📄 `{fn}`\n")
            f.write(f"- **Document Type**: {dtype}\n")
            f.write(f"- **Recorded Date**: {rdate}\n")
            f.write(f"- **Parties**: {grantor} ➔ {grantee}\n")
            f.write(f"- **Hit Reasons**: `{reasons}`\n")
            f.write(f"- **Lots Found**: {lots}\n")
            f.write(f"- **Parcels Found**: {parcels}\n")
            f.write(f"- **Legal Summary**: {summary}\n")
            if excerpt:
                f.write(f"- **Document Excerpt**:\n> *\"{excerpt}\"*\n\n")
            f.write("\n")

        f.write("---\n\n")
        f.write("## ⚪ Non-Hits (Other Townships / Out-of-Scope Tracts)\n\n")
        for nh in sorted(non_hits, key=lambda x: x["filename"]):
            f.write(f"- **`{nh['filename']}`** ({nh.get('document_type')}): {nh.get('legal_summary')}\n")

    total_time = round(time.time() - pipe_start, 2)
    print("\n" + "=" * 85)
    print(f"🏁 STREAMING PIPELINE COMPLETE!")
    print(f"📁 Destination Folder: {DEST_DIR}")
    print(f"🏆 Confirmed Hits: {len(confirmed_hits)} / {len(all_ai_results)}")
    print(f"📁 Hits Saved In: {HITS_DIR}")
    print(f"📄 Full Report: {report_path}")
    print(f"⏱️ Total Streaming Pipeline Time: {total_time}s ({round(total_time/60, 2)} mins)")
    print("=" * 85)


if __name__ == "__main__":
    main()
