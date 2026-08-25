#!/usr/bin/env python3
"""
Full Automated Pipeline Test for Debolt Charles (01/01/1993 - 08/24/2026):
1. Search Belmont County Kofile.
2. Parallel download all unique documents (4 concurrent workers) into /Volumes/davidlls/assignments/test_debolt_charles/.
3. Send all downloaded documents to Gemini 3.6 Flash for Lot 142 / Parcel 42-00124.000 + Mortgage Release cross-referencing.
4. Auto-populate Hits/ folder and generate Markdown & JSON Title Triage reports.
"""

import os
import sys
import time
import json
import re
import shutil
import concurrent.futures
from playwright.sync_api import sync_playwright
from google import genai

SEARCH_NAME = "Debolt Charles"
START_DATE = "01/01/1993"
END_DATE = "08/24/2026"
OUTPUT_DIR = "/Volumes/davidlls/assignments/test_debolt_charles"
HITS_DIR = os.path.join(OUTPUT_DIR, "Hits")
CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(HITS_DIR, exist_ok=True)

with open(CONFIG_PATH) as f:
    api_key = json.load(f)["GEMINI_API_KEY"]

client = genai.Client(api_key=api_key)

print("=" * 80)
print(f"🚀 FULL AI TITLE PIPELINE: {SEARCH_NAME} ({START_DATE} to {END_DATE})")
print(f"📁 Destination Folder: {OUTPUT_DIR}")
print(f"🎯 Target Tract: Lot 142 / Parcel 42-00124.000 (and variants)")
print("=" * 80)


def search_kofile(name, start_date, end_date):
    parsed_records = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

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

            print("📋 Navigating to Search Public Records...")
            try:
                page.frame_locator("iframe[name='bodyframe']").locator("text='Search Public Records'").first.click(timeout=8000)
            except:
                pass
            page.wait_for_timeout(2000)

            print("🏷️ Selecting Name Tab...")
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
                try:
                    crit_frame.locator("input[aria-label='Recorded Date From'].textbox-text, input[name*='From']").first.fill(start_date)
                except:
                    pass
            if end_date:
                try:
                    crit_frame.locator("input[aria-label='Recorded Date To'].textbox-text, input[name*='To']").first.fill(end_date)
                except:
                    pass

            print("🔎 Clicking Search...")
            page.frame_locator("iframe[name='bodyframe']").frame_locator("iframe[name='dynSearchFrame']").locator("img#imgSearch").click()
            page.wait_for_timeout(3000)

            reslist = page.frame_locator("iframe[name='bodyframe']").frame_locator("iframe[name='resultFrame']").frame_locator("iframe[name='resultListFrame']")
            try:
                reslist.locator("tr").first.wait_for(state="visible", timeout=30000)
            except Exception:
                print("⚠️ No results found.")
                browser.close()
                return []

            page_num = 1
            while True:
                rows_data = reslist.locator("tr").evaluate_all("""
                    trs => trs.map(tr => {
                        let cells = Array.from(tr.querySelectorAll('td, th'));
                        return cells.map(c => c.innerText.split('\\n').join(' ').trim());
                    })
                """)

                print(f"📄 Page {page_num}: Found {len(rows_data)} rows.")
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

                        parsed_records.append({
                            "inst": inst,
                            "type": doc_type,
                            "vol": book,
                            "pg": pg,
                            "volpg": volpg,
                            "date": rec_date,
                            "grantor": grantor,
                            "grantee": grantee,
                            "legal": legal_desc
                        })

                nav_frame = page.frame_locator("iframe[name='bodyframe']").frame_locator("iframe[name='resultFrame']").frame_locator("iframe[name='resultNavFrame']")
                try:
                    next_btn = nav_frame.locator("input[value='>'], a:has-text('>'), img[title*='Next'], img[alt*='Next']").first
                    if next_btn.is_visible() and next_btn.is_enabled():
                        next_btn.click()
                        page.wait_for_timeout(2500)
                        page_num += 1
                        reslist = page.frame_locator("iframe[name='bodyframe']").frame_locator("iframe[name='resultFrame']").frame_locator("iframe[name='resultListFrame']")
                    else:
                        break
                except Exception:
                    break

            browser.close()
            return parsed_records

        except Exception as e:
            print(f"❌ Error during search: {e}")
            try:
                browser.close()
            except:
                pass
            return []


def download_doc_worker(vol, pg, doc_type, dest_dir, worker_id=1, stagger_sec=0.0):
    if stagger_sec > 0:
        time.sleep(stagger_sec)

    start_time = time.time()
    tag = f"[Worker-{worker_id} | Vol {vol} Pg {pg}]"

    clean_type = "".join(c for c in doc_type if c.isalnum() or c in " _-").strip() or "DOC"
    out_file = os.path.join(dest_dir, f"{vol}-{pg} {clean_type}.pdf")
    if os.path.exists(out_file) and os.path.getsize(out_file) > 1000:
        print(f"{tag} ⏩ Already downloaded: {os.path.basename(out_file)}")
        return {"vol": vol, "pg": pg, "success": True, "time": 0.1, "size_kb": os.path.getsize(out_file)/1024, "file": out_file}

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
                accept_btn.wait_for(state="visible", timeout=15000)
                accept_btn.click()
                page.wait_for_load_state('domcontentloaded')
                page.wait_for_timeout(1000)
            except Exception:
                pass

            search_pub = page.frame_locator("iframe[name='bodyframe']").locator("text='Search Public Records'").first
            search_pub.wait_for(state="visible", timeout=15000)
            search_pub.click()
            page.wait_for_timeout(1000)

            bp_tab = page.frame_locator("iframe[name='bodyframe']").frame_locator("iframe[name='dynSearchFrame']").get_by_role("tab", name="Book / Page")
            bp_tab.wait_for(state="visible", timeout=15000)
            bp_tab.click()
            page.wait_for_timeout(500)

            crit = page.frame_locator("iframe[name='bodyframe']").frame_locator("iframe[name='dynSearchFrame']").frame_locator("iframe[name='criteriaframe']")
            book_input = crit.get_by_role("textbox", name="Book")
            book_input.wait_for(state="visible", timeout=15000)
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
                return {"vol": vol, "pg": pg, "success": False, "error": "No result row", "time": round(time.time() - start_time, 2)}

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

            with page.expect_download(timeout=60000) as download_info:
                page.frame(name='bodyframe').evaluate("""
                    var instrId = document.getElementById("documentFrame").contentWindow.getInstrumentId();
                    var numPages = document.getElementById("documentFrame").contentWindow.getNumPages();
                    continueDownloadDocImage(instrId, true, numPages, "printall", false);
                """)

            download = download_info.value
            download.save_as(out_file)

            elapsed = round(time.time() - start_time, 2)
            size_kb = round(os.path.getsize(out_file) / 1024, 1)
            print(f"{tag} ✅ SUCCESS in {elapsed}s ({size_kb} KB) -> {os.path.basename(out_file)}")

            browser.close()
            return {"vol": vol, "pg": pg, "success": True, "time": elapsed, "size_kb": size_kb, "file": out_file}

        except Exception as e:
            elapsed = round(time.time() - start_time, 2)
            print(f"{tag} ❌ ERROR after {elapsed}s: {e}")
            try:
                browser.close()
            except:
                pass
            return {"vol": vol, "pg": pg, "success": False, "error": str(e), "time": elapsed}


def analyze_pdf_with_gemini(pdf_path):
    filename = os.path.basename(pdf_path)
    tag = f"[{filename}]"
    print(f"{tag} 🤖 Analyzing with Gemini 3.6 Flash...")

    try:
        sample_file = client.files.upload(
            file=pdf_path,
            config={'display_name': filename}
        )

        prompt = """
        You are an expert real estate and oil & gas title attorney examining title records in Belmont County, Ohio.
        Read this entire legal document and analyze it carefully.

        CRITICAL SEARCH TARGETS:
        1. Does this document convey, describe, encumber, or reference "Lot 142" (or "Lot # 142", "Lot Number 142", "Lot One Hundred Forty Two", "Out Lot 142", "In-Lot 142", "In Lot 142", "Outlot 142", or subdivision lot 142)?
        2. Does this document reference Parcel ID "42-00124.000" (or "42-0124.000", "42-000124.000", "42-00124", "42-0124", "00124.000", "0124.000")?
        3. If this document is a MORTGAGE: Does it encumber Lot 142 or Parcel 42-00124.000?
        4. If this document is a MORTGAGE RELEASE, SATISFACTION, ASSIGNMENT, OR AMENDMENT: 
           - What original Mortgage Volume/Book and Page does it release or reference? (e.g. Vol 565 Pg 119, Vol 547 Pg 882, etc.)

        Return ONLY a JSON object with this exact structure:
        {
          "is_direct_hit": true or false,
          "hit_reasons": ["List of matching criteria if hit, e.g. 'Conveys Lot 142', 'References Parcel 42-00124.000'"],
          "lot_numbers_found": ["List all lot numbers mentioned"],
          "parcel_numbers_found": ["List all parcel IDs mentioned"],
          "township_section_found": "Township, Section, or Municipality if mentioned (e.g. 'Barnesville', 'Warren Twp', 'Wayne Twp')",
          "document_type": "DEED / LEASE / MORTGAGE / MTG RELEASE / POOL UNIT / etc.",
          "recorded_date": "MM/DD/YYYY if visible",
          "grantor": "Grantor / Direct party",
          "grantee": "Grantee / Reverse party",
          "is_mortgage": true or false,
          "mortgage_referenced_vol_pg": ["List of prior mortgage Vol/Pg numbers released or modified"],
          "legal_summary": "1-2 sentence clear summary of the actual land/tracts described",
          "exact_excerpt": "Exact quote from the document showing lot number, parcel number, or land description"
        }
        """

        response = client.models.generate_content(
            model='gemini-3.6-flash',
            contents=[sample_file, prompt],
            config={'response_mime_type': 'application/json'}
        )

        try:
            client.files.delete(name=sample_file.name)
        except:
            pass

        data = json.loads(response.text.strip())
        data["filename"] = filename
        data["filepath"] = pdf_path

        is_hit = data.get("is_direct_hit", False)
        lots = ", ".join(data.get("lot_numbers_found", [])) or "None"
        parcels = ", ".join(data.get("parcel_numbers_found", [])) or "None"
        status_icon = "⭐ HIT" if is_hit else "⚪ Non-Hit"
        print(f"{tag} -> {status_icon} | Lots: {lots} | Parcels: {parcels}")
        return data

    except Exception as e:
        print(f"{tag} ❌ AI Error: {e}")
        return {
            "filename": filename,
            "filepath": pdf_path,
            "is_direct_hit": False,
            "hit_reasons": [f"Error: {str(e)}"],
            "lot_numbers_found": [],
            "parcel_numbers_found": [],
            "township_section_found": "",
            "document_type": "UNKNOWN",
            "recorded_date": "",
            "grantor": "",
            "grantee": "",
            "is_mortgage": False,
            "mortgage_referenced_vol_pg": [],
            "legal_summary": f"Analysis failed: {e}",
            "exact_excerpt": ""
        }


def main():
    # 1. Search Kofile
    records = search_kofile(SEARCH_NAME, START_DATE, END_DATE)
    print(f"\n🎉 Search complete! Found {len(records)} total records for {SEARCH_NAME}.")
    if not records:
        print("No documents found.")
        sys.exit(0)

    # Save manifest
    manifest_path = os.path.join(OUTPUT_DIR, "Search_Results_Manifest.txt")
    with open(manifest_path, "w") as f:
        f.write(f"Kofile Search Results for: {SEARCH_NAME} ({START_DATE} - {END_DATE})\n")
        f.write(f"Total Records: {len(records)}\n\n")
        for r in records:
            f.write(f"[{r['type']}] Vol/Pg: {r['volpg']} | Date: {r['date']} | Grantor: {r['grantor']} | Grantee: {r['grantee']} | Legal: {r['legal']}\n")
    print(f"📝 Manifest saved to: {manifest_path}")

    # Deduplicate Book/Page
    valid_items = []
    seen = set()
    for r in records:
        vol = str(r.get("vol", "")).strip()
        pg = str(r.get("pg", "")).strip()
        if vol and pg and (vol, pg) not in seen:
            seen.add((vol, pg))
            valid_items.append((vol, pg, r["type"]))

    print("\n" + "=" * 80)
    print(f"🚀 STEP 2: PARALLEL DOWNLOADING {len(valid_items)} UNIQUE DOCUMENTS (4 WORKERS)")
    print("=" * 80)

    dl_start = time.time()
    dl_results = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
        futures = {
            executor.submit(download_doc_worker, vol, pg, dtype, OUTPUT_DIR, (idx % 4) + 1, (idx % 4) * 0.4): (vol, pg)
            for idx, (vol, pg, dtype) in enumerate(valid_items)
        }
        for future in concurrent.futures.as_completed(futures):
            res = future.result()
            dl_results.append(res)

    dl_time = round(time.time() - dl_start, 2)
    successes = [r for r in dl_results if r["success"]]
    print(f"\n✅ Downloaded {len(successes)} / {len(valid_items)} documents in {dl_time}s ({round(dl_time/60, 2)} mins).")

    # Retry any failed downloads once
    failed_items = [(r["vol"], r["pg"], [v[2] for v in valid_items if v[0] == r["vol"] and v[1] == r["pg"]][0]) for r in dl_results if not r["success"]]
    if failed_items:
        print(f"\n🔄 Retrying {len(failed_items)} items...")
        with concurrent.futures.ThreadPoolExecutor(max_workers=3) as executor:
            retries = [executor.submit(download_doc_worker, vol, pg, dtype, OUTPUT_DIR, idx+1, (idx%3)*0.5) for idx, (vol, pg, dtype) in enumerate(failed_items)]
            concurrent.futures.wait(retries)

    print("\n" + "=" * 80)
    print("🧠 STEP 3: GEMINI 3.6 FLASH TITLE SCREENING & TRIAGE")
    print("=" * 80)

    pdf_files = sorted([os.path.join(OUTPUT_DIR, f) for f in os.listdir(OUTPUT_DIR) if f.endswith(".pdf") and not f.startswith("._")])
    ai_results = []

    with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
        futures = {executor.submit(analyze_pdf_with_gemini, p): p for p in pdf_files}
        for future in concurrent.futures.as_completed(futures):
            ai_results.append(future.result())

    # Mortgage Relational Cross-Referencing
    hit_mtg_volpgs = set()
    for r in ai_results:
        if r.get("is_direct_hit") and (r.get("is_mortgage") or "MORTGAGE" in r.get("document_type", "").upper() or "MTG" in r.get("filename", "").upper()):
            m = re.search(r'(\d+)-(\d+)', r["filename"])
            if m:
                hit_mtg_volpgs.add(f"{m.group(1)}/{m.group(2)}")
                hit_mtg_volpgs.add(f"{int(m.group(1))}/{int(m.group(2))}")
                print(f"📌 Registered Hit Mortgage: Vol {m.group(1)} Pg {m.group(2)}")

    for r in ai_results:
        if not r.get("is_direct_hit"):
            doc_type = r.get("document_type", "").upper()
            fn = r["filename"].upper()
            if "RELEASE" in doc_type or "RELEASE" in fn or "SATISFACTION" in doc_type:
                ref_volpgs = r.get("mortgage_referenced_vol_pg", [])
                matched_hit_mtg = None
                for ref in ref_volpgs:
                    clean_ref = ref.replace(" ", "").replace("-", "/")
                    if clean_ref in hit_mtg_volpgs:
                        matched_hit_mtg = clean_ref
                        break
                if matched_hit_mtg:
                    r["is_direct_hit"] = True
                    r["hit_reasons"].append(f"Releases Hit Mortgage Vol/Pg {matched_hit_mtg}")
                    print(f"⭐ PROMOTED TO HIT: {r['filename']} (Releases Hit Mortgage {matched_hit_mtg})")

    # Move hits to Hits/
    confirmed_hits = [r for r in ai_results if r.get("is_direct_hit")]
    non_hits = [r for r in ai_results if not r.get("is_direct_hit")]

    for h in confirmed_hits:
        shutil.copy2(h["filepath"], os.path.join(HITS_DIR, h["filename"]))

    # Write Markdown Report
    report_md_path = os.path.join(OUTPUT_DIR, "Title_AI_Hits_Report.md")
    with open(report_md_path, "w") as f:
        f.write("# 🏆 AI Title Examination & Triage Report\n\n")
        f.write(f"**Target Tract**: **Lot 142** / Parcel **`42-00124.000`**\n")
        f.write(f"**Owner Searched**: {SEARCH_NAME} ({START_DATE} - {END_DATE})\n")
        f.write(f"**Total Documents Examined**: {len(ai_results)}\n")
        f.write(f"**Total Confirmed Hits**: **{len(confirmed_hits)}**\n")
        f.write(f"**Hits Output Folder**: `Hits/`\n\n")
        f.write("---\n\n")

        f.write("## ⭐ Confirmed Hits (Affecting Lot 142 / Parcel 42-00124.000)\n\n")
        if confirmed_hits:
            for idx, h in enumerate(sorted(confirmed_hits, key=lambda x: x['filename']), 1):
                fn = h["filename"]
                dtype = h.get("document_type", "N/A")
                rdate = h.get("recorded_date", "N/A")
                grantor = h.get("grantor", "N/A")
                grantee = h.get("grantee", "N/A")
                lots_str = ", ".join(h.get("lot_numbers_found", [])) or "None"
                parcels_str = ", ".join(h.get("parcel_numbers_found", [])) or "None"
                summary = h.get("legal_summary", "N/A")

                f.write(f"### {idx}. 📄 `{fn}`\n")
                f.write(f"- **Document Type**: {dtype}\n")
                f.write(f"- **Recorded Date**: {rdate}\n")
                f.write(f"- **Parties**: {grantor} ➔ {grantee}\n")
                f.write(f"- **Hit Reasons**: `{' | '.join(h.get('hit_reasons', []))}`\n")
                f.write(f"- **Lots Found**: {lots_str}\n")
                f.write(f"- **Parcels Found**: {parcels_str}\n")
                f.write(f"- **Legal Summary**: {summary}\n")
                if h.get("exact_excerpt"):
                    f.write(f"- **Document Excerpt**:\n> *\"{h.get('exact_excerpt').strip()}\"*\n\n")
                f.write("\n")
        else:
            f.write("No documents matched Lot 142 or Parcel 42-00124.000.\n\n")

        f.write("---\n\n")
        f.write("## ⚪ Non-Hits (Other Tracts / Townships / Lots)\n\n")
        for nh in sorted(non_hits, key=lambda x: x['filename']):
            fn = nh["filename"]
            dtype = nh.get("document_type")
            twp = nh.get("township_section_found") or "Other"
            summary = nh.get("legal_summary")
            f.write(f"- **`{fn}`** ({dtype}): {twp} | {summary}\n")

    # Save JSON Report
    report_json_path = os.path.join(OUTPUT_DIR, "Title_AI_Hits_Report.json")
    with open(report_json_path, "w") as f:
        json.dump({
            "total_documents": len(ai_results),
            "hits_count": len(confirmed_hits),
            "hits": confirmed_hits,
            "non_hits": non_hits
        }, f, indent=2)

    total_pipeline_time = round(time.time() - dl_start, 2)
    print("\n" + "=" * 80)
    print("🏁 PIPELINE COMPLETE!")
    print("=" * 80)
    print(f"📁 Destination Folder: {OUTPUT_DIR}")
    print(f"🏆 Confirmed Hits: {len(confirmed_hits)} / {len(pdf_files)}")
    print(f"📁 Hits Saved In: {HITS_DIR}")
    print(f"📄 Full Report: {report_md_path}")
    print(f"⏱️ Total Pipeline Time: {total_pipeline_time}s ({round(total_pipeline_time/60, 2)} mins)")
    print("=" * 80)


if __name__ == "__main__":
    main()
