#!/usr/bin/env python3
"""
Live Kofile Name Search for Huntsman Jackie (01/01/1974 - 08/24/2026)
and parallel downloading with 4 concurrent workers into /Volumes/davidlls/assignments/test_huntsman_jackie.
"""

import os
import sys
import time
import re
import concurrent.futures
from playwright.sync_api import sync_playwright

SEARCH_NAME = "Huntsman Jackie"
START_DATE = "01/01/1974"
END_DATE = "08/24/2026"
OUTPUT_DIR = "/Volumes/davidlls/assignments/test_huntsman_jackie"

os.makedirs(OUTPUT_DIR, exist_ok=True)

print("=" * 80)
print(f"🔍 STEP 1: SEARCHING KOFILE FOR: {SEARCH_NAME} ({START_DATE} to {END_DATE})")
print(f"📁 Destination Folder: {OUTPUT_DIR}")
print("=" * 80)

def search_kofile_name(name, start_date, end_date):
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

            print(f"✍️ Entering Name: {name}...")
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
                print("⚠️ No results found or search timed out.")
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

                print(f"📄 Page {page_num}: Extracted {len(rows_data)} rows from table.")
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
            print(f"❌ Error during Kofile search: {e}")
            try:
                browser.close()
            except:
                pass
            return []


def download_single_doc_worker(vol, pg, doc_type, dest_dir, worker_id=1, stagger_sec=0.0):
    if stagger_sec > 0:
        time.sleep(stagger_sec)

    start_time = time.time()
    tag = f"[Worker-{worker_id} | Vol {vol} Pg {pg}]"

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

            clean_type = "".join(c for c in doc_type if c.isalnum() or c in " _-").strip() or "DOC"
            with page.expect_download(timeout=60000) as download_info:
                page.frame(name='bodyframe').evaluate("""
                    var instrId = document.getElementById("documentFrame").contentWindow.getInstrumentId();
                    var numPages = document.getElementById("documentFrame").contentWindow.getNumPages();
                    continueDownloadDocImage(instrId, true, numPages, "printall", false);
                """)

            download = download_info.value
            out_file = os.path.join(dest_dir, f"{clean_type} {vol}-{pg}.pdf")
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


if __name__ == "__main__":
    records = search_kofile_name(SEARCH_NAME, START_DATE, END_DATE)
    print(f"\n🎉 Search complete! Found {len(records)} total records for {SEARCH_NAME}.")
    
    if not records:
        print("No documents found to download.")
        sys.exit(0)

    # Save text manifest
    manifest_path = os.path.join(OUTPUT_DIR, "Search_Results_Manifest.txt")
    with open(manifest_path, "w") as f:
        f.write(f"Kofile Search Results for: {SEARCH_NAME} ({START_DATE} - {END_DATE})\n")
        f.write(f"Total Records: {len(records)}\n\n")
        for r in records:
            f.write(f"[{r['type']}] Vol/Pg: {r['volpg']} | Date: {r['date']} | Grantor: {r['grantor']} | Grantee: {r['grantee']} | Legal: {r['legal']}\n")
    print(f"📝 Saved full records manifest to: {manifest_path}")

    # Deduplicate valid Book/Page pairs
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
    results = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
        futures = {
            executor.submit(download_single_doc_worker, vol, pg, dtype, OUTPUT_DIR, (idx % 4) + 1, (idx % 4) * 0.4): (vol, pg)
            for idx, (vol, pg, dtype) in enumerate(valid_items)
        }
        for future in concurrent.futures.as_completed(futures):
            res = future.result()
            results.append(res)

    total_dl_time = round(time.time() - dl_start, 2)
    successes = [r for r in results if r["success"]]
    failures = [r for r in results if not r["success"]]

    print("\n" + "=" * 80)
    print("🏁 FINAL SUMMARY & REPORT")
    print("=" * 80)
    print(f"📁 Destination Folder: {OUTPUT_DIR}")
    print(f"⏱️ Total Download Time: {total_dl_time} seconds ({round(total_dl_time/60, 2)} mins)")
    print(f"✅ Success Rate: {len(successes)} / {len(valid_items)} ({round(len(successes)/len(valid_items)*100, 1)}%)")
    if successes:
        print(f"⚡ Effective Speed: 1 document every {round(total_dl_time/len(successes), 2)} seconds ({round(len(successes)/(total_dl_time/60), 1)} docs/min)")
        total_kb = sum(r["size_kb"] for r in successes)
        print(f"📦 Total Data Downloaded: {round(total_kb/1024, 2)} MB")
    print("=" * 80)
