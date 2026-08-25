#!/usr/bin/env python3
"""
Gemini AI Title Screener & Triage Engine.
Scans all downloaded PDFs in target directory for:
1. Lot 195 (and variants)
2. Parcel 42-001139.000 (and variants)
3. Relational cross-referencing: Mortgage Releases/Modifications referencing Hit Mortgages.

Moves confirmed Hits to Hits/ subfolder and generates a comprehensive Markdown report.
"""

import os
import sys
import json
import time
import shutil
import concurrent.futures
from google import genai

TARGET_DIR = "/Volumes/davidlls/assignments/test_huntsman_jackie"
HITS_DIR = os.path.join(TARGET_DIR, "Hits")
CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

os.makedirs(HITS_DIR, exist_ok=True)

with open(CONFIG_PATH) as f:
    api_key = json.load(f)["GEMINI_API_KEY"]

client = genai.Client(api_key=api_key)

print("=" * 80)
print("🧠 GEMINI 3.6 FLASH TITLE SCREENER & TRIAGE ENGINE")
print(f"📁 Target Folder: {TARGET_DIR}")
print(f"🎯 Hit Criteria: [Lot 195] OR [Parcel 42-001139.000] OR [Release of Hit Mortgage]")
print(f"📁 Hits Destination: {HITS_DIR}")
print("=" * 80)

def analyze_document_with_gemini(pdf_path):
    filename = os.path.basename(pdf_path)
    tag = f"[{filename}]"
    print(f"{tag} 🤖 Sending to Gemini 3.6 Flash for title analysis...")

    try:
        sample_file = client.files.upload(
            file=pdf_path,
            config={'display_name': filename}
        )

        prompt = """
        You are an expert oil, gas, and real estate title attorney and landman examining title records from Belmont County, Ohio.
        Read this entire legal document and analyze it carefully.

        CRITICAL SEARCH TARGETS:
        1. Does this document convey, describe, encumber, or reference "Lot 195" (or "Lot # 195", "Lot Number 195", "Lot One Hundred Ninety Five", "Out Lot 195", "In-Lot 195", or subdivision lot 195)?
        2. Does this document reference Parcel ID "42-001139.000" (or "42-01139.000", "42-001139", "42-01139", "001139.000")?
        3. If this document is a MORTGAGE: Does it encumber Lot 195 or Parcel 42-001139.000?
        4. If this document is a MORTGAGE RELEASE, SATISFACTION, ASSIGNMENT, OR AMENDMENT: 
           - What original Mortgage Volume/Book and Page does it release or reference? (e.g. Vol 565 Pg 119, Vol 547 Pg 882, etc.)

        Return ONLY a JSON object with this exact structure:
        {
          "is_direct_hit": true or false,
          "hit_reasons": ["List of matching criteria if hit, e.g. 'Conveys Lot 195', 'References Parcel 42-001139.000'"],
          "lot_numbers_found": ["List all lot numbers mentioned, e.g. 'Lot 195', 'Lot 196'"],
          "parcel_numbers_found": ["List all parcel IDs mentioned in document"],
          "township_section_found": "Township and Section if mentioned (e.g. 'Wayne Twp Sec 18' or 'Barnesville')",
          "document_type": "DEED / LEASE / MORTGAGE / MTG RELEASE / POOL UNIT / etc.",
          "recorded_date": "MM/DD/YYYY if visible",
          "grantor": "Grantor / Direct party",
          "grantee": "Grantee / Reverse party",
          "is_mortgage": true or false,
          "mortgage_referenced_vol_pg": ["List of any prior mortgage Vol/Pg numbers this document releases or modifies, e.g. '565/119'"],
          "legal_summary": "1-2 sentence clear summary of the actual lands/tracts described in this instrument",
          "exact_excerpt": "Exact quote from the document showing the lot number, parcel number, or land description"
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
        print(f"{tag} ❌ Error analyzing: {e}")
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


def run_full_title_screening():
    pdf_files = sorted([os.path.join(TARGET_DIR, f) for f in os.listdir(TARGET_DIR) if f.endswith(".pdf") and not f.startswith("._")])
    print(f"📚 Found {len(pdf_files)} PDF documents to scan in {TARGET_DIR}...\n")

    start_time = time.time()
    all_results = []

    # Run parallel scanning (4 concurrent requests to Gemini)
    with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
        futures = {executor.submit(analyze_document_with_gemini, p): p for p in pdf_files}
        for future in concurrent.futures.as_completed(futures):
            res = future.result()
            all_results.append(res)

    print("\n" + "=" * 80)
    print("🔗 PASS 2: MORTGAGE <-> RELEASE RELATIONAL CROSS-REFERENCING")
    print("=" * 80)

    # 1. Collect all Hit Mortgages (Vol/Pg pairs)
    hit_mortgage_volpgs = set()
    for r in all_results:
        if r.get("is_direct_hit") and (r.get("is_mortgage") or "MORTGAGE" in r.get("document_type", "").upper() or "MTG" in r.get("filename", "").upper()):
            # Extract Vol and Pg from filename or data
            import re
            m = re.search(r'(\d+)-(\d+)', r["filename"])
            if m:
                hit_mortgage_volpgs.add(f"{m.group(1)}/{m.group(2)}")
                hit_mortgage_volpgs.add(f"{int(m.group(1))}/{int(m.group(2))}")
                print(f"📌 Registered Hit Mortgage: Vol {m.group(1)} Pg {m.group(2)} ({r['filename']})")

    # 2. Check if any Release references a Hit Mortgage
    for r in all_results:
        if not r.get("is_direct_hit"):
            doc_type = r.get("document_type", "").upper()
            fn = r["filename"].upper()
            if "RELEASE" in doc_type or "RELEASE" in fn or "SATISFACTION" in doc_type or "ASSIGN" in doc_type:
                ref_volpgs = r.get("mortgage_referenced_vol_pg", [])
                # Also check filename/summary for referenced vol/pg
                matched_hit_mtg = None
                for ref in ref_volpgs:
                    clean_ref = ref.replace(" ", "").replace("-", "/")
                    if clean_ref in hit_mortgage_volpgs:
                        matched_hit_mtg = clean_ref
                        break

                if matched_hit_mtg:
                    r["is_direct_hit"] = True
                    r["is_relational_hit"] = True
                    r["hit_reasons"].append(f"Releases/Modifies Hit Mortgage Vol/Pg {matched_hit_mtg}")
                    print(f"⭐ PROMOTED TO HIT: {r['filename']} (Releases Hit Mortgage Vol/Pg {matched_hit_mtg})")

    # 3. Move/Copy Hits into Hits/ folder
    confirmed_hits = [r for r in all_results if r.get("is_direct_hit")]
    non_hits = [r for r in all_results if not r.get("is_direct_hit")]

    print("\n" + "=" * 80)
    print(f"📦 ORGANIZING HITS INTO {HITS_DIR}")
    print("=" * 80)

    for h in confirmed_hits:
        src = h["filepath"]
        dst = os.path.join(HITS_DIR, h["filename"])
        shutil.copy2(src, dst)
        print(f"✅ Copied to Hits/: {h['filename']}")

    # 4. Generate Markdown Report
    report_md_path = os.path.join(TARGET_DIR, "Title_AI_Hits_Report.md")
    with open(report_md_path, "w") as f:
        f.write("# 🏆 AI Title Examination & Triage Report\n\n")
        f.write(f"**Target Tract**: Lot 195 / Parcel `42-001139.000`\n")
        f.write(f"**Owner Searched**: Huntsman Jackie (1974 - 2026)\n")
        f.write(f"**Total Documents Examined**: {len(all_results)}\n")
        f.write(f"**Total Confirmed Hits**: **{len(confirmed_hits)}**\n")
        f.write(f"**Hits Output Folder**: `Hits/`\n\n")
        f.write("---\n\n")

        f.write("## ⭐ Confirmed Hits (Affecting Lot 195 / Parcel 42-001139.000)\n\n")
        if confirmed_hits:
            for idx, h in enumerate(sorted(confirmed_hits, key=lambda x: x['filename']), 1):
                f.write(f"### {idx}. 📄 `{h['filename']}`\n")
                f.write(f"- **Document Type**: {h.get('document_type', 'N/A')}\n")
                f.write(f"- **Recorded Date**: {h.get('recorded_date', 'N/A')}\n")
                f.write(f"- **Parties**: {h.get('grantor', 'N/A')} ➔ {h.get('grantee', 'N/A')}\n")
                f.write(f"- **Hit Reasons**: `{' | '.join(h.get('hit_reasons', []))}`\n")
                f.write(f"- **Lots Found**: {', '.join(h.get('lot_numbers_found', [])) or 'None'}\n")
                f.write(f"- **Parcels Found**: {', '.join(h.get('parcel_numbers_found', [])) or 'None'}\n")
                f.write(f"- **Legal Summary**: {h.get('legal_summary', 'N/A')}\n")
                if h.get('exact_excerpt'):
                    f.write(f"- **Document Excerpt**:\n> *\"{h.get('exact_excerpt')}\"*\n\n")
                f.write("\n")
        else:
            f.write("No documents matched Lot 195 or Parcel 42-001139.000.\n\n")

        f.write("---\n\n")
        f.write("## ⚪ Non-Hits (Other Tracts / Townships)\n\n")
        for nh in sorted(non_hits, key=lambda x: x['filename']):
            f.write(f"- **`{nh['filename']}`** ({nh.get('document_type')}): {nh.get('township_section_found') or 'Other'} | {nh.get('legal_summary')}\n")

    # Save JSON report
    report_json_path = os.path.join(TARGET_DIR, "Title_AI_Hits_Report.json")
    with open(report_json_path, "w") as f:
        json.dump({
            "total_documents": len(all_results),
            "hits_count": len(confirmed_hits),
            "hits": confirmed_hits,
            "non_hits": non_hits
        }, f, indent=2)

    total_time = round(time.time() - start_time, 2)
    print("\n" + "=" * 80)
    print("🏁 SCREENING COMPLETE!")
    print("=" * 80)
    print(f"⏱️ Total AI Scanning Time: {total_time} seconds ({round(total_time/len(pdf_files), 2)}s / doc)")
    print(f"🏆 Confirmed Hits: {len(confirmed_hits)} / {len(pdf_files)}")
    print(f"📁 Hits Saved In: {HITS_DIR}")
    print(f"📄 Full Report Saved: {report_md_path}")
    print("=" * 80)

if __name__ == "__main__":
    run_full_title_screening()
