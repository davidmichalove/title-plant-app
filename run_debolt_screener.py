#!/usr/bin/env python3
import os
import json
import time
import shutil
import re
import concurrent.futures
from google import genai

TARGET_DIR = "/Volumes/davidlls/assignments/test_debolt_charles"
HITS_DIR = os.path.join(TARGET_DIR, "Hits")
CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

os.makedirs(HITS_DIR, exist_ok=True)

with open(CONFIG_PATH) as f:
    api_key = json.load(f)["GEMINI_API_KEY"]

client = genai.Client(api_key=api_key)

def analyze_pdf(pdf_path):
    filename = os.path.basename(pdf_path)
    tag = f"[{filename}]"
    print(f"{tag} 🤖 Analyzing with Gemini 3.6 Flash...")

    try:
        sample_file = client.files.upload(file=pdf_path, config={'display_name': filename})
        prompt = """
        You are an expert real estate and oil & gas title attorney examining title records in Belmont County, Ohio.
        Read this entire legal document and analyze it carefully.

        CRITICAL SEARCH TARGETS:
        1. Does this document convey, describe, encumber, or reference "Lot 142" (or "Lot # 142", "Lot Number 142", "Lot One Hundred Forty Two", "Out Lot 142", "In-Lot 142", "In Lot 142", "Outlot 142", or subdivision lot 142)?
        2. Does this document reference Parcel ID "42-00124.000" (or "42-0124.000", "42-000124.000", "42-00124", "42-0124", "00124.000", "0124.000")?
        3. If this document is a MORTGAGE: Does it encumber Lot 142 or Parcel 42-00124.000?
        4. If this document is a MORTGAGE RELEASE, SATISFACTION, ASSIGNMENT, OR AMENDMENT: 
           - What original Mortgage Volume/Book and Page does it release or reference? (e.g. Vol 865 Pg 633, Vol 839 Pg 300, etc.)

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
        try: client.files.delete(name=sample_file.name)
        except: pass

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
        print(f"{tag} ❌ Error: {e}")
        return {
            "filename": filename, "filepath": pdf_path, "is_direct_hit": False,
            "hit_reasons": [f"Error: {e}"], "lot_numbers_found": [], "parcel_numbers_found": [],
            "township_section_found": "", "document_type": "UNKNOWN", "recorded_date": "",
            "grantor": "", "grantee": "", "is_mortgage": False, "mortgage_referenced_vol_pg": [],
            "legal_summary": f"Failed: {e}", "exact_excerpt": ""
        }

def main():
    pdf_files = sorted([os.path.join(TARGET_DIR, f) for f in os.listdir(TARGET_DIR) if f.endswith(".pdf") and not f.startswith("._")])
    print(f"Scanning {len(pdf_files)} downloaded PDFs with Gemini 3.6 Flash...")

    ai_results = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
        futures = {executor.submit(analyze_pdf, p): p for p in pdf_files}
        for future in concurrent.futures.as_completed(futures):
            ai_results.append(future.result())

    # Cross-reference Mortgages & Releases
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

    confirmed_hits = [r for r in ai_results if r.get("is_direct_hit")]
    non_hits = [r for r in ai_results if not r.get("is_direct_hit")]

    for h in confirmed_hits:
        shutil.copy2(h["filepath"], os.path.join(HITS_DIR, h["filename"]))

    report_md_path = os.path.join(TARGET_DIR, "Title_AI_Hits_Report.md")
    with open(report_md_path, "w") as f:
        f.write("# 🏆 AI Title Examination & Triage Report\n\n")
        f.write("**Target Tract**: **Lot 142** / Parcel **`42-00124.000`**\n")
        f.write("**Owner Searched**: Debolt Charles (01/01/1993 - 08/24/2026)\n")
        f.write(f"**Total Documents Examined**: {len(ai_results)}\n")
        f.write(f"**Total Confirmed Hits**: **{len(confirmed_hits)}**\n")
        f.write(f"**Hits Output Folder**: `Hits/`\n\n")
        f.write("---\n\n")
        f.write("## ⭐ Confirmed Hits (Affecting Lot 142 / Parcel 42-00124.000)\n\n")

        if confirmed_hits:
            for idx, h in enumerate(sorted(confirmed_hits, key=lambda x: x["filename"]), 1):
                fn = h["filename"]
                dtype = h.get("document_type", "N/A")
                rdate = h.get("recorded_date", "N/A")
                grantor = h.get("grantor", "N/A")
                grantee = h.get("grantee", "N/A")
                lots_str = ", ".join(h.get("lot_numbers_found", [])) or "None"
                parcels_str = ", ".join(h.get("parcel_numbers_found", [])) or "None"
                summary = h.get("legal_summary", "N/A")
                reasons_str = " | ".join(h.get("hit_reasons", []))

                f.write(f"### {idx}. 📄 `{fn}`\n")
                f.write(f"- **Document Type**: {dtype}\n")
                f.write(f"- **Recorded Date**: {rdate}\n")
                f.write(f"- **Parties**: {grantor} ➔ {grantee}\n")
                f.write(f"- **Hit Reasons**: `{reasons_str}`\n")
                f.write(f"- **Lots Found**: {lots_str}\n")
                f.write(f"- **Parcels Found**: {parcels_str}\n")
                f.write(f"- **Legal Summary**: {summary}\n")
                if h.get("exact_excerpt"):
                    clean_ex = h.get("exact_excerpt").strip()
                    f.write(f"- **Document Excerpt**:\n> *\"{clean_ex}\"*\n\n")
                f.write("\n")
        else:
            f.write("No documents matched Lot 142 or Parcel 42-00124.000.\n\n")

        f.write("---\n\n")
        f.write("## ⚪ Non-Hits (Other Tracts / Townships / Lots)\n\n")
        for nh in sorted(non_hits, key=lambda x: x["filename"]):
            fn = nh["filename"]
            dtype = nh.get("document_type")
            twp = nh.get("township_section_found") or "Other"
            summary = nh.get("legal_summary")
            f.write(f"- **`{fn}`** ({dtype}): {twp} | {summary}\n")

    with open(os.path.join(TARGET_DIR, "Title_AI_Hits_Report.json"), "w") as f:
        json.dump({"total": len(ai_results), "hits_count": len(confirmed_hits), "hits": confirmed_hits, "non_hits": non_hits}, f, indent=2)

    print(f"\n🎉 Screening complete! Confirmed Hits: {len(confirmed_hits)} / {len(pdf_files)}")

if __name__ == "__main__":
    main()
