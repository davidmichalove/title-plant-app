import os
import glob
import re
import datetime
import openpyxl
from openpyxl.cell.rich_text import CellRichText

BASE_APP_DIR = os.path.dirname(os.path.abspath(__file__))
SHAPE_FILE_DBF = os.path.join(BASE_APP_DIR, "shape_files", "Belmont_County_Parcels.dbf")

def parse_date(date_val):
    if not date_val:
        return None
    if isinstance(date_val, (datetime.datetime, datetime.date)):
        return date_val.strftime("%m/%d/%Y")
    
    s = str(date_val).strip()
    s = re.sub(r'^(DOD|Dated|Effective|Filed|Recorded)\s*', '', s, flags=re.IGNORECASE).strip()
    
    # Try formats
    for fmt in ["%m/%d/%Y", "%m-%d-%Y", "%Y-%m-%d", "%m/%d/%y", "%m-%d-%y"]:
        try:
            dt = datetime.datetime.strptime(s, fmt)
            # Normalize 2-digit years
            if dt.year > 2050:
                dt = dt.replace(year=dt.year - 100)
            return dt.strftime("%m/%d/%Y")
        except ValueError:
            pass
            
    # Regex extract MM/DD/YYYY
    m = re.search(r'(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})', s)
    if m:
        mo, day, yr = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if yr < 50: yr += 2000
        elif yr < 100: yr += 1900
        try:
            dt = datetime.datetime(yr, mo, day)
            return dt.strftime("%m/%d/%Y")
        except: pass
    return None

def get_gis_owner_info(parcel_num):
    if not parcel_num or not os.path.exists(SHAPE_FILE_DBF):
        return {}
        
    try:
        import struct
        with open(SHAPE_FILE_DBF, "rb") as f:
            header = f.read(32)
            num_records, header_len, record_len = struct.unpack("<IHH", header[4:12])
            fields = []
            while True:
                field_desc = f.read(32)
                if field_desc[0] == 0x0D:
                    break
                name = field_desc[:11].replace(b"\x00", b"").decode("ascii", errors="ignore")
                length = field_desc[16]
                fields.append((name, length))
            
            f.seek(header_len)
            target_clean = parcel_num.strip().upper()
            for _ in range(num_records):
                rec_bytes = f.read(record_len)
                if not rec_bytes: break
                offset = 1
                rec = {}
                for name, length in fields:
                    val = rec_bytes[offset:offset+length].decode("ascii", errors="ignore").strip()
                    rec[name] = val
                    offset += length
                
                p_no = rec.get("parcel_no", "").strip().upper()
                if p_no == target_clean or (len(target_clean) > 5 and target_clean in p_no):
                    # Format address
                    add1 = rec.get("ownadd1", "").strip()
                    add2 = rec.get("ownadd2", "").strip()
                    add3 = rec.get("ownadd3", "").strip()
                    zip_code = rec.get("zip", "").strip()
                    
                    addr_lines = []
                    if add1 and add1.upper() != "NAN": addr_lines.append(add1.title())
                    if add2 and add2.upper() != "NAN": addr_lines.append(add2.title())
                    if add3 and add3.upper() != "NAN": 
                        if zip_code and zip_code not in add3:
                            addr_lines.append(f"{add3.title()} {zip_code}")
                        else:
                            addr_lines.append(add3.title())
                    elif zip_code:
                        addr_lines.append(zip_code)
                        
                    return {
                        "name": rec.get("name", "").strip().title(),
                        "lname": rec.get("lname", "").strip().title(),
                        "fname": rec.get("fname", "").strip().title(),
                        "address_lines": addr_lines,
                        "raw": rec
                    }
    except Exception as e:
        print(f"GIS lookup error: {e}")
    return {}

class ORCompilerEngine:
    @classmethod
    def compile_data(cls, pid_dir, parcel_num=None):
        if not parcel_num and pid_dir:
            m = re.search(r'PID\s*([0-9\-\.]+)', os.path.basename(pid_dir), re.IGNORECASE)
            if m: parcel_num = m.group(1).strip()
            
        rs_files = glob.glob(os.path.join(pid_dir, "*RS*.xlsx"))
        # Exclude temp / backup files
        valid_rs = [f for f in rs_files if not os.path.basename(f).startswith("~") and not os.path.basename(f).startswith("._") and "backup" not in f.lower() and "blank" not in f.lower()]
        
        if not valid_rs:
            return None
            
        rs_path = valid_rs[0]
        wb = openpyxl.load_workbook(rs_path, data_only=True)
        ws = wb.active
        
        rows = []
        for r in range(2, ws.max_row + 1):
            itype = str(ws.cell(r, 1).value or '').strip()
            btype = str(ws.cell(r, 2).value or '').strip()
            vol = str(ws.cell(r, 3).value or '').strip()
            pg = str(ws.cell(r, 4).value or '').strip()
            inst_num = str(ws.cell(r, 5).value or '').strip()
            eff_dt = ws.cell(r, 6).value
            file_dt = ws.cell(r, 7).value
            grantor = str(ws.cell(r, 8).value or '').strip()
            grantee = str(ws.cell(r, 9).value or '').strip()
            comments = str(ws.cell(r, 10).value or '').strip()
            
            if not any([itype, btype, vol, pg, grantor, grantee]):
                continue
                
            parsed_eff = parse_date(eff_dt)
            parsed_file = parse_date(file_dt)
            
            rows.append({
                "row_idx": r,
                "itype": itype,
                "btype": btype,
                "vol": vol,
                "pg": pg,
                "inst_num": inst_num,
                "eff_dt_raw": str(eff_dt or ''),
                "eff_dt": parsed_eff,
                "file_dt": parsed_file,
                "grantor": grantor,
                "grantee": grantee,
                "comments": comments
            })
            
        if not rows:
            return None
            
        # 1. Date Range
        all_dates = []
        for r in rows:
            if r["eff_dt"]:
                try:
                    dt_obj = datetime.datetime.strptime(r["eff_dt"], "%m/%d/%Y")
                    all_dates.append((dt_obj, r["eff_dt"]))
                except: pass
                
        all_dates.sort(key=lambda x: x[0])
        earliest_date = all_dates[0][1] if all_dates else "01/01/1900"
        today_date = datetime.date.today().strftime("%m/%d/%Y")
        
        # 2. Vesting Deed & Surface Owner
        vesting_row = None
        deed_keywords = ["deed", "survivorship", "warranty", "quit claim", "fiduciary", "sheriff", "affidavit for transfer", "certificate of transfer"]
        for r in reversed(rows):
            it_l = r["itype"].lower()
            if any(k in it_l for k in deed_keywords):
                vesting_row = r
                break
                
        gis_info = get_gis_owner_info(parcel_num)
        
        surface_owner_name = ""
        surface_tenancy = ""
        acquired_year = ""
        
        if vesting_row:
            raw_grantee = vesting_row["grantee"]
            if vesting_row["eff_dt"]:
                try:
                    acquired_year = f"({vesting_row['eff_dt'].split('/')[-1]})"
                except: pass
                
            # Parse Tenancy
            m_ten = re.search(r',\s*(husband and wife.*|for their joint lives.*|as survivorship tenants.*|a single person.*|unmarried.*|widow.*|a corporation.*|an ohio.*|a delaware.*)', raw_grantee, re.IGNORECASE)
            if m_ten:
                surface_tenancy = m_ten.group(1).strip().upper()
                surface_owner_name = raw_grantee[:m_ten.start()].strip().upper()
            else:
                surface_owner_name = raw_grantee.strip().upper()
        elif gis_info.get("name"):
            surface_owner_name = gis_info["name"].upper()
            
        address_lines = gis_info.get("address_lines", [])
        if not address_lines:
            address_lines = ["Belmont County, OH"]
            
        # 3. Encumbrances Scanning
        # A) Easements & Rights of Way
        easement_rows = []
        for r in rows:
            it_l = r["itype"].lower()
            if any(k in it_l for k in ["right of way", "easement", "pipeline", "powerline", "electric", "highway", "utility", "telephone", "roadway"]):
                summary = f"{r['grantor']} to {r['grantee']}, dated {r['eff_dt'] or 'NA'}, filed {r['file_dt'] or 'NA'}, {r['btype']} Vol {r['vol']}, Pg {r['pg']}"
                easement_rows.append({"row": r, "summary": summary, "included": True})
                
        # B) Oil & Gas Leases
        lease_rows = []
        releases = []
        for r in rows:
            it_l = r["itype"].lower()
            if any(k in it_l for k in ["release of lease", "surrender of lease", "cancellation of lease"]):
                releases.append(r)
                
        for r in rows:
            it_l = r["itype"].lower()
            if any(k in it_l for k in ["lease", "memorandum of lease", "oil and gas lease", "ratification of oil", "addendum to and ratification"]) and not any(k in it_l for k in ["release", "surrender", "assignment"]):
                summary = f"{r['grantor']} to {r['grantee']}, dated {r['eff_dt'] or 'NA'}, filed {r['file_dt'] or 'NA'}, {r['btype']} Vol {r['vol']}, Pg {r['pg']}"
                lease_rows.append({"row": r, "summary": summary, "included": True, "status": "Active"})
                
        # C) Mortgages
        mortgage_rows = []
        satisfactions = []
        for r in rows:
            it_l = r["itype"].lower()
            if any(k in it_l for k in ["satisfaction", "release of mortgage", "certificate of satisfaction", "discharge of mortgage"]):
                satisfactions.append(r)
                
        for r in rows:
            it_l = r["itype"].lower()
            if "mortgage" in it_l and not any(k in it_l for k in ["satisfaction", "release", "assignment", "modification"]):
                is_satisfied = False
                # Check matching volume/page in satisfactions or comments
                ref_target = f"{r['vol']}-{r['pg']}"
                for sat in satisfactions:
                    if (r["vol"] and r["vol"] in sat["comments"]) or (r["pg"] and r["pg"] in sat["comments"]) or (ref_target in sat["comments"]):
                        is_satisfied = True
                        break
                    # Also check if satisfaction date is after mortgage date
                    if sat["grantee"] and r["grantor"] and sat["row_idx"] > r["row_idx"]:
                        # borrower match
                        b_m = sat["grantee"].split()[0].lower() in r["grantor"].lower()
                        if b_m:
                            is_satisfied = True
                            
                summary = f"{r['grantor']} to {r['grantee']}, dated {r['eff_dt'] or 'NA'}, filed {r['file_dt'] or 'NA'}, {r['btype']} Vol {r['vol']}, Pg {r['pg']}"
                mortgage_rows.append({
                    "row": r,
                    "summary": summary,
                    "is_satisfied": is_satisfied,
                    "included": not is_satisfied, # Only unreleased mortgages included by default!
                    "status": "Satisfied" if is_satisfied else "Unreleased"
                })
                
        return {
            "pid_dir": pid_dir,
            "parcel_num": parcel_num,
            "from_date": earliest_date,
            "to_date": today_date,
            "surface_owner": {
                "name": surface_owner_name,
                "tenancy": surface_tenancy,
                "address_lines": address_lines,
                "year": acquired_year,
                "interest": "1"
            },
            "mineral_owner": {
                "name": surface_owner_name,
                "tenancy": surface_tenancy,
                "address_lines": address_lines,
                "year": acquired_year,
                "interest": "1"
            },
            "easements": easement_rows,
            "leases": lease_rows,
            "mortgages": mortgage_rows
        }

    @classmethod
    def apply_to_excel(cls, or_path, data):
        if not os.path.exists(or_path):
            raise FileNotFoundError(f"Ownership report file not found: {or_path}")
            
        wb = openpyxl.load_workbook(or_path)
        ws = wb.active
        
        # 1. Date Range in cell B62 (or finding "RECORDS EXAMINED FROM AND TO:")
        date_str = f"{data['from_date']} TO {data['to_date']}"
        date_cell_found = False
        for r in range(55, ws.max_row + 1):
            if ws.cell(r, 1).value and "RECORDS EXAMINED FROM AND TO" in str(ws.cell(r, 1).value):
                ws.cell(r, 2, date_str)
                date_cell_found = True
                break
        if not date_cell_found:
            ws.cell(62, 2, date_str)
            
        # 2. Surface Owner (Rows 15-19)
        so = data.get("surface_owner", {})
        if so.get("name"):
            ws.cell(15, 1, so["name"])
            if so.get("tenancy"):
                ws.cell(16, 1, so["tenancy"])
            addrs = so.get("address_lines", [])
            if len(addrs) > 0: ws.cell(17, 1, addrs[0])
            if len(addrs) > 1: ws.cell(18, 1, addrs[1])
            if so.get("year"): ws.cell(19, 1, so["year"])
            ws.cell(15, 3, so.get("interest", 1))
            
        # 3. Mineral Owner (Rows 23-27)
        mo = data.get("mineral_owner", {})
        if mo.get("name"):
            ws.cell(23, 1, mo["name"])
            if mo.get("tenancy"):
                ws.cell(24, 1, mo["tenancy"])
            addrs = mo.get("address_lines", [])
            if len(addrs) > 0: ws.cell(25, 1, addrs[0])
            if len(addrs) > 1: ws.cell(26, 1, addrs[1])
            if mo.get("year"): ws.cell(27, 1, mo["year"])
            ws.cell(23, 3, mo.get("interest", 1))
            
        # 4. Easements (Row 52)
        inc_easements = [e["summary"] for e in data.get("easements", []) if e.get("included")]
        for r in range(45, 60):
            if ws.cell(r, 1).value and "EASEMENTS & RIGHTS OF WAY" in str(ws.cell(r, 1).value):
                if inc_easements:
                    txt = "\n".join([f"{i+1}) {s}" for i, s in enumerate(inc_easements)])
                    ws.cell(r + 1, 1, txt)
                else:
                    ws.cell(r + 1, 1, "1) None")
                break
                
        # 5. Oil & Gas Leases (Row 55)
        inc_leases = [l["summary"] for l in data.get("leases", []) if l.get("included")]
        for r in range(48, 60):
            if ws.cell(r, 1).value and "UNRELEASED OIL & GAS LEASES" in str(ws.cell(r, 1).value):
                if inc_leases:
                    txt = "\n".join([f"{i+1}) {s}" for i, s in enumerate(inc_leases)])
                    ws.cell(r + 1, 1, txt)
                else:
                    ws.cell(r + 1, 1, "1) None")
                break
                
        # 6. Unreleased Mortgages (Row 58)
        inc_mortgages = [m["summary"] for m in data.get("mortgages", []) if m.get("included")]
        for r in range(50, 62):
            if ws.cell(r, 1).value and "UNRELEASED MORTGAGES" in str(ws.cell(r, 1).value):
                if inc_mortgages:
                    txt = "\n".join([f"{i+1}) {s}" for i, s in enumerate(inc_mortgages)])
                    ws.cell(r + 1, 1, txt)
                else:
                    ws.cell(r + 1, 1, "1) None")
                break
                
        wb.save(or_path)
        return True
