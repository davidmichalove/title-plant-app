import os
import glob
import subprocess
import datetime
import re
import tkinter as tk
from tkinter import ttk, messagebox
from or_compiler_engine import ORCompilerEngine, get_gis_owner_info, format_encumbrance_short, parse_lease_details

class ORSyncDialog(tk.Toplevel):
    def __init__(self, master, pid_dir, parcel_num=None, rs_path=None):
        super().__init__(master)
        self.title("📊 Ownership Report (OR) Auto-Compiler & Sync")
        self.geometry("920x820")
        self.minsize(780, 600)
        self.transient(master)
        self.attributes("-topmost", True)

        self.pid_dir = pid_dir
        self.parcel_num = parcel_num
        self.rs_path = rs_path

        self.data = ORCompilerEngine.compile_data(self.pid_dir, self.parcel_num, rs_path=self.rs_path)
        if not self.data:
            messagebox.showerror("Error", f"Could not extract runsheet rows from:\n{self.rs_path or self.pid_dir}", parent=self)
            self.destroy()
            return

        self.all_rows = self.data.get("all_rows", [])
        self._find_or_file()
        self._build_ui()

    def _find_or_file(self):
        or_files = glob.glob(os.path.join(self.pid_dir, "*OR*.xlsx"))
        valid_or = [f for f in or_files if not os.path.basename(f).startswith("~") and not os.path.basename(f).startswith("._") and "backup" not in f.lower()]
        self.or_path = valid_or[0] if valid_or else None

    def _format_row_label(self, r):
        idx = r.get("row_idx", "?")
        it = r.get("itype", "Doc")
        bt = r.get("btype", "")
        vol = r.get("vol", "")
        pg = r.get("pg", "")
        dt = r.get("eff_dt") or r.get("file_dt") or ""
        g1 = r.get("grantor", "")
        g2 = r.get("grantee", "")
        bp_str = f"{bt} {vol}/{pg}".strip() if vol and pg else (bt or it)
        parties = f"{g1} ➔ {g2}".strip() if g1 or g2 else ""
        return f"Row {idx:02d}: [{it}] {bp_str} | {dt} | {parties}"

    def _build_ui(self):
        # 1. Top Header: Date Range
        top_frame = ttk.LabelFrame(self, text="📅 Records Examined Date Range (Cell B62)", padding=10)
        top_frame.pack(fill=tk.X, padx=12, pady=(10, 4))

        d_row = ttk.Frame(top_frame)
        d_row.pack(fill=tk.X)

        ttk.Label(d_row, text="FROM (Earliest RS Date):", font=("Helvetica", 11, "bold")).pack(side=tk.LEFT, padx=(0, 5))
        self.from_date_var = tk.StringVar(value=self.data["from_date"])
        ttk.Entry(d_row, textvariable=self.from_date_var, width=14, font=("Helvetica", 11)).pack(side=tk.LEFT, padx=(0, 20))

        ttk.Label(d_row, text="TO (Current Date):", font=("Helvetica", 11, "bold")).pack(side=tk.LEFT, padx=(0, 5))
        self.to_date_var = tk.StringVar(value=self.data["to_date"])
        ttk.Entry(d_row, textvariable=self.to_date_var, width=14, font=("Helvetica", 11)).pack(side=tk.LEFT)

        # 2. Notebook Tabs
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=12, pady=4)

        self.tab_owners = ttk.Frame(self.notebook, padding=10)
        self.tab_easements = ttk.Frame(self.notebook, padding=10)
        self.tab_leases = ttk.Frame(self.notebook, padding=10)
        self.tab_mortgages = ttk.Frame(self.notebook, padding=10)

        self.notebook.add(self.tab_owners, text="🏠 Owners & Leasehold")
        self.notebook.add(self.tab_easements, text=f"🛣️ Easements ({len(self.data['easements'])})")
        self.notebook.add(self.tab_leases, text=f"🛢️ O&G Leases ({len(self.data['leases'])})")
        self.notebook.add(self.tab_mortgages, text=f"🏦 Mortgages ({len(self.data['mortgages'])})")

        # =========================================================================
        # --- TAB 1: OWNERS & LEASEHOLD OPTIONS ---
        # =========================================================================
        # Vesting Deed Row Selector
        row_sel_frame = ttk.LabelFrame(self.tab_owners, text="📜 Select Vesting Deed / Current Owner Row from Runsheet", padding=8)
        row_sel_frame.pack(fill=tk.X, pady=(0, 8))

        row_options = [self._format_row_label(r) for r in self.all_rows]
        self.vesting_combo_var = tk.StringVar()

        # Find default matching vesting row
        curr_vesting = self.data.get("vesting_deed")
        default_vesting_idx = 0
        if curr_vesting:
            for idx, r in enumerate(self.all_rows):
                if r.get("row_idx") == curr_vesting.get("row_idx"):
                    default_vesting_idx = idx
                    break

        self.vesting_cb = ttk.Combobox(row_sel_frame, textvariable=self.vesting_combo_var, values=row_options, font=("Helvetica", 11), state="readonly")
        self.vesting_cb.pack(fill=tk.X, padx=5, pady=3)
        if row_options:
            self.vesting_cb.current(default_vesting_idx)
        self.vesting_cb.bind("<<ComboboxSelected>>", self._on_vesting_row_selected)

        # Owner Entry Fields
        fields_frame = ttk.Frame(self.tab_owners)
        fields_frame.pack(fill=tk.X, pady=2)

        so = self.data["surface_owner"]
        ttk.Label(fields_frame, text="Surface Owner (Vesting Grantee):", font=("Helvetica", 11, "bold")).grid(row=0, column=0, sticky="w", pady=3)
        self.so_name_var = tk.StringVar(value=so.get("name", ""))
        ttk.Entry(fields_frame, textvariable=self.so_name_var, font=("Helvetica", 11)).grid(row=0, column=1, sticky="ew", padx=(5, 0), pady=3)

        ttk.Label(fields_frame, text="Marital / Vesting Tenancy:", font=("Helvetica", 11, "bold")).grid(row=1, column=0, sticky="w", pady=3)
        self.so_tenancy_var = tk.StringVar(value=so.get("tenancy", ""))
        ttk.Entry(fields_frame, textvariable=self.so_tenancy_var, font=("Helvetica", 11)).grid(row=1, column=1, sticky="ew", padx=(5, 0), pady=3)

        addrs = so.get("address_lines", ["", ""])
        ttk.Label(fields_frame, text="Mailing Address Line 1:", font=("Helvetica", 11, "bold")).grid(row=2, column=0, sticky="w", pady=3)
        self.so_add1_var = tk.StringVar(value=addrs[0] if len(addrs) > 0 else "")
        ttk.Entry(fields_frame, textvariable=self.so_add1_var, font=("Helvetica", 11)).grid(row=2, column=1, sticky="ew", padx=(5, 0), pady=3)

        ttk.Label(fields_frame, text="Mailing Address Line 2:", font=("Helvetica", 11, "bold")).grid(row=3, column=0, sticky="w", pady=3)
        self.so_add2_var = tk.StringVar(value=addrs[1] if len(addrs) > 1 else "")
        ttk.Entry(fields_frame, textvariable=self.so_add2_var, font=("Helvetica", 11)).grid(row=3, column=1, sticky="ew", padx=(5, 0), pady=3)

        year_default = so.get("year", f"({datetime.date.today().year})") or f"({datetime.date.today().year})"
        ttk.Label(fields_frame, text="Acquired Year:", font=("Helvetica", 11, "bold")).grid(row=4, column=0, sticky="w", pady=3)
        self.so_year_var = tk.StringVar(value=year_default)
        ttk.Entry(fields_frame, textvariable=self.so_year_var, font=("Helvetica", 11), width=16).grid(row=4, column=1, sticky="w", padx=(5, 0), pady=3)

        fields_frame.columnconfigure(1, weight=1)

        ttk.Separator(self.tab_owners, orient="horizontal").pack(fill=tk.X, pady=8)

        # Leasehold Option Frame with Active Lease Selector
        l_frame = ttk.LabelFrame(self.tab_owners, text="🛢️ Leasehold Schedule A Automation", padding=10)
        l_frame.pack(fill=tk.X, pady=3)

        self.leasehold_mode_var = tk.StringVar(value=self.data.get("leasehold_mode", "open_of_record"))
        
        # Build list of parsed lease options
        self.parsed_leases = self.data.get("parsed_leases", [])
        self.lease_options = []
        for pl in self.parsed_leases:
            r = pl["row"]
            idx = r.get("row_idx", "?")
            bk = pl.get("bk_pg", "")
            lessee = pl.get("lessee", "")[:28]
            term = pl.get("term", "")
            royalty = pl.get("royalty", "")
            self.lease_options.append(f"Row {idx:02d}: {bk} ({lessee}...) | Term: {term} | Royalty: {royalty}")

        self.active_lease_var = tk.StringVar()
        if self.lease_options:
            self.active_lease_var.set(self.lease_options[-1])

        rb_populate = ttk.Radiobutton(l_frame, text="✅ Auto-Populate Leasehold Schedule A from active lease:", value="populate", variable=self.leasehold_mode_var)
        rb_populate.pack(anchor="w", pady=(2, 4))

        lease_sel_box = ttk.Frame(l_frame)
        lease_sel_box.pack(fill=tk.X, padx=20, pady=(0, 6))
        
        self.lease_combo = ttk.Combobox(lease_sel_box, textvariable=self.active_lease_var, values=self.lease_options, font=("Helvetica", 10), state="readonly")
        self.lease_combo.pack(fill=tk.X)
        self.lease_combo.bind("<<ComboboxSelected>>", self._on_active_lease_selected)

        rb_open = ttk.Radiobutton(l_frame, text="📜 OPEN OF RECORD (No active lease / Delete 'Leasehold Schedule A' tab)", value="open_of_record", variable=self.leasehold_mode_var)
        rb_open.pack(anchor="w", pady=2)

        # Quick Cleanup Options Section
        opt_frame = ttk.LabelFrame(self.tab_owners, text="⚡ Quick Cleanups", padding=10)
        opt_frame.pack(fill=tk.X, pady=6)

        self.sync_mineral_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt_frame, text="✅ Mirror Mineral Ownership to match Surface Owner (100% Fee Simple)", variable=self.sync_mineral_var).pack(anchor="w", pady=2)

        self.sole_mineral_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt_frame, text="🗑️ Sole Mineral Holder: Delete 'Jim Doe' placeholder (Rows 32–40) & set to 100%", variable=self.sole_mineral_var).pack(anchor="w", pady=2)

        self.delete_notes_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt_frame, text="📝 Delete Notes Block: Delete Rows 42–46 (NOTE #1 / Add note text here)", variable=self.delete_notes_var).pack(anchor="w", pady=2)

        # =========================================================================
        # --- TAB 2: EASEMENTS ---
        # =========================================================================
        self._build_easements_tab()

        # =========================================================================
        # --- TAB 3: O&G LEASES ---
        # =========================================================================
        self._build_leases_tab()

        # =========================================================================
        # --- TAB 4: MORTGAGES ---
        # =========================================================================
        self._build_mortgages_tab()

        # 3. Bottom Action Bar
        bottom_bar = ttk.Frame(self, padding=12)
        bottom_bar.pack(fill=tk.X)

        target_name = os.path.basename(self.or_path) if self.or_path else "Not Found"
        self.target_lbl = ttk.Label(bottom_bar, text=f"Target: {target_name}", font=("Helvetica", 10, "italic"))
        self.target_lbl.pack(side=tk.LEFT)

        btn_apply_open = ttk.Button(bottom_bar, text="🚀 Apply & Open OR Excel", command=lambda: self.apply_changes(open_after=True))
        btn_apply_open.pack(side=tk.RIGHT, padx=(5, 0))

        btn_apply = ttk.Button(bottom_bar, text="💾 Apply to Excel", command=lambda: self.apply_changes(open_after=False))
        btn_apply.pack(side=tk.RIGHT, padx=(5, 0))

        btn_cancel = ttk.Button(bottom_bar, text="Cancel (Esc)", command=self.destroy)
        btn_cancel.pack(side=tk.RIGHT)

        self.bind("<Escape>", lambda e: self.destroy())

    # -------------------------------------------------------------------------
    # TAB 1 Event Handlers
    # -------------------------------------------------------------------------
    def _on_vesting_row_selected(self, event=None):
        sel_idx = self.vesting_cb.current()
        if sel_idx < 0 or sel_idx >= len(self.all_rows):
            return

        sel_row = self.all_rows[sel_idx]
        self.data["vesting_deed"] = sel_row

        raw_grantee = sel_row.get("grantee", "").strip()
        surface_tenancy = ""
        surface_owner_name = raw_grantee

        m_ten = re.search(r',\s*(husband and wife.*|for their joint lives.*|as survivorship tenants.*|a single person.*|unmarried.*|widow.*|a corporation.*|an ohio.*|a delaware.*)', raw_grantee, re.IGNORECASE)
        if m_ten:
            surface_tenancy = m_ten.group(1).strip().upper()
            surface_owner_name = raw_grantee[:m_ten.start()].strip().upper()
        else:
            surface_owner_name = raw_grantee.strip().upper()

        self.so_name_var.set(surface_owner_name)
        self.so_tenancy_var.set(surface_tenancy)

        # Extract Year
        dt_val = sel_row.get("eff_dt") or sel_row.get("file_dt") or ""
        m_yr = re.search(r'\b(19\d{2}|20\d{2})\b', dt_val)
        if m_yr:
            self.so_year_var.set(f"({m_yr.group(1)})")

        # Lookup GIS address for this owner
        gis_info = get_gis_owner_info(self.parcel_num)
        if gis_info.get("address_lines"):
            addrs = gis_info["address_lines"]
            self.so_add1_var.set(addrs[0] if len(addrs) > 0 else "")
            self.so_add2_var.set(addrs[1] if len(addrs) > 1 else "")

    def _on_active_lease_selected(self, event=None):
        sel_idx = self.lease_combo.current()
        if 0 <= sel_idx < len(self.parsed_leases):
            self.data["primary_lease"] = self.parsed_leases[sel_idx]
            self.leasehold_mode_var.set("populate")

    # -------------------------------------------------------------------------
    # TAB 2: EASEMENTS BUILDER
    # -------------------------------------------------------------------------
    def _build_easements_tab(self):
        for widget in self.tab_easements.winfo_children():
            widget.destroy()

        ttk.Label(self.tab_easements, text="Detected Easements to write as 'BookType Vol/Pg':", font=("Helvetica", 11, "bold")).pack(anchor="w", pady=(0, 5))

        self.easement_vars = []
        list_frame = ttk.Frame(self.tab_easements)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=4)

        if not self.data["easements"]:
            ttk.Label(list_frame, text="No Easements currently selected (will write '1) None').", font=("Helvetica", 11, "italic")).pack(anchor="w", pady=5)
        else:
            for idx, e in enumerate(self.data["easements"]):
                var = tk.BooleanVar(value=e.get("included", True))
                self.easement_vars.append((var, e))
                r = e["row"]
                cb = ttk.Checkbutton(list_frame, text=f"Row {r.get('row_idx', '?')}: {e['summary']} ({r.get('grantor', '')} ➔ {r.get('grantee', '')})", variable=var)
                cb.pack(anchor="w", pady=3)

        # Row Selector to add ANY runsheet row
        add_frame = ttk.LabelFrame(self.tab_easements, text="➕ Add Any Other Row from Runsheet as Easement", padding=8)
        add_frame.pack(fill=tk.X, pady=(10, 0))

        row_options = [self._format_row_label(r) for r in self.all_rows]
        add_cb_var = tk.StringVar()
        add_cb = ttk.Combobox(add_frame, textvariable=add_cb_var, values=row_options, font=("Helvetica", 10), state="readonly")
        add_cb.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        if row_options: add_cb.current(0)

        def add_easement():
            sel = add_cb.current()
            if 0 <= sel < len(self.all_rows):
                r = self.all_rows[sel]
                summary = format_encumbrance_short(r)
                # Check if already in list
                if not any(e["row"].get("row_idx") == r.get("row_idx") for e in self.data["easements"]):
                    self.data["easements"].append({"row": r, "summary": summary, "included": True})
                    self.notebook.tab(self.tab_easements, text=f"🛣️ Easements ({len(self.data['easements'])})")
                    self._build_easements_tab()

        ttk.Button(add_frame, text="Add to Easements", command=add_easement).pack(side=tk.RIGHT)

    # -------------------------------------------------------------------------
    # TAB 3: LEASES BUILDER
    # -------------------------------------------------------------------------
    def _build_leases_tab(self):
        for widget in self.tab_leases.winfo_children():
            widget.destroy()

        ttk.Label(self.tab_leases, text="Detected Oil & Gas Leases to write as 'BookType Vol/Pg':", font=("Helvetica", 11, "bold")).pack(anchor="w", pady=(0, 5))

        self.lease_vars = []
        list_frame = ttk.Frame(self.tab_leases)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=4)

        if not self.data["leases"]:
            ttk.Label(list_frame, text="No Oil & Gas Leases currently selected (will write '1) None').", font=("Helvetica", 11, "italic")).pack(anchor="w", pady=5)
        else:
            for idx, l in enumerate(self.data["leases"]):
                var = tk.BooleanVar(value=l.get("included", True))
                self.lease_vars.append((var, l))
                r = l["row"]
                cb = ttk.Checkbutton(list_frame, text=f"Row {r.get('row_idx', '?')}: {l['summary']} ({r.get('grantor', '')} ➔ {r.get('grantee', '')})", variable=var)
                cb.pack(anchor="w", pady=3)

        # Row Selector to add ANY runsheet row
        add_frame = ttk.LabelFrame(self.tab_leases, text="➕ Add Any Other Row from Runsheet as O&G Lease", padding=8)
        add_frame.pack(fill=tk.X, pady=(10, 0))

        row_options = [self._format_row_label(r) for r in self.all_rows]
        add_cb_var = tk.StringVar()
        add_cb = ttk.Combobox(add_frame, textvariable=add_cb_var, values=row_options, font=("Helvetica", 10), state="readonly")
        add_cb.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        if row_options: add_cb.current(0)

        def add_lease():
            sel = add_cb.current()
            if 0 <= sel < len(self.all_rows):
                r = self.all_rows[sel]
                summary = format_encumbrance_short(r)
                if not any(l["row"].get("row_idx") == r.get("row_idx") for l in self.data["leases"]):
                    self.data["leases"].append({"row": r, "summary": summary, "included": True, "status": "Active"})
                    new_parsed = parse_lease_details(r)
                    self.parsed_leases.append(new_parsed)
                    self.data["primary_lease"] = new_parsed
                    self.notebook.tab(self.tab_leases, text=f"🛢️ O&G Leases ({len(self.data['leases'])})")
                    self._build_leases_tab()

        ttk.Button(add_frame, text="Add to Leases", command=add_lease).pack(side=tk.RIGHT)

    # -------------------------------------------------------------------------
    # TAB 4: MORTGAGES BUILDER
    # -------------------------------------------------------------------------
    def _build_mortgages_tab(self):
        for widget in self.tab_mortgages.winfo_children():
            widget.destroy()

        ttk.Label(self.tab_mortgages, text="Detected Mortgages (Satisfied mortgages are unchecked by default):", font=("Helvetica", 11, "bold")).pack(anchor="w", pady=(0, 5))

        self.mortgage_vars = []
        list_frame = ttk.Frame(self.tab_mortgages)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=4)

        if not self.data["mortgages"]:
            ttk.Label(list_frame, text="No Mortgages currently selected (will write '1) None').", font=("Helvetica", 11, "italic")).pack(anchor="w", pady=5)
        else:
            for idx, m in enumerate(self.data["mortgages"]):
                var = tk.BooleanVar(value=m.get("included", False))
                self.mortgage_vars.append((var, m))
                r = m["row"]
                tag = "⭐ [UNRELEASED]" if not m.get("is_satisfied") else " [Satisfied / Released]"
                cb = ttk.Checkbutton(list_frame, text=f"Row {r.get('row_idx', '?')} {tag}: {m['summary']} ({r.get('grantor', '')} ➔ {r.get('grantee', '')})", variable=var)
                cb.pack(anchor="w", pady=3)

        # Row Selector to add ANY runsheet row
        add_frame = ttk.LabelFrame(self.tab_mortgages, text="➕ Add Any Other Row from Runsheet as Mortgage", padding=8)
        add_frame.pack(fill=tk.X, pady=(10, 0))

        row_options = [self._format_row_label(r) for r in self.all_rows]
        add_cb_var = tk.StringVar()
        add_cb = ttk.Combobox(add_frame, textvariable=add_cb_var, values=row_options, font=("Helvetica", 10), state="readonly")
        add_cb.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        if row_options: add_cb.current(0)

        def add_mortgage():
            sel = add_cb.current()
            if 0 <= sel < len(self.all_rows):
                r = self.all_rows[sel]
                summary = format_encumbrance_short(r)
                if not any(m["row"].get("row_idx") == r.get("row_idx") for m in self.data["mortgages"]):
                    self.data["mortgages"].append({
                        "row": r,
                        "summary": summary,
                        "is_satisfied": False,
                        "included": True,
                        "status": "Unreleased"
                    })
                    self.notebook.tab(self.tab_mortgages, text=f"🏦 Mortgages ({len(self.data['mortgages'])})")
                    self._build_mortgages_tab()

        ttk.Button(add_frame, text="Add to Mortgages", command=add_mortgage).pack(side=tk.RIGHT)

    # -------------------------------------------------------------------------
    # Apply Changes to Ownership Report
    # -------------------------------------------------------------------------
    def apply_changes(self, open_after=False):
        if not self.or_path or not os.path.exists(self.or_path):
            messagebox.showerror("Error", f"Ownership Report Excel file not found in:\n{self.pid_dir}", parent=self)
            return

        # Update data payload from UI
        self.data["from_date"] = self.from_date_var.get().strip()
        self.data["to_date"] = self.to_date_var.get().strip()

        addrs = []
        if self.so_add1_var.get().strip(): addrs.append(self.so_add1_var.get().strip())
        if self.so_add2_var.get().strip(): addrs.append(self.so_add2_var.get().strip())

        self.data["surface_owner"] = {
            "name": self.so_name_var.get().strip(),
            "tenancy": self.so_tenancy_var.get().strip(),
            "address_lines": addrs,
            "year": self.so_year_var.get().strip(),
            "interest": "1"
        }

        if self.sync_mineral_var.get():
            self.data["mineral_owner"] = dict(self.data["surface_owner"])

        self.data["sole_mineral_owner"] = self.sole_mineral_var.get()
        self.data["leasehold_mode"] = self.leasehold_mode_var.get()
        self.data["delete_notes"] = self.delete_notes_var.get()

        for var, e in self.easement_vars:
            e["included"] = var.get()

        for var, l in self.lease_vars:
            l["included"] = var.get()

        for var, m in self.mortgage_vars:
            m["included"] = var.get()

        try:
            ORCompilerEngine.apply_to_excel(self.or_path, self.data)
            messagebox.showinfo("Success", f"Successfully synced data to Ownership Report:\n{os.path.basename(self.or_path)}", parent=self)

            if open_after:
                try:
                    if os.name == 'nt':
                        os.startfile(self.or_path)
                    else:
                        subprocess.call(('open', self.or_path))
                except Exception: pass

            self.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to write to Ownership Report:\n{e}", parent=self)
