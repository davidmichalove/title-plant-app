import os
import glob
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
from or_compiler_engine import ORCompilerEngine

class ORSyncDialog(tk.Toplevel):
    def __init__(self, master, pid_dir, parcel_num=None):
        super().__init__(master)
        self.title("📊 Ownership Report (OR) Auto-Compiler & Sync")
        self.geometry("820x680")
        self.minsize(700, 500)
        self.transient(master)
        self.attributes("-topmost", True)

        self.pid_dir = pid_dir
        self.parcel_num = parcel_num

        self.data = ORCompilerEngine.compile_data(self.pid_dir, self.parcel_num)
        if not self.data:
            messagebox.showerror("Error", f"Could not find valid Runsheet data in:\n{pid_dir}", parent=self)
            self.destroy()
            return

        self._find_or_file()
        self._build_ui()

    def _find_or_file(self):
        or_files = glob.glob(os.path.join(self.pid_dir, "*OR*.xlsx"))
        valid_or = [f for f in or_files if not os.path.basename(f).startswith("~") and not os.path.basename(f).startswith("._") and "backup" not in f.lower()]
        self.or_path = valid_or[0] if valid_or else None

    def _build_ui(self):
        # 1. Top Header: Date Range
        top_frame = ttk.LabelFrame(self, text="📅 Records Examined Date Range (Cell B62)", padding=10)
        top_frame.pack(fill=tk.X, padx=12, pady=(10, 5))

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
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=12, pady=5)

        tab_owners = ttk.Frame(self.notebook, padding=10)
        tab_easements = ttk.Frame(self.notebook, padding=10)
        tab_leases = ttk.Frame(self.notebook, padding=10)
        tab_mortgages = ttk.Frame(self.notebook, padding=10)

        self.notebook.add(tab_owners, text="🏠 Surface & Mineral Owners")
        self.notebook.add(tab_easements, text=f"🛣️ Easements ({len(self.data['easements'])})")
        self.notebook.add(tab_leases, text=f"🛢️ O&G Leases ({len(self.data['leases'])})")
        self.notebook.add(tab_mortgages, text=f"🏦 Mortgages ({len(self.data['mortgages'])})")

        # --- TAB 1: OWNERS ---
        so = self.data["surface_owner"]
        ttk.Label(tab_owners, text="Surface Owner (Vesting Grantee):", font=("Helvetica", 11, "bold")).grid(row=0, column=0, sticky="w", pady=2)
        self.so_name_var = tk.StringVar(value=so.get("name", ""))
        ttk.Entry(tab_owners, textvariable=self.so_name_var, font=("Helvetica", 11)).grid(row=0, column=1, sticky="ew", padx=(5, 0), pady=2)

        ttk.Label(tab_owners, text="Marital / Vesting Tenancy:", font=("Helvetica", 11, "bold")).grid(row=1, column=0, sticky="w", pady=2)
        self.so_tenancy_var = tk.StringVar(value=so.get("tenancy", ""))
        ttk.Entry(tab_owners, textvariable=self.so_tenancy_var, font=("Helvetica", 11)).grid(row=1, column=1, sticky="ew", padx=(5, 0), pady=2)

        addrs = so.get("address_lines", ["", ""])
        ttk.Label(tab_owners, text="Mailing Address Line 1:", font=("Helvetica", 11, "bold")).grid(row=2, column=0, sticky="w", pady=2)
        self.so_add1_var = tk.StringVar(value=addrs[0] if len(addrs) > 0 else "")
        ttk.Entry(tab_owners, textvariable=self.so_add1_var, font=("Helvetica", 11)).grid(row=2, column=1, sticky="ew", padx=(5, 0), pady=2)

        ttk.Label(tab_owners, text="Mailing Address Line 2:", font=("Helvetica", 11, "bold")).grid(row=3, column=0, sticky="w", pady=2)
        self.so_add2_var = tk.StringVar(value=addrs[1] if len(addrs) > 1 else "")
        ttk.Entry(tab_owners, textvariable=self.so_add2_var, font=("Helvetica", 11)).grid(row=3, column=1, sticky="ew", padx=(5, 0), pady=2)

        ttk.Label(tab_owners, text="Acquired Year:", font=("Helvetica", 11, "bold")).grid(row=4, column=0, sticky="w", pady=2)
        self.so_year_var = tk.StringVar(value=so.get("year", ""))
        ttk.Entry(tab_owners, textvariable=self.so_year_var, font=("Helvetica", 11)).grid(row=4, column=1, sticky="w", padx=(5, 0), pady=2)

        tab_owners.columnconfigure(1, weight=1)

        ttk.Separator(tab_owners, orient="horizontal").grid(row=5, column=0, columnspan=2, sticky="ew", pady=10)

        self.sync_mineral_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(tab_owners, text="Mirror Mineral Ownership to match Surface Owner (100% Fee Simple)", variable=self.sync_mineral_var).grid(row=6, column=0, columnspan=2, sticky="w", pady=5)

        # --- TAB 2: EASEMENTS ---
        ttk.Label(tab_easements, text="Detected Easements & Rights of Way to include in OR:", font=("Helvetica", 11, "bold")).pack(anchor="w", pady=(0, 5))
        self.easement_vars = []
        if not self.data["easements"]:
            ttk.Label(tab_easements, text="No Easements or Rights of Way found (will write '1) None').", font=("Helvetica", 11, "italic")).pack(anchor="w", pady=5)
        else:
            for idx, e in enumerate(self.data["easements"]):
                var = tk.BooleanVar(value=e["included"])
                self.easement_vars.append((var, e))
                cb = ttk.Checkbutton(tab_easements, text=f"Row {e['row']['row_idx']}: {e['summary']}", variable=var)
                cb.pack(anchor="w", pady=3)

        # --- TAB 3: LEASES ---
        ttk.Label(tab_leases, text="Detected Oil & Gas Leases to include in OR:", font=("Helvetica", 11, "bold")).pack(anchor="w", pady=(0, 5))
        self.lease_vars = []
        if not self.data["leases"]:
            ttk.Label(tab_leases, text="No Oil & Gas Leases found (will write '1) None').", font=("Helvetica", 11, "italic")).pack(anchor="w", pady=5)
        else:
            for idx, l in enumerate(self.data["leases"]):
                var = tk.BooleanVar(value=l["included"])
                self.lease_vars.append((var, l))
                cb = ttk.Checkbutton(tab_leases, text=f"Row {l['row']['row_idx']}: {l['summary']}", variable=var)
                cb.pack(anchor="w", pady=3)

        # --- TAB 4: MORTGAGES ---
        ttk.Label(tab_mortgages, text="Detected Mortgages (Satisfied mortgages are unchecked by default):", font=("Helvetica", 11, "bold")).pack(anchor="w", pady=(0, 5))
        self.mortgage_vars = []
        if not self.data["mortgages"]:
            ttk.Label(tab_mortgages, text="No Mortgages found (will write '1) None').", font=("Helvetica", 11, "italic")).pack(anchor="w", pady=5)
        else:
            for idx, m in enumerate(self.data["mortgages"]):
                var = tk.BooleanVar(value=m["included"])
                self.mortgage_vars.append((var, m))
                tag = "⭐ [UNRELEASED]" if not m["is_satisfied"] else " [Satisfied / Released]"
                cb = ttk.Checkbutton(tab_mortgages, text=f"Row {m['row']['row_idx']} {tag}: {m['summary']}", variable=var)
                cb.pack(anchor="w", pady=3)

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
