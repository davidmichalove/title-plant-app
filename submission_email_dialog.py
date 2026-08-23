import os
import re
import glob
import subprocess
import urllib.parse
import tkinter as tk
from tkinter import ttk, messagebox

import or_compiler_engine

class SubmissionEmailDialog(tk.Toplevel):
    def __init__(self, master, pid_dir, parcel_num=None):
        super().__init__(master)
        self.title("✉️ Completion & Submission Email Generator")
        self.geometry("820x720")
        self.minsize(700, 550)
        self.transient(master)
        self.attributes("-topmost", True)

        self.pid_dir = pid_dir
        self.parcel_num = parcel_num or self._extract_parcel_num()

        # Extract initial data
        self._load_parcel_data()
        self._build_ui()
        self.update_preview()

    def _extract_parcel_num(self):
        if not self.pid_dir: return "42-00000.000"
        m = re.search(r'PID\s*([0-9\-\.]+)', os.path.basename(self.pid_dir), re.IGNORECASE)
        return m.group(1).strip() if m else "42-00000.000"

    def _load_parcel_data(self):
        # 1. Project Name from folder
        self.project_name = "Norma North I, II"
        if self.pid_dir:
            m_proj = re.search(r'\((.*?)\)', os.path.basename(self.pid_dir))
            if m_proj:
                self.project_name = m_proj.group(1).strip()

        # 2. Location & Acreage defaults
        self.twp = "Warren"
        self.t_num = "8N"
        self.r_num = "6W"
        self.sec_num = "21"
        self.village = "Within the Village of Barnesville"
        self.subdivision_lot = "Lot 143 of the Shoe Factory Addition (Subdivision)"
        self.acreage = "0.134068"
        self.encumbrance_text = "None of record."

        # Query GIS and OR compiler
        try:
            gis_info = or_compiler_engine.get_gis_owner_info(self.parcel_num)
            raw = gis_info.get("raw", {})
            if raw:
                desc = raw.get("desc_", "")
                if desc:
                    self.subdivision_lot = desc.title()
                ac = raw.get("acres", "") or raw.get("calcacres", "")
                if ac and ac != "0":
                    try:
                        self.acreage = f"{float(ac):.6f}"
                    except: pass

            # Run compiler engine for leases
            data = or_compiler_engine.ORCompilerEngine.compile_data(self.pid_dir, self.parcel_num)
            if data:
                p_l = data.get("primary_lease")
                if p_l:
                    is_notice = "notice" in p_l["row"].get("itype", "").lower() or "notice" in p_l["row"].get("comments", "").lower()
                    type_str = "Oil and Gas Lease via Notice" if is_notice else "Oil and Gas Lease"
                    memo_str = f" ({p_l['bk_pg']})"
                    exp_str = f" expires {p_l['exp_date']}" if p_l.get("exp_date") else ""
                    term_str = f" Includes a {p_l['term']}." if p_l.get("term") else ""
                    self.encumbrance_text = f"- {type_str}{memo_str}{exp_str} (no release found of record).{term_str}"
                elif data.get("leases"):
                    lines = []
                    for l in data["leases"]:
                        r = l["row"]
                        lines.append(f"- Oil and Gas Lease ({r['btype']} {r['vol']}/{r['pg']}) (no release found of record).")
                    self.encumbrance_text = "\n".join(lines)
        except Exception as e:
            print("Error loading data for email:", e)

    def _build_ui(self):
        main_frame = ttk.Frame(self, padding=12)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Top Parameters Frame
        top_grid = ttk.LabelFrame(main_frame, text="⚙️ Email Parameters", padding=10)
        top_grid.pack(fill=tk.X, pady=(0, 10))

        # Row 0: Recipient & Project Name
        ttk.Label(top_grid, text="Recipient Name:", font=("Helvetica", 10, "bold")).grid(row=0, column=0, sticky="w", pady=3)
        self.recipient_var = tk.StringVar(value="Tawnie,")
        ttk.Entry(top_grid, textvariable=self.recipient_var, width=15, font=("Helvetica", 10)).grid(row=0, column=1, sticky="w", padx=(5, 15), pady=3)

        ttk.Label(top_grid, text="Project Name:", font=("Helvetica", 10, "bold")).grid(row=0, column=2, sticky="w", pady=3)
        self.project_var = tk.StringVar(value=self.project_name)
        ttk.Entry(top_grid, textvariable=self.project_var, width=25, font=("Helvetica", 10)).grid(row=0, column=3, sticky="ew", padx=(5, 0), pady=3)

        # Row 1: Township / Range / Section
        ttk.Label(top_grid, text="Township / Range:", font=("Helvetica", 10, "bold")).grid(row=1, column=0, sticky="w", pady=3)
        self.twp_var = tk.StringVar(value=f"T. {self.t_num}., R. {self.r_num}., {self.twp}")
        ttk.Entry(top_grid, textvariable=self.twp_var, width=22, font=("Helvetica", 10)).grid(row=1, column=1, sticky="w", padx=(5, 15), pady=3)

        ttk.Label(top_grid, text="Municipality / Village:", font=("Helvetica", 10, "bold")).grid(row=1, column=2, sticky="w", pady=3)
        self.village_var = tk.StringVar(value=self.village)
        ttk.Entry(top_grid, textvariable=self.village_var, width=25, font=("Helvetica", 10)).grid(row=1, column=3, sticky="ew", padx=(5, 0), pady=3)

        # Row 2: Section & Lot / Acreage
        ttk.Label(top_grid, text="Sec / Subdivision / Lot:", font=("Helvetica", 10, "bold")).grid(row=2, column=0, sticky="w", pady=3)
        self.lot_var = tk.StringVar(value=f"Sec. {self.sec_num}: {self.subdivision_lot}")
        ttk.Entry(top_grid, textvariable=self.lot_var, width=35, font=("Helvetica", 10)).grid(row=2, column=1, columnspan=3, sticky="ew", padx=(5, 0), pady=3)

        # Row 3: Acreage & Days Billed
        ttk.Label(top_grid, text="Containing Acreage:", font=("Helvetica", 10, "bold")).grid(row=3, column=0, sticky="w", pady=3)
        self.acres_var = tk.StringVar(value=f"{self.acreage} acres, more or less")
        ttk.Entry(top_grid, textvariable=self.acres_var, width=22, font=("Helvetica", 10)).grid(row=3, column=1, sticky="w", padx=(5, 15), pady=3)

        ttk.Label(top_grid, text="Days / Hours Billed:", font=("Helvetica", 10, "bold")).grid(row=3, column=2, sticky="w", pady=3)
        self.billed_var = tk.StringVar(value="6 hrs")
        b_frame = ttk.Frame(top_grid)
        b_frame.grid(row=3, column=3, sticky="w", padx=(5, 0), pady=3)
        ttk.Entry(b_frame, textvariable=self.billed_var, width=10, font=("Helvetica", 10)).pack(side=tk.LEFT)
        for h in ["4 hrs", "6 hrs", "8 hrs", "1 Day"]:
            ttk.Button(b_frame, text=h, width=5, command=lambda val=h: [self.billed_var.set(val), self.update_preview()]).pack(side=tk.LEFT, padx=2)

        # Row 4: Prior Title & Parse Used
        ttk.Label(top_grid, text="Prior Title Used:", font=("Helvetica", 10, "bold")).grid(row=4, column=0, sticky="w", pady=3)
        self.prior_title_var = tk.StringVar(value="Yes.")
        cb_pt = ttk.Combobox(top_grid, textvariable=self.prior_title_var, values=["Yes.", "No.", "None."], width=12, font=("Helvetica", 10))
        cb_pt.grid(row=4, column=1, sticky="w", padx=(5, 15), pady=3)

        ttk.Label(top_grid, text="Parse Used:", font=("Helvetica", 10, "bold")).grid(row=4, column=2, sticky="w", pady=3)
        self.parse_used_var = tk.StringVar(value="Yes. The application functioned smoothly.")
        cb_parse = ttk.Combobox(top_grid, textvariable=self.parse_used_var, values=["Yes. The application functioned smoothly.", "No.", "Yes."], width=35, font=("Helvetica", 10))
        cb_parse.grid(row=4, column=3, sticky="ew", padx=(5, 0), pady=3)

        top_grid.columnconfigure(3, weight=1)

        # Trace variable changes to auto update preview
        for v in [self.recipient_var, self.project_var, self.twp_var, self.village_var, self.lot_var, self.acres_var, self.billed_var, self.prior_title_var, self.parse_used_var]:
            v.trace_add("write", lambda *args: self.update_preview())

        # Email Text Preview Frame
        preview_frame = ttk.LabelFrame(main_frame, text="📄 Live Email Preview (Edit directly or copy)", padding=10)
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.text_area = tk.Text(preview_frame, wrap=tk.WORD, font=("Monaco", 12), padx=8, pady=8)
        self.text_area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.text_area.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.text_area.config(yscrollcommand=scrollbar.set)

        # Bottom Action Bar
        bottom_bar = ttk.Frame(main_frame, padding=5)
        bottom_bar.pack(fill=tk.X, pady=(5, 0))

        self.copy_btn = ttk.Button(bottom_bar, text="📋 Copy to Clipboard", command=self.copy_to_clipboard)
        self.copy_btn.pack(side=tk.LEFT, padx=(0, 5))

        ttk.Button(bottom_bar, text="✉️ Open in Default Mail Client", command=self.open_in_mail_client).pack(side=tk.LEFT, padx=5)
        ttk.Button(bottom_bar, text="💾 Save to Folder", command=self.save_to_file).pack(side=tk.LEFT, padx=5)

        ttk.Button(bottom_bar, text="Close (Esc)", command=self.destroy).pack(side=tk.RIGHT)

        self.bind("<Escape>", lambda e: self.destroy())

    def update_preview(self):
        recip = self.recipient_var.get().strip()
        proj = self.project_var.get().strip()
        twp = self.twp_var.get().strip()
        village = self.village_var.get().strip()
        lot = self.lot_var.get().strip()
        acres = self.acres_var.get().strip()
        pt = self.prior_title_var.get().strip()
        parse = self.parse_used_var.get().strip()
        billed = self.billed_var.get().strip()

        village_line = f"  {village}\n" if village else ""

        body = f"""{recip}

PID {self.parcel_num} ({proj}) is complete and ready for review.

    * Subdivision Name/Lot:
      {twp}
{village_line}      {lot}
      Containing {acres}

    * Prior Title used: {pt}

    * Parse Used: {parse}

    * Days Billed: {billed}

Encumbrances:
{self.encumbrance_text}

Best,
David Michalove"""

        # Save cursor position if any
        self.text_area.delete("1.0", tk.END)
        self.text_area.insert("1.0", body)

    def copy_to_clipboard(self):
        text = self.text_area.get("1.0", tk.END).strip()
        self.clipboard_clear()
        self.clipboard_append(text)
        self.copy_btn.config(text="✅ Copied to Clipboard!")
        self.after(2000, lambda: self.copy_btn.config(text="📋 Copy to Clipboard"))

    def open_in_mail_client(self):
        subject = f"PID {self.parcel_num} ({self.project_var.get().strip()}) - Complete"
        body = self.text_area.get("1.0", tk.END).strip()
        
        # Build mailto URI
        mailto_url = f"mailto:Tawnie.Rizzardo@gmail.com?subject={urllib.parse.quote(subject)}&body={urllib.parse.quote(body)}"
        try:
            if os.name == 'posix':
                subprocess.Popen(['open', mailto_url])
            else:
                os.startfile(mailto_url)
        except Exception as e:
            messagebox.showerror("Error", f"Could not launch mail client: {e}", parent=self)

    def save_to_file(self):
        if not self.pid_dir or not os.path.exists(self.pid_dir):
            messagebox.showerror("Error", "PID directory not found", parent=self)
            return
        target_path = os.path.join(self.pid_dir, f"PID {self.parcel_num} Submission Email.txt")
        try:
            with open(target_path, "w") as f:
                f.write(self.text_area.get("1.0", tk.END).strip())
            messagebox.showinfo("Saved", f"Saved submission email to:\n{os.path.basename(target_path)}", parent=self)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save email: {e}", parent=self)
