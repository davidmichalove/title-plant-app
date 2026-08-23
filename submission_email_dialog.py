import os
import re
import glob
import webbrowser
import urllib.parse
import tkinter as tk
from tkinter import ttk, messagebox
import openpyxl

import or_compiler_engine

class SubmissionEmailDialog(tk.Toplevel):
    def __init__(self, master, pid_dir, parcel_num=None):
        super().__init__(master)
        self.title("✉️ Completion & Submission Email Generator")
        self.geometry("860x780")
        self.minsize(740, 580)
        self.transient(master)
        self.attributes("-topmost", True)

        self.pid_dir = pid_dir
        self.parcel_num = parcel_num or self._extract_parcel_num()

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
        self.is_village = True
        self.village = "Within the Village of Barnesville"
        self.is_subdivision = False
        self.subdiv_name = ""
        self.lot_number = ""
        self.subdivision_lot_text = "Lot 143 of the Shoe Factory Addition (Subdivision)"
        self.acreage = "0.134068"
        self.encumbrance_text = "- None of record."

        # Query GIS
        try:
            gis_info = or_compiler_engine.get_gis_owner_info(self.parcel_num)
            raw = gis_info.get("raw", {})
            if raw:
                # Township / Sec / T / R
                raw_twp = raw.get("twp", "")
                if raw_twp.upper() == "WAR": self.twp = "Warren"
                elif raw_twp.upper() == "MEA": self.twp = "Mead"
                elif raw_twp.upper() == "SOM": self.twp = "Somerset"
                elif raw_twp.upper() == "WAS": self.twp = "Washington"
                elif raw_twp.upper() == "YOR": self.twp = "York"
                elif raw_twp: self.twp = raw_twp.title()

                if raw.get("t"): self.t_num = f"{raw.get('t')}N"
                if raw.get("r"): self.r_num = f"{raw.get('r')}W"
                if raw.get("sec"): self.sec_num = str(raw.get("sec"))

                # Municipality / Village
                polsub = raw.get("polsub", "")
                if polsub and "BARNE" in polsub.upper():
                    self.is_village = True
                    self.village = "Within the Village of Barnesville"
                elif polsub:
                    self.is_village = True
                    self.village = f"Within the {polsub.title()}"
                else:
                    self.is_village = False

                # Subdivision detection
                subdiv = raw.get("subdiv", "").strip()
                parcel_lot = raw.get("parcel", "").strip()
                if subdiv and subdiv.upper() != "NONE" and subdiv.upper() != "NAN":
                    self.is_subdivision = True
                    subdiv_clean = re.sub(r'\(MAP\)', '', subdiv, flags=re.IGNORECASE).strip().title()
                    self.subdiv_name = subdiv_clean
                    self.lot_number = parcel_lot or "143"
                    self.subdivision_lot_text = f"Lot {self.lot_number} of the {subdiv_clean} (Subdivision)"
                else:
                    desc = raw.get("desc_", "").strip()
                    if desc:
                        self.subdivision_lot_text = desc.title()

                # Acreage
                ac = raw.get("acres", "") or raw.get("calcacres", "")
                if ac and ac != "0" and ac != "0.00000000":
                    try:
                        self.acreage = f"{float(ac):.6f}"
                    except: pass

            # Try reading exact acreage from OR sheet
            if self.pid_dir:
                or_files = glob.glob(os.path.join(self.pid_dir, "*OR*.xlsx"))
                valid_or = [f for f in or_files if not os.path.basename(f).startswith("~") and not os.path.basename(f).startswith("._") and "backup" not in f.lower()]
                if valid_or:
                    wb = openpyxl.load_workbook(valid_or[0], data_only=True)
                    ws = wb.active
                    or_ac = ws["B7"].value
                    if or_ac and str(or_ac).strip() not in ["0.0", "0", "<ACRES_IN2>"]:
                        self.acreage = str(or_ac).strip()

            # Run compiler engine for leases & encumbrances
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

        # 1. Subject Line Frame
        subj_frame = ttk.LabelFrame(main_frame, text="🏷️ Email Subject Line", padding=10)
        subj_frame.pack(fill=tk.X, pady=(0, 8))

        self.subject_var = tk.StringVar()
        self.subject_entry = ttk.Entry(subj_frame, textvariable=self.subject_var, font=("Helvetica", 11, "bold"))
        self.subject_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(subj_frame, text="📋 Copy Subject", width=14, command=self.copy_subject).pack(side=tk.RIGHT)

        # 2. Parameters Grid Frame
        top_grid = ttk.LabelFrame(main_frame, text="⚙️ Dynamic Parcel Parameters", padding=10)
        top_grid.pack(fill=tk.X, pady=(0, 8))

        # Row 0: Recipient & Project Name
        ttk.Label(top_grid, text="Recipient:", font=("Helvetica", 10, "bold")).grid(row=0, column=0, sticky="w", pady=3)
        self.recipient_var = tk.StringVar(value="Tawnie,")
        ttk.Entry(top_grid, textvariable=self.recipient_var, width=15, font=("Helvetica", 10)).grid(row=0, column=1, sticky="w", padx=(5, 15), pady=3)

        ttk.Label(top_grid, text="Project Name:", font=("Helvetica", 10, "bold")).grid(row=0, column=2, sticky="w", pady=3)
        self.project_var = tk.StringVar(value=self.project_name)
        ttk.Entry(top_grid, textvariable=self.project_var, width=28, font=("Helvetica", 10)).grid(row=0, column=3, sticky="ew", padx=(5, 0), pady=3)

        # Row 1: Township / Range & Village Toggle
        ttk.Label(top_grid, text="Township / Range:", font=("Helvetica", 10, "bold")).grid(row=1, column=0, sticky="w", pady=3)
        self.twp_var = tk.StringVar(value=f"T. {self.t_num}., R. {self.r_num}., {self.twp}")
        ttk.Entry(top_grid, textvariable=self.twp_var, width=22, font=("Helvetica", 10)).grid(row=1, column=1, sticky="w", padx=(5, 15), pady=3)

        # Village Toggle & Entry
        v_frame = ttk.Frame(top_grid)
        v_frame.grid(row=1, column=2, columnspan=2, sticky="ew", pady=3)
        self.is_village_var = tk.BooleanVar(value=self.is_village)
        ttk.Checkbutton(v_frame, text="Include Village/Muni:", variable=self.is_village_var).pack(side=tk.LEFT, padx=(0, 5))
        self.village_var = tk.StringVar(value=self.village)
        ttk.Entry(v_frame, textvariable=self.village_var, font=("Helvetica", 10)).pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Row 2: Subdivision Toggle & Lot/Subdivision Name
        sub_row = ttk.Frame(top_grid)
        sub_row.grid(row=2, column=0, columnspan=4, sticky="ew", pady=3)
        self.is_subdivision_var = tk.BooleanVar(value=self.is_subdivision)
        ttk.Checkbutton(sub_row, text="Platted Subdivision (Adds to Subject):", variable=self.is_subdivision_var).pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Label(sub_row, text="Sec #:").pack(side=tk.LEFT, padx=(5, 2))
        self.sec_var = tk.StringVar(value=f"Sec. {self.sec_num}:")
        ttk.Entry(sub_row, textvariable=self.sec_var, width=9, font=("Helvetica", 10)).pack(side=tk.LEFT, padx=(0, 5))

        self.lot_var = tk.StringVar(value=self.subdivision_lot_text)
        ttk.Entry(sub_row, textvariable=self.lot_var, font=("Helvetica", 10)).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))

        # Row 3: Acreage & Days Billed
        ttk.Label(top_grid, text="Acreage:", font=("Helvetica", 10, "bold")).grid(row=3, column=0, sticky="w", pady=3)
        self.acres_var = tk.StringVar(value=f"Containing {self.acreage} acres, more or less")
        ttk.Entry(top_grid, textvariable=self.acres_var, width=28, font=("Helvetica", 10)).grid(row=3, column=1, sticky="w", padx=(5, 15), pady=3)

        ttk.Label(top_grid, text="Days / Hours Billed:", font=("Helvetica", 10, "bold")).grid(row=3, column=2, sticky="w", pady=3)
        self.billed_var = tk.StringVar(value="6 hrs")
        b_frame = ttk.Frame(top_grid)
        b_frame.grid(row=3, column=3, sticky="w", padx=(5, 0), pady=3)
        ttk.Entry(b_frame, textvariable=self.billed_var, width=9, font=("Helvetica", 10)).pack(side=tk.LEFT)
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

        # Trace variable changes
        for v in [self.recipient_var, self.project_var, self.twp_var, self.village_var, self.is_village_var, 
                  self.is_subdivision_var, self.sec_var, self.lot_var, self.acres_var, self.billed_var, 
                  self.prior_title_var, self.parse_used_var]:
            v.trace_add("write", lambda *args: self.update_preview())

        # 3. Live Email Preview Frame
        preview_frame = ttk.LabelFrame(main_frame, text="📄 Live Email Preview (Editable)", padding=10)
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.text_area = tk.Text(preview_frame, wrap=tk.WORD, font=("Monaco", 12), padx=8, pady=8)
        self.text_area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.text_area.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.text_area.config(yscrollcommand=scrollbar.set)

        # 4. Bottom Action Bar
        bottom_bar = ttk.Frame(main_frame, padding=5)
        bottom_bar.pack(fill=tk.X, pady=(5, 0))

        self.copy_btn = ttk.Button(bottom_bar, text="📋 Copy Full Email Body", command=self.copy_to_clipboard)
        self.copy_btn.pack(side=tk.LEFT, padx=(0, 5))

        ttk.Button(bottom_bar, text="🌐 Open in Gmail (Browser)", command=self.open_in_gmail).pack(side=tk.LEFT, padx=5)
        ttk.Button(bottom_bar, text="✉️ Open in Mail Client", command=self.open_in_mail_client).pack(side=tk.LEFT, padx=5)
        ttk.Button(bottom_bar, text="💾 Save to Folder", command=self.save_to_file).pack(side=tk.LEFT, padx=5)

        ttk.Button(bottom_bar, text="Close (Esc)", command=self.destroy).pack(side=tk.RIGHT)

        self.bind("<Escape>", lambda e: self.destroy())

    def update_preview(self):
        recip = self.recipient_var.get().strip()
        proj = self.project_var.get().strip()
        twp = self.twp_var.get().strip()
        use_village = self.is_village_var.get()
        village = self.village_var.get().strip()
        use_subdiv = self.is_subdivision_var.get()
        sec_prefix = self.sec_var.get().strip()
        lot = self.lot_var.get().strip()
        acres = self.acres_var.get().strip()
        pt = self.prior_title_var.get().strip()
        parse = self.parse_used_var.get().strip()
        billed = self.billed_var.get().strip()

        # Update subject line
        if use_subdiv and lot:
            # Clean subdivision lot description for subject
            subdiv_clean = re.sub(r'\(Subdivision\)', '', lot, flags=re.IGNORECASE).strip()
            subject = f"Abstract Completed: PID {self.parcel_num} ({proj}); {subdiv_clean}"
        else:
            subject = f"Abstract Completed: PID {self.parcel_num} ({proj})"
        self.subject_var.set(subject)

        # Build location block
        loc_lines = [f"      {twp}"]
        if use_village and village:
            loc_lines.append(f"      {village}")
        loc_lines.append(f"      {sec_prefix} {lot}")
        if acres.lower().startswith("containing"):
            loc_lines.append(f"      {acres}")
        else:
            loc_lines.append(f"      Containing {acres}")

        loc_text = "\n".join(loc_lines)

        body = f"""{recip}

PID {self.parcel_num} ({proj}) is complete and ready for review.

    * Subdivision Name/Lot:
{loc_text}

    * Prior Title used: {pt}

    * Parse Used: {parse}

    * Days Billed: {billed}

Encumbrances:
{self.encumbrance_text}

Best,
David Michalove"""

        self.text_area.delete("1.0", tk.END)
        self.text_area.insert("1.0", body)

    def copy_subject(self):
        subj = self.subject_var.get().strip()
        self.clipboard_clear()
        self.clipboard_append(subj)
        messagebox.showinfo("Copied", f"Subject copied to clipboard:\n\n{subj}", parent=self)

    def copy_to_clipboard(self):
        text = self.text_area.get("1.0", tk.END).strip()
        self.clipboard_clear()
        self.clipboard_append(text)
        self.copy_btn.config(text="✅ Copied Body to Clipboard!")
        self.after(2000, lambda: self.copy_btn.config(text="📋 Copy Full Email Body"))

    def open_in_gmail(self):
        subject = self.subject_var.get().strip()
        body = self.text_area.get("1.0", tk.END).strip()
        to_email = "Tawnie.Rizzardo@gmail.com"
        
        gmail_url = f"https://mail.google.com/mail/?view=cm&fs=1&to={urllib.parse.quote(to_email)}&su={urllib.parse.quote(subject)}&body={urllib.parse.quote(body)}"
        try:
            webbrowser.open(gmail_url)
        except Exception as e:
            messagebox.showerror("Error", f"Could not launch browser for Gmail: {e}", parent=self)

    def open_in_mail_client(self):
        subject = self.subject_var.get().strip()
        body = self.text_area.get("1.0", tk.END).strip()
        
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
                f.write(f"Subject: {self.subject_var.get().strip()}\n\n" + self.text_area.get("1.0", tk.END).strip())
            messagebox.showinfo("Saved", f"Saved submission email to:\n{os.path.basename(target_path)}", parent=self)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save email: {e}", parent=self)
