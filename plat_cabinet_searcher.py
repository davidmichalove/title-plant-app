import os
import shutil
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox

PLAT_BASE_DIR = "/Volumes/davidlls/belcogis_plats"

class PlatIndex:
    _cached_plats = None

    @classmethod
    def get_plats(cls, refresh=False):
        if cls._cached_plats is not None and not refresh:
            return cls._cached_plats

        plats = []
        if not os.path.exists(PLAT_BASE_DIR):
            cls._cached_plats = []
            return cls._cached_plats

        for root, dirs, files in os.walk(PLAT_BASE_DIR):
            cab_raw = os.path.basename(root)
            cab = cab_raw.replace("_", " ")
            for f in files:
                if f.endswith(".pdf") and not f.startswith("."):
                    full_path = os.path.join(root, f)
                    name_no_ext = f[:-4]
                    parts = name_no_ext.split("_", 1)
                    slide_code = parts[0]
                    title = parts[1].replace("_", " ") if len(parts) > 1 else name_no_ext

                    plats.append({
                        "cabinet": cab,
                        "slide": slide_code,
                        "title": title,
                        "filename": f,
                        "path": full_path
                    })

        # Sort by cabinet then slide
        plats.sort(key=lambda x: (x["cabinet"], x["slide"], x["title"]))
        cls._cached_plats = plats
        return cls._cached_plats

class PlatCabinetSearchWindow(tk.Toplevel):
    def __init__(self, master, parcel_dir=None):
        super().__init__(master)
        self.title("📑 Belmont County Plat Cabinet Searcher")
        self.geometry("950x650")
        self.minsize(700, 450)
        self.parcel_dir = parcel_dir

        self.plats = PlatIndex.get_plats()
        self.filtered_plats = list(self.plats)

        self._build_ui()
        self._populate_tree()
        self.search_entry.focus_set()

    def _build_ui(self):
        # 1. Search & Filter Bar
        filter_frame = ttk.LabelFrame(self, text="Search & Filter Plats", padding=10)
        filter_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(filter_frame, text="Search:", font=("Helvetica", 13, "bold")).pack(side=tk.LEFT, padx=(0, 5))
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(filter_frame, textvariable=self.search_var, font=("Helvetica", 13), width=35)
        self.search_entry.pack(side=tk.LEFT, padx=(0, 15), fill=tk.X, expand=True)
        self.search_entry.bind("<KeyRelease>", self.on_search)

        ttk.Label(filter_frame, text="Cabinet:").pack(side=tk.LEFT, padx=(0, 5))
        self.cab_var = tk.StringVar(value="All Cabinets")
        cabs = ["All Cabinets", "Cabinet A", "Cabinet B", "Cabinet C", "Cabinet D", "Cabinet E", "Cabinet F", "plat"]
        self.cab_cb = ttk.Combobox(filter_frame, textvariable=self.cab_var, values=cabs, state="readonly", width=14)
        self.cab_cb.pack(side=tk.LEFT, padx=(0, 10))
        self.cab_cb.bind("<<ComboboxSelected>>", self.on_search)

        btn_clear = ttk.Button(filter_frame, text="Clear", command=self.clear_search)
        btn_clear.pack(side=tk.LEFT)

        # 2. Treeview Results Frame
        tree_frame = ttk.Frame(self)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        cols = ("cabinet", "slide", "title", "filename")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings", selectmode="browse")
        self.tree.heading("cabinet", text="Cabinet", command=lambda: self.sort_tree("cabinet"))
        self.tree.heading("slide", text="Slide / Number", command=lambda: self.sort_tree("slide"))
        self.tree.heading("title", text="Subdivision / Plat Name", command=lambda: self.sort_tree("title"))
        self.tree.heading("filename", text="File Name", command=lambda: self.sort_tree("filename"))

        self.tree.column("cabinet", width=110, anchor=tk.W)
        self.tree.column("slide", width=120, anchor=tk.W)
        self.tree.column("title", width=420, anchor=tk.W)
        self.tree.column("filename", width=260, anchor=tk.W)

        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        self.tree.bind("<Double-1>", lambda e: self.open_selected_pdf())
        self.tree.bind("<Return>", lambda e: self.open_selected_pdf())
        self.bind("<Escape>", lambda e: self.destroy())

        # 3. Bottom Actions & Status
        bottom_frame = ttk.Frame(self, padding=10)
        bottom_frame.pack(fill=tk.X, padx=10, pady=5)

        self.status_lbl = ttk.Label(bottom_frame, text=f"Total: {len(self.plats)} plats available", font=("Helvetica", 11, "italic"))
        self.status_lbl.pack(side=tk.LEFT)

        btn_open = ttk.Button(bottom_frame, text="📂 Open PDF (Enter)", command=self.open_selected_pdf)
        btn_open.pack(side=tk.RIGHT, padx=(5, 0))

        if self.parcel_dir and os.path.exists(self.parcel_dir):
            btn_copy_docs = ttk.Button(bottom_frame, text="📥 Copy to DOCS", command=lambda: self.copy_to_parcel("DOCS"))
            btn_copy_docs.pack(side=tk.RIGHT, padx=(5, 0))

            btn_copy_maps = ttk.Button(bottom_frame, text="🗺️ Copy to MAPS", command=lambda: self.copy_to_parcel("MAPS"))
            btn_copy_maps.pack(side=tk.RIGHT, padx=(5, 0))

        btn_reveal = ttk.Button(bottom_frame, text="🔍 Reveal in Finder", command=self.reveal_in_finder)
        btn_reveal.pack(side=tk.RIGHT, padx=(5, 0))

    def on_search(self, event=None):
        q = self.search_var.get().strip().lower()
        cab = self.cab_var.get()

        results = []
        for p in self.plats:
            if cab != "All Cabinets" and p["cabinet"].lower() != cab.lower():
                continue
            if not q or q in p["title"].lower() or q in p["slide"].lower() or q in p["filename"].lower() or q in p["cabinet"].lower():
                results.append(p)

        self.filtered_plats = results
        self._populate_tree()

    def clear_search(self):
        self.search_var.set("")
        self.cab_var.set("All Cabinets")
        self.filtered_plats = list(self.plats)
        self._populate_tree()
        self.search_entry.focus_set()

    def _populate_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        for idx, p in enumerate(self.filtered_plats):
            self.tree.insert("", tk.END, iid=str(idx), values=(p["cabinet"], p["slide"], p["title"], p["filename"]))

        self.status_lbl.config(text=f"Showing {len(self.filtered_plats)} of {len(self.plats)} plats")

    def sort_tree(self, col):
        rev = getattr(self, f"_sort_rev_{col}", False)
        self.filtered_plats.sort(key=lambda x: x[col].lower(), reverse=not rev)
        setattr(self, f"_sort_rev_{col}", not rev)
        self._populate_tree()

    def get_selected_plat(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Select Plat", "Please select a plat from the list first.", parent=self)
            return None
        idx = int(sel[0])
        return self.filtered_plats[idx]

    def open_selected_pdf(self):
        plat = self.get_selected_plat()
        if not plat: return

        path = plat["path"]
        if not os.path.exists(path):
            messagebox.showerror("Error", f"Plat file not found:\n{path}", parent=self)
            return

        try:
            if os.name == 'nt':
                os.startfile(path)
            else:
                subprocess.call(('open', path))
        except Exception as e:
            messagebox.showerror("Error", f"Could not open PDF:\n{e}", parent=self)

    def reveal_in_finder(self):
        plat = self.get_selected_plat()
        if not plat: return

        path = plat["path"]
        try:
            if os.name != 'nt':
                subprocess.call(['open', '-R', path])
        except Exception as e:
            messagebox.showerror("Error", f"Could not reveal file in Finder:\n{e}", parent=self)

    def copy_to_parcel(self, subfolder):
        plat = self.get_selected_plat()
        if not plat or not self.parcel_dir: return

        target_dir = os.path.join(self.parcel_dir, subfolder)
        os.makedirs(target_dir, exist_ok=True)
        dest = os.path.join(target_dir, plat["filename"])

        try:
            shutil.copy2(plat["path"], dest)
            messagebox.showinfo("Success", f"Copied {plat['filename']} to:\n{target_dir}", parent=self)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to copy plat:\n{e}", parent=self)
