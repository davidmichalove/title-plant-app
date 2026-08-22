import os
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import fitz # PyMuPDF

try:
    from tkinterdnd2 import DND_FILES
except ImportError:
    DND_FILES = None

class PDFCombinerWindow(tk.Toplevel):
    def __init__(self, master, parcel_dir=None, initial_files=None):
        super().__init__(master)
        self.title("📑 PDF Combiner")
        self.geometry("700x600")
        self.minsize(550, 450)
        self.parcel_dir = parcel_dir

        self.pdf_files = []

        self._build_ui()

        if initial_files:
            self.add_files(initial_files)

    def _build_ui(self):
        # 1. Top instructions & File list
        top_frame = ttk.Frame(self, padding=12)
        top_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(top_frame, text="PDF Files to Combine (in order):", font=("Helvetica", 14, "bold")).pack(anchor=tk.W, pady=(0, 5))

        list_container = ttk.Frame(top_frame)
        list_container.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        self.listbox = tk.Listbox(list_container, font=("Helvetica", 13), selectmode=tk.SINGLE, activestyle="none")
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(list_container, orient=tk.VERTICAL, command=self.listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.config(yscrollcommand=scrollbar.set)

        # Drag and Drop support
        if DND_FILES and hasattr(self.listbox, 'drop_target_register'):
            try:
                self.listbox.drop_target_register(DND_FILES)
                self.listbox.dnd_bind('<<Drop>>', self._on_drop)
            except Exception:
                pass

        # Reorder and Action Buttons Bar
        btn_bar = ttk.Frame(top_frame)
        btn_bar.pack(fill=tk.X, pady=(0, 10))

        ttk.Button(btn_bar, text="➕ Add PDFs...", command=self.browse_add_files).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_bar, text="⬆️ Move Up", command=self.move_up).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_bar, text="⬇️ Move Down", command=self.move_down).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_bar, text="❌ Remove", command=self.remove_selected).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_bar, text="Clear All", command=self.clear_all).pack(side=tk.LEFT)

        # 2. Output Settings Frame
        out_frame = ttk.LabelFrame(self, text="Output Settings", padding=12)
        out_frame.pack(fill=tk.X, padx=12, pady=(0, 10))

        # Output Name Box (Auto-takes the first PDF's name)
        name_row = ttk.Frame(out_frame)
        name_row.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(name_row, text="Output File Name:", font=("Helvetica", 12, "bold"), width=16).pack(side=tk.LEFT)
        self.name_var = tk.StringVar(value="")
        self.name_entry = ttk.Entry(name_row, textvariable=self.name_var, font=("Helvetica", 12))
        self.name_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))

        # Destination Folder
        dest_row = ttk.Frame(out_frame)
        dest_row.pack(fill=tk.X)

        ttk.Label(dest_row, text="Save In Folder:", font=("Helvetica", 12, "bold"), width=16).pack(side=tk.LEFT)
        
        default_dest = ""
        if self.parcel_dir and os.path.exists(os.path.join(self.parcel_dir, "DOCS")):
            default_dest = os.path.join(self.parcel_dir, "DOCS")
        elif self.parcel_dir and os.path.exists(self.parcel_dir):
            default_dest = self.parcel_dir

        self.dest_var = tk.StringVar(value=default_dest)
        self.dest_entry = ttk.Entry(dest_row, textvariable=self.dest_var, font=("Helvetica", 11))
        self.dest_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 5))

        ttk.Button(dest_row, text="Browse...", command=self.browse_dest_folder).pack(side=tk.LEFT)

        # 3. Bottom Merge Buttons
        bottom_bar = ttk.Frame(self, padding=12)
        bottom_bar.pack(fill=tk.X)

        self.status_lbl = ttk.Label(bottom_bar, text="0 PDFs ready to combine", font=("Helvetica", 11, "italic"))
        self.status_lbl.pack(side=tk.LEFT)

        btn_combine_open = ttk.Button(bottom_bar, text="🚀 Combine & Open PDF", command=lambda: self.combine_pdfs(open_after=True))
        btn_combine_open.pack(side=tk.RIGHT, padx=(5, 0))

        btn_combine = ttk.Button(bottom_bar, text="💾 Combine & Save", command=lambda: self.combine_pdfs(open_after=False))
        btn_combine.pack(side=tk.RIGHT)

    def _on_drop(self, event):
        files = []
        raw_data = event.data
        if raw_data:
            # Handle macOS/Windows file paths from DnD
            if raw_data.startswith('{') and raw_data.endswith('}'):
                import re
                files = re.findall(r'\{(.*?)\}', raw_data)
            else:
                files = raw_data.split()

        cleaned_files = [f.strip('{}') for f in files if f.strip('{}').lower().endswith('.pdf')]
        if cleaned_files:
            self.add_files(cleaned_files)

    def browse_add_files(self):
        start_dir = self.parcel_dir if self.parcel_dir and os.path.exists(self.parcel_dir) else "/"
        files = filedialog.askopenfilenames(
            title="Select PDF Files to Combine",
            initialdir=start_dir,
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            parent=self
        )
        if files:
            self.add_files(list(files))

    def add_files(self, file_paths):
        first_time = len(self.pdf_files) == 0
        for fp in file_paths:
            if fp not in self.pdf_files and os.path.exists(fp) and fp.lower().endswith(".pdf"):
                self.pdf_files.append(fp)

        # If this is the first file added, automatically take its name for the Name Box!
        if first_time and self.pdf_files:
            first_fn = os.path.basename(self.pdf_files[0])
            self.name_var.set(first_fn)
            if not self.dest_var.get():
                self.dest_var.set(os.path.dirname(self.pdf_files[0]))

        self._refresh_listbox()

    def browse_dest_folder(self):
        curr = self.dest_var.get() or "/"
        f = filedialog.askdirectory(title="Select Destination Folder", initialdir=curr, parent=self)
        if f:
            self.dest_var.set(f)

    def _refresh_listbox(self):
        self.listbox.delete(0, tk.END)
        for idx, fp in enumerate(self.pdf_files, 1):
            fn = os.path.basename(fp)
            self.listbox.insert(tk.END, f" [{idx}] {fn}  ({os.path.dirname(fp)})")

        self.status_lbl.config(text=f"{len(self.pdf_files)} PDFs ready to combine")

    def move_up(self):
        sel = self.listbox.curselection()
        if not sel or sel[0] == 0: return
        idx = sel[0]
        self.pdf_files[idx - 1], self.pdf_files[idx] = self.pdf_files[idx], self.pdf_files[idx - 1]
        self._refresh_listbox()
        self.listbox.select_set(idx - 1)

    def move_down(self):
        sel = self.listbox.curselection()
        if not sel or sel[0] >= len(self.pdf_files) - 1: return
        idx = sel[0]
        self.pdf_files[idx + 1], self.pdf_files[idx] = self.pdf_files[idx], self.pdf_files[idx + 1]
        self._refresh_listbox()
        self.listbox.select_set(idx + 1)

    def remove_selected(self):
        sel = self.listbox.curselection()
        if not sel: return
        idx = sel[0]
        del self.pdf_files[idx]
        self._refresh_listbox()
        if self.pdf_files:
            new_idx = min(idx, len(self.pdf_files) - 1)
            self.listbox.select_set(new_idx)
        else:
            self.name_var.set("")

    def clear_all(self):
        self.pdf_files = []
        self.name_var.set("")
        self._refresh_listbox()

    def combine_pdfs(self, open_after=False):
        if len(self.pdf_files) < 2:
            messagebox.showwarning("Add PDFs", "Please add at least 2 PDF files to combine.", parent=self)
            return

        out_name = self.name_var.get().strip()
        if not out_name:
            messagebox.showerror("File Name Required", "Please enter an output file name.", parent=self)
            return

        if not out_name.lower().endswith(".pdf"):
            out_name += ".pdf"

        dest_dir = self.dest_var.get().strip()
        if not dest_dir or not os.path.exists(dest_dir):
            messagebox.showerror("Destination Required", "Please select a valid destination folder.", parent=self)
            return

        out_path = os.path.join(dest_dir, out_name)

        # Check overwrite
        if os.path.exists(out_path):
            if not messagebox.askyesno("Overwrite Existing File?", f"File '{out_name}' already exists in destination.\nDo you want to overwrite it?", parent=self):
                return

        try:
            merged_doc = fitz.open()
            total_pages = 0
            for p in self.pdf_files:
                doc_in = fitz.open(p)
                merged_doc.insert_pdf(doc_in)
                total_pages += doc_in.page_count
                doc_in.close()

            merged_doc.save(out_path, incremental=False, deflate=True)
            merged_doc.close()

            if open_after:
                try:
                    if os.name == 'nt':
                        os.startfile(out_path)
                    else:
                        subprocess.call(('open', out_path))
                except Exception:
                    pass

            messagebox.showinfo("Success", f"Successfully combined {len(self.pdf_files)} PDFs ({total_pages} total pages) into:\n{out_path}", parent=self)

        except Exception as e:
            messagebox.showerror("Merge Failed", f"Failed to combine PDFs:\n{e}", parent=self)
