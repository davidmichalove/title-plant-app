import os
import json
import tkinter as tk
from tkinter import ttk, messagebox
import glob

class StatusOverviewWindow(tk.Toplevel):
    def __init__(self, parent, base_dir):
        super().__init__(parent)
        self.title("Parcel Status Overview")
        self.geometry("600x500")
        self.base_dir = base_dir
        self.status_file = os.path.join(self.base_dir, "parcel_statuses.json")
        
        self.pids = self.get_all_pids()
        self.statuses = self.load_statuses()
        
        # Ensure all PIDs are in the status dictionary
        changed = False
        for pid in self.pids:
            if pid not in self.statuses:
                self.statuses[pid] = "Not Started"
                changed = True
                
        if changed:
            self.save_statuses()
                
        self.build_ui()
        
    def load_statuses(self):
        if os.path.exists(self.status_file):
            try:
                with open(self.status_file, "r") as f:
                    return json.load(f)
            except:
                return {}
        return {}
        
    def save_statuses(self):
        try:
            with open(self.status_file, "w") as f:
                json.dump(self.statuses, f, indent=4)
        except Exception as e:
            messagebox.showerror("Error", f"Could not save statuses: {e}", parent=self)
            
    def get_all_pids(self):
        pids = []
        for item in os.listdir(self.base_dir):
            path = os.path.join(self.base_dir, item)
            if os.path.isdir(path) and item.startswith("PID "):
                pid = item[4:].strip()
                pids.append(pid)
        return sorted(pids)
        
    def build_ui(self):
        columns = ("pid", "status")
        self.tree = ttk.Treeview(self, columns=columns, show="headings", selectmode="extended")
        self.tree.heading("pid", text="Parcel ID (PID)")
        self.tree.heading("status", text="Status")
        
        self.tree.column("pid", width=300, anchor=tk.W)
        self.tree.column("status", width=200, anchor=tk.CENTER)
        
        # Tag configuration for coloring based on status
        self.tree.tag_configure("Not Started", foreground="gray")
        self.tree.tag_configure("In Progress", foreground="blue")
        self.tree.tag_configure("Completed", foreground="green")
        
        self.populate_tree()
        
        scrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Control panel
        control_frame = ttk.Frame(self)
        control_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=10)
        
        ttk.Label(control_frame, text="Change Status:").pack(side=tk.LEFT, padx=5)
        
        self.status_var = tk.StringVar(value="Not Started")
        combo = ttk.Combobox(control_frame, textvariable=self.status_var, state="readonly", values=["Not Started", "In Progress", "Completed"])
        combo.pack(side=tk.LEFT, padx=5)
        
        btn = ttk.Button(control_frame, text="Apply to Selected", command=self.update_status)
        btn.pack(side=tk.LEFT, padx=5)
        
    def populate_tree(self):
        self.tree.delete(*self.tree.get_children())
        for pid in self.pids:
            status = self.statuses.get(pid, "Not Started")
            self.tree.insert("", tk.END, values=(pid, status), tags=(status,))
        
    def update_status(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a parcel from the list first.", parent=self)
            return
            
        new_status = self.status_var.get()
        for item_id in selected:
            pid = self.tree.item(item_id, "values")[0]
            self.statuses[pid] = new_status
            
        self.save_statuses()
        self.populate_tree()
