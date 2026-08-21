import tkinter as tk
from tkinter import ttk

def show_popup():
    root = tk.Tk()
    root.title("Select Documents")
    
    columns = ("Inst", "Type", "Vol/Pg", "Date")
    tree = ttk.Treeview(root, columns=columns, show="headings")
    tree.heading("Inst", text="Instrument")
    tree.heading("Type", text="Type")
    tree.heading("Vol/Pg", text="Vol/Pg")
    tree.heading("Date", text="Recorded Date")
    
    # Mock data
    table_data_json = [
        ["", "Instrument", "Type", "Vol", "Pg", "Date", "Grantor", "Grantee"],
        ["", "20230001", "DEED", "100", "200", "01/01/2023", "Smith", "Jones"]
    ]
    
    for idx, row in enumerate(table_data_json):
        if idx == 0: continue # header
        if len(row) >= 6:
            inst = row[1]
            dtype = row[2]
            volpg = f"{row[3]}/{row[4]}"
            date = row[5]
            tree.insert("", tk.END, values=(inst, dtype, volpg, date))
            
    tree.pack(fill=tk.BOTH, expand=True)
    
    def on_download():
        selected = tree.selection()
        items = [tree.item(s)["values"] for s in selected]
        print("Selected:", items)
        root.destroy()
        
    tk.Button(root, text="Download Selected", command=on_download).pack()
    
    # Auto-click download for test
    def auto_test():
        children = tree.get_children()
        if children:
            tree.selection_set(children[0])
            on_download()
    
    root.after(500, auto_test)
    root.mainloop()

show_popup()
