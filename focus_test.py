import tkinter as tk
from tkinter import ttk

root = tk.Tk()
t = tk.Text(root, height=5)
t.pack()
t.insert("1.0", "Hello ")

def open_popup():
    top = tk.Toplevel(root)
    def insert():
        w = root.focus_lastfor()
        if isinstance(w, tk.Text):
            w.insert(tk.INSERT, "World")
        else:
            print(f"Not text: {w}")
    ttk.Button(top, text="Insert", command=insert).pack()

ttk.Button(root, text="Popup", command=open_popup).pack()
root.after(1000, lambda: t.focus_set())
# root.after(2000, root.destroy)
# root.mainloop()
