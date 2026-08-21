import tkinter as tk

root = tk.Tk()
t = tk.Text(root, undo=True)
t.pack()
def _undo(e):
    try:
        e.widget.edit_undo()
    except:
        pass
    return "break"
def _redo(e):
    try:
        e.widget.edit_redo()
    except:
        pass
    return "break"
t.bind("<Command-z>", _undo)
t.bind("<Command-Shift-Z>", _redo)
t.insert("1.0", "Type something and hit Cmd-Z")
root.update()
root.after(1000, root.destroy)
root.mainloop()
