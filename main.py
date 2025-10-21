import tkinter as tk
from gui import STAFCommentApp  # loads the GUI class from gui.py

# Entry point for the STAF Insert Comment Tool V3.1 (xlwings)
if __name__ == "__main__":
    root = tk.Tk()
    app = STAFCommentApp(root)
    root.mainloop()
