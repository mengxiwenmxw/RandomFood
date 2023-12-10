import tkinter as tk
from DemoClass import Gui


if __name__ == '__main__':
    root = tk.Tk()
    window = Gui(root)
    window.set_window()
    # window
    root.mainloop()