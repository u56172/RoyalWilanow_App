import tkinter as tk
from royalapp.controller import AppController

def main():
    root = tk.Tk()
    AppController(root)
    root.mainloop()

if __name__ == '__main__':
    main()
