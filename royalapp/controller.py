import tkinter as tk
from tkinter import filedialog, messagebox
from royalapp import model


class AppController:
    def __init__(self, master):
        #Inicjalizacja
        self.master = master
        self.master.title('Rejestry Royal Wilanów')
        self.files = []
        self.out_dir= model.get_paths()[0]
        #UI
        self.text = tk.Text(master, width=100, height=30)
        self.text.pack(fill='both', expand=True, padx=8, pady=8)
        bar = tk.Frame(master)
        bar.pack(fill='x', padx=8, pady=(0, 8))
        tk.Button(bar, text='Wybierz pliki', command=self.pick).pack(side='left')
        tk.Button(bar, text='Wybierz folder zapisu', command=self.pick_out).pack(side='left', padx=6)
        tk.Button(bar, text='Przetwórz i zapisz', command=self.run).pack(side='left', padx=6)
        self.text.insert(tk.END, f"Folder zapisu: {self.out_dir}\n")

    def pick(self):
        files = filedialog.askopenfilenames(title='Wybierz Pliki Excel', filetypes=(('Pliki Excel', '*.xlsx *.xls'),))
        if files:
            self.files = list(files)
            self.text.insert(tk.END, 'Wybrane pliki:\n')
            for p in self.files:
                self.text.insert(tk.END, f'• {p}\n')

    def pick_out(self):
        d = filedialog.askdirectory(title='Wybierz folder zapisu')
        if d:
            self.out_dir = d
            self.text.insert(tk.END, f'\nFolder zapisu ustawiony na: {self.out_dir}\n')

    def run(self):
        if not self.files:
            messagebox.showwarning('Brak plików', 'Najpierw wybierz pliki.')
            return

        wynik = model.przetworz(self.files)
        model.zapisz_pliki(wynik, self.out_dir)

        messagebox.showinfo('Gotowe', f'Zapisano raporty i wykresy w:\n{self.out_dir}')


