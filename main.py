# main.py
import tkinter as tk
from insert_data import InsertDataApp
from search_data import SearchDataApp
from flotta_search import FlottaSearchApp

class MainApp:
    def __init__(self, root):
        self.root = root
        self.root.title("RCTOPCAR B2B DATABASE")
        self.root.geometry("300x200")

        self.file_path = "data.xlsx"

        insert_button = tk.Button(self.root, text="Inserisci Dati", command=self.open_insert_window, width=20, height=2)
        insert_button.pack(pady=10)

        search_button = tk.Button(self.root, text="Aggiorna Dati", command=self.open_search_window, width=20, height=2)
        search_button.pack(pady=10)

        flotta_button = tk.Button(self.root, text="Genera Flusso Flotte", command=self.open_flotta_window, width=20, height=2)
        flotta_button.pack(pady=10)

    def open_insert_window(self):
        self.insert_window = tk.Toplevel(self.root)
        InsertDataApp(self.insert_window, self.file_path)

    def open_search_window(self):
        self.search_window = tk.Toplevel(self.root)
        SearchDataApp(self.search_window, self.file_path)

    def open_flotta_window(self):
        self.flotta_window = tk.Toplevel(self.root)
        FlottaSearchApp(self.flotta_window, self.file_path)

if __name__ == "__main__":
    root = tk.Tk()
    app = MainApp(root)
    root.mainloop()
