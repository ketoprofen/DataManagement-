# search_data.py
import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd
from datetime import datetime
import numpy as np
import os

class SearchDataApp:
    def __init__(self, root, file_path):
        self.root = root
        self.root.title("Search Data")
        self.root.geometry("800x600")
        self.file_path = file_path

        self.labels = [
            "FLOTTA", "TARGA", "MODELLO", "ENTRATA", "PREV.USCITA", "FER. VET", "DITTA",
            "GG.INIZ.MECC", "INIZIO.MECC", "GG.LAV.MECC", "FINE MECC",
            "GG.INIZIO.CARR.", "INIZIO CARR", "GG.LAV.CAR", "FINE CARR",
            "PZ CARR", "STATO", "DOWN TIME", "DATA ULTIMA ATTIVITA'", "RICAMBI", "FERMO TECNICO"
        ]

        self.editable_labels = [
            "FLOTTA", "TARGA", "MODELLO", "ENTRATA", "DITTA", "INIZIO.MECC", "FINE MECC",
            "INIZIO CARR", "FINE CARR", "PZ CARR", "STATO", "RICAMBI"
        ]

        search_frame = tk.Frame(self.root)
        search_frame.pack(pady=20)

        tk.Label(search_frame, text="Enter TARGA:").grid(row=0, column=0, padx=10, pady=10)
        self.targa_entry = tk.Entry(search_frame)
        self.targa_entry.grid(row=0, column=1, padx=10, pady=10)
        search_button = tk.Button(search_frame, text="Search", command=self.display_search_results)
        search_button.grid(row=0, column=2, padx=10, pady=10)

        self.tree = ttk.Treeview(self.root, columns=self.labels, show='headings')
        for col in self.labels:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)
        
        self.tree.pack(fill=tk.BOTH, expand=True)
        self.tree.bind("<Double-1>", self.on_double_click)

    def display_search_results(self):
        targa = self.targa_entry.get()
        if os.path.exists(self.file_path):
            df = pd.read_excel(self.file_path)
            results = df[df["TARGA"].str.contains(targa, case=False, na=False)]
            if not results.empty:
                self.tree.delete(*self.tree.get_children())
                for _, row in results.iterrows():
                    self.tree.insert("", tk.END, values=list(row))
            else:
                messagebox.showinfo("Info", "No results found!")
        else:
            messagebox.showinfo("Info", "No data file found!")

    def on_double_click(self, event):
        item = self.tree.selection()[0]
        self.edit_window = tk.Toplevel(self.root)
        self.edit_window.title("Edit Row")
        self.entries = {}
        row_data = self.tree.item(item, "values")

        for i, col in enumerate(self.labels):
            if col in self.editable_labels:
                lbl = tk.Label(self.edit_window, text=col)
                lbl.grid(row=i, column=0, padx=5, pady=5)
                if col == "STATO":
                    entry = ttk.Combobox(self.edit_window, values=[
                        "ATT.PERZ.", "ATT.AUT.", "ATT.RIC.", "LAV.CAR1", "LAV.CAR2", "LAV.CAR3",
                        "LAV.CAR4", "LAV.MECC.", "FIN", "ALTRI LAVORI", "DA FATTURARE", "PRONTA", "PRE-CONSEGNA"
                    ], width=47)
                    entry.set(row_data[i])
                elif col == "RICAMBI":
                    entry = ttk.Combobox(self.edit_window, values=["SÌ", "NO"], width=47)
                    entry.set(row_data[i])
                else:
                    entry = tk.Entry(self.edit_window, width=50)
                    entry.insert(0, row_data[i])
                entry.grid(row=i, column=1, padx=5, pady=5)
                self.entries[col] = entry

        self.update_button = tk.Button(self.edit_window, text="Update", command=lambda: self.update_row(item, row_data))
        self.update_button.grid(row=len(self.labels), columnspan=2, pady=10)

    def calculate_business_days(self, start_date, end_date):
        return np.busday_count(start_date.date(), end_date.date())

    def update_row(self, item, original_values):
        updated_values = list(original_values)  # Copy original values to update only editable fields
        for col in self.editable_labels:
            index = self.labels.index(col)
            updated_values[index] = self.entries[col].get()
        
        self.tree.item(item, values=updated_values)
        
        if os.path.exists(self.file_path):
            df = pd.read_excel(self.file_path)
            index = df[(df["TARGA"] == original_values[1]) & (df["ENTRATA"] == original_values[3])].index
            if not index.empty:
                for col in self.editable_labels:
                    df.at[index[0], col] = self.entries[col].get()

                # Calculate additional fields
                entrata = self.entries["ENTRATA"].get()
                inizio_mecc = self.entries["INIZIO.MECC"].get()
                fine_mecc = self.entries["FINE MECC"].get()
                inizio_carr = self.entries["INIZIO CARR"].get()
                fine_carr = self.entries["FINE CARR"].get()
                pz_carr = self.entries["PZ CARR"].get()

                try:
                    entrata_date = datetime.strptime(entrata, "%d/%m/%Y") if entrata else None
                    inizio_mecc_date = datetime.strptime(inizio_mecc, "%d/%m/%Y") if inizio_mecc else None
                    fine_mecc_date = datetime.strptime(fine_mecc, "%d/%m/%Y") if fine_mecc else None
                    inizio_carr_date = datetime.strptime(inizio_carr, "%d/%m/%Y") if inizio_carr else None
                    fine_carr_date = datetime.strptime(fine_carr, "%d/%m/%Y") if fine_carr else None

                    # DATA ULTIMA ATTIVITA'
                    data_ultima_attivita_date = max(fine_mecc_date, fine_carr_date) if fine_mecc_date and fine_carr_date else fine_mecc_date or fine_carr_date
                    df.at[index[0], "DATA ULTIMA ATTIVITA'"] = data_ultima_attivita_date.strftime("%d/%m/%Y") if data_ultima_attivita_date else None

                    # FER. VET
                    if entrata_date and data_ultima_attivita_date:
                        df.at[index[0], "FER. VET"] = self.calculate_business_days(entrata_date, data_ultima_attivita_date)

                    # GG.INIZ.MECC
                    if entrata_date and inizio_mecc_date:
                        df.at[index[0], "GG.INIZ.MECC"] = self.calculate_business_days(entrata_date, inizio_mecc_date)

                    # GG.INIZIO.CARR
                    if entrata_date and inizio_carr_date:
                        df.at[index[0], "GG.INIZIO.CARR."] = self.calculate_business_days(entrata_date, inizio_carr_date)

                    # GG.LAV.MECC
                    if inizio_mecc_date and fine_mecc_date:
                        df.at[index[0], "GG.LAV.MECC"] = self.calculate_business_days(inizio_mecc_date, fine_mecc_date)

                    # GG.LAV.CARR
                    if inizio_carr_date and fine_carr_date:
                        df.at[index[0], "GG.LAV.CAR"] = self.calculate_business_days(inizio_carr_date, fine_carr_date)

                    # DOWN TIME
                    if entrata_date and data_ultima_attivita_date:
                        df.at[index[0], "DOWN TIME"] = self.calculate_business_days(entrata_date, data_ultima_attivita_date)

                    # FERMO TECNICO
                    start_date = min(date for date in [inizio_mecc_date, inizio_carr_date] if date)
                    end_date = max(date for date in [fine_mecc_date, fine_carr_date] if date)
                    if start_date and end_date:
                        df.at[index[0], "FERMO TECNICO"] = self.calculate_business_days(start_date, end_date)

                    # Handle PZ CARR being blank or 0
                    if not pz_carr:
                        df.at[index[0], "INIZIO CARR"] = None
                        df.at[index[0], "FINE CARR"] = None
                        df.at[index[0], "GG.LAV.CAR"] = None

                except Exception as e:
                    messagebox.showerror("Error", f"Error in calculations: {e}")

                # Ensure PZ CARR is stored as a number
                try:
                    df.at[index[0], "PZ CARR"] = int(pz_carr)
                except ValueError:
                    df.at[index[0], "PZ CARR"] = None

                df.to_excel(self.file_path, index=False)
                messagebox.showinfo("Info", "Data updated successfully!")
        
        self.edit_window.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = SearchDataApp(root, "data.xlsx")
    root.mainloop()

