# insert_data.py
import tkinter as tk
from tkinter import messagebox, ttk
from tkcalendar import DateEntry
import pandas as pd
import numpy as np
import os
from datetime import datetime, timedelta

class InsertDataApp:
    def __init__(self, root, file_path):
        self.root = root
        self.root.title("Insert Data")
        self.root.geometry("600x400")
        self.file_path = file_path

        self.gui_labels = [
            "FLOTTA", "TARGA", "MODELLO", "ENTRATA", "DITTA", "PZ CARR", "STATO", "RICAMBI"
        ]

        self.excel_labels = [
            "FLOTTA", "TARGA", "MODELLO", "ENTRATA", "PREV.USCITA", "FER. VET", "DITTA",
            "GG.INIZ.MECC", "INIZIO.MECC", "GG.LAV.MECC", "FINE MECC",
            "GG.INIZIO.CARR.", "INIZIO CARR", "GG.LAV.CAR", "FINE CARR",
            "PZ CARR", "STATO", "DOWN TIME", "DATA ULTIMA ATTIVITA'", "RICAMBI", "FERMO TECNICO"
        ]

        self.entries = {}
        self.create_widgets()

    def create_widgets(self):
        form_frame = tk.Frame(self.root)
        form_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        for i, label in enumerate(self.gui_labels):
            lbl = tk.Label(form_frame, text=label)
            lbl.grid(row=i, column=0, sticky='w', pady=5, padx=5)

            if label == "ENTRATA":
                entry = DateEntry(form_frame, width=47, background='darkblue', foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
                entry._top_cal.overrideredirect(False)  # Allow window to close
            elif label == "STATO":
                entry = ttk.Combobox(form_frame, values=[
                    "ATT.PERZ.", "ATT.AUT.", "ATT.RIC.", "LAV.CAR1", "LAV.CAR2", "LAV.CAR3",
                    "LAV.CAR4", "LAV.MECC.", "FIN", "ALTRI LAVORI", "DA FATTURARE", "PRONTA", "PRE-CONSEGNA"
                ], width=47)
            elif label == "RICAMBI":
                entry = ttk.Combobox(form_frame, values=["Sì", "NO"], width=47)
            else:
                entry = tk.Entry(form_frame, width=50)

            entry.grid(row=i, column=1, sticky='ew', pady=5, padx=5)
            self.entries[label] = entry

        button_frame = tk.Frame(self.root)
        button_frame.pack(fill=tk.X, padx=20, pady=20)

        self.save_button = tk.Button(button_frame, text="Save Data", command=self.save_data)
        self.save_button.pack(pady=20)

    def check_duplicate(self, data):
        if os.path.exists(self.file_path):
            df = pd.read_excel(self.file_path)
            if not df.empty:
                duplicate = df[(df["TARGA"] == data["TARGA"]) & (df["ENTRATA"] == data["ENTRATA"])]
                if not duplicate.empty:
                    return True
        return False

    def calculate_business_days(self, start_date, end_date):
        return np.busday_count(start_date.date(), end_date.date())

    def save_data(self):
        data = {label: entry.get().upper() if isinstance(entry, tk.Entry) else entry.get() for label, entry in self.entries.items()}

        if self.check_duplicate(data):
            messagebox.showwarning("Warning", "The combination of TARGA and ENTRATA already exists in the file!")
            return

        additional_data = {
            "PREV.USCITA": '',
            "FER. VET": '',
            "GG.INIZ.MECC": '',
            "INIZIO.MECC": '',
            "GG.LAV.MECC": '',
            "FINE MECC": '',
            "GG.INIZIO.CARR.": '',
            "INIZIO CARR": '',
            "GG.LAV.CAR": '',
            "FINE CARR": '',
            "DOWN TIME FATTURAZIONE": '',
            "NR. VETTURE": '',
            "DOWN TIME": '',
            "DATA ULTIMA ATTIVITA'": '',
            "FERMO TECNICO": ''
        }

        if data["ENTRATA"] == '':
            data["ENTRATA"] = None
        else:
            try:
                entrata_date = datetime.strptime(data["ENTRATA"], "%d/%m/%Y")
                additional_data["PREV.USCITA"] = (entrata_date + timedelta(days=10)).strftime("%d/%m/%Y")
            except Exception as e:
                messagebox.showerror("Error", f"Date format error: {e}")
                return

        complete_data = {**data, **additional_data}

        for label in self.excel_labels:
            if label not in complete_data:
                complete_data[label] = ''

        if os.path.exists(self.file_path):
            df = pd.read_excel(self.file_path)
        else:
            df = pd.DataFrame(columns=self.excel_labels)

        new_df = pd.DataFrame([complete_data])
        df = pd.concat([df, new_df], ignore_index=True)

        column_order = ["FLOTTA"] + [col for col in self.excel_labels if col != "FLOTTA"]
        df = df[column_order]

        df.to_excel(self.file_path, index=False)
        messagebox.showinfo("Info", "Data saved successfully!")
