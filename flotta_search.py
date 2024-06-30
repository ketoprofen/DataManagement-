# flotta_search.py
import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd
import os
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from datetime import datetime

class FlottaSearchApp:
    def __init__(self, root, file_path):
        self.root = root
        self.root.title("Search FLOTTA")
        self.root.geometry("800x600")
        self.file_path = file_path

        self.labels = [
            "FLOTTA", "TARGA", "MODELLO", "ENTRATA", "PREV.USCITA", "FER. VET", "DITTA",
            "GG.INIZ.MECC", "INIZIO.MECC", "GG.LAV.MECC", "FINE MECC",
            "GG.INIZIO.CARR.", "INIZIO CARR", "GG.LAV.CAR", "FINE CARR",
            "PZ CARR", "STATO", "DOWN TIME", "DATA ULTIMA ATTIVITA'", "RICAMBI", "FERMO TECNICO"
        ]

        search_frame = tk.Frame(self.root)
        search_frame.pack(pady=20)

        tk.Label(search_frame, text="Enter FLOTTA:").grid(row=0, column=0, padx=10, pady=10)
        self.flotta_entry = tk.Entry(search_frame)
        self.flotta_entry.grid(row=0, column=1, padx=10, pady=10)
        search_button = tk.Button(search_frame, text="Search", command=self.display_search_results)
        search_button.grid(row=0, column=2, padx=10, pady=10)

        self.tree = ttk.Treeview(self.root, columns=self.labels, show='headings')
        for col in self.labels:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)
        
        self.tree.pack(fill=tk.BOTH, expand=True)

        generate_pdf_button = tk.Button(self.root, text="Generate PDF", command=self.generate_pdf)
        generate_pdf_button.pack(pady=10)
        
        generate_excel_button = tk.Button(self.root, text="Generate Excel", command=self.generate_excel)
        generate_excel_button.pack(pady=10)

    def display_search_results(self):
        self.flotta = self.flotta_entry.get()
        if os.path.exists(self.file_path):
            df = pd.read_excel(self.file_path)
            results = df[df["FLOTTA"].str.contains(self.flotta, case=False, na=False)]
            if not results.empty:
                self.tree.delete(*self.tree.get_children())
                for _, row in results.iterrows():
                    self.tree.insert("", tk.END, values=list(row))

                # Prepare data for PDF generation
                self.stato_counts = results["STATO"].value_counts()
                self.total_targa = len(results["TARGA"].unique())
                self.results = results

                # Generate table data for PDF and Excel
                self.generate_table_data()

            else:
                messagebox.showinfo("Info", "No results found!")
        else:
            messagebox.showinfo("Info", "No data file found!")

    def generate_table_data(self):
        unique_stato = [
            "ATT.PERZ.", "ATT.AUT.", "ATT.RIC.", "LAV.CAR1", "LAV.CAR2", "LAV.CAR3",
            "LAV.CAR4", "LAV.MECC.", "FIN", "ALTRI LAVORI", "DA FATTURARE", "PRONTA", "PRE-CONSEGNA"
        ]
        table_data = [["Row"] + unique_stato]
        max_rows = self.results["STATO"].value_counts().max()

        for i in range(max_rows):
            row = [f"Row {i + 1}"]
            for stato in unique_stato:
                targa_values = self.results[self.results["STATO"] == stato]["TARGA"].values
                row.append(targa_values[i] if i < len(targa_values) else "")
            table_data.append(row)

        # Add the counts at the bottom
        counts_row = ["Count"]
        for stato in unique_stato:
            counts_row.append(self.results[self.results["STATO"] == stato]["TARGA"].count())
        table_data.append(counts_row)

        total_count_row = ["Total"] + [""] * (len(unique_stato) - 1) + [self.total_targa]
        table_data.append(total_count_row)

        self.table_data = table_data

    def generate_pdf(self):
        pdf_path = f"{self.flotta}.pdf"
        doc = SimpleDocTemplate(pdf_path, pagesize=A4)
        elements = []

        styles = getSampleStyleSheet()
        title_style = styles['Title']
        normal_style = styles['Normal']

        current_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        elements.append(Paragraph(f"Search Results for FLOTTA: {self.flotta} (Generated on {current_date})", title_style))
        elements.append(Paragraph(f"Total TARGA: {self.total_targa}", normal_style))

        table = Table(self.table_data)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ]))
        elements.append(table)

        doc.build(elements)
        messagebox.showinfo("Info", f"PDF generated successfully: {pdf_path}")

    def generate_excel(self):
        excel_path = f"{self.flotta}.xlsx"
        
        # Create a DataFrame from the table data
        df_table = pd.DataFrame(self.table_data[1:], columns=self.table_data[0])
        
        # Add a title row
        current_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        df_title = pd.DataFrame([[f"Search Results for FLOTTA: {self.flotta} (Generated on {current_date})"] + [""] * (len(df_table.columns) - 1)], columns=df_table.columns)
        
        # Concatenate title row with the data
        df_final = pd.concat([df_title, df_table], ignore_index=True)
        
        # Save the DataFrame to an Excel file
        with pd.ExcelWriter(excel_path) as writer:
            df_final.to_excel(writer, sheet_name=self.flotta, index=False)
        
        messagebox.showinfo("Info", f"Excel file generated successfully: {excel_path}")

if __name__ == "__main__":
    root = tk.Tk()
    app = FlottaSearchApp(root, "data.xlsx")
    root.mainloop()
