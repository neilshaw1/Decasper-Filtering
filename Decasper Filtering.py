import pandas as pd
import xlsxwriter
from tkinter import filedialog, Tk, messagebox
import tkinter as tk
import threading
import os

law_office_map = {
    "ADAME GARZA": "DECASPER", "ADAME GARZA (D)": "DECASPER", "ADAME GARZA(D)": "DECASPER", "ADAMSON AHDOOT": "DECASPER",
    "AK LAW": "DECASPER", "ALADDIN LAW": "OURS", "ALADDINS": "OURS", "ALADDINS LAW": "OURS", "AMARO LAW": "OURS",
    "ANDRE THOMAS": "OURS", "ANDREA I JONES": "OURS", "ANGEL REYES": "DECASPER", "AP LAW": "OURS", "AP LAW GRP": "OURS",
    "AQRAWI": "OURS", "BERGQUIST": "DECASPER", "BEVERLY CARU(D)": "DECASPER", "BEVERLY CARUTHE": "DECASPER",
    "BRANN SULLIVAN": "DECASPER", "BRIAN WHITE": "OURS", "CALDERON": "OURS", "CAQUIAS LAW": "DECASPER",
    "CARLSON": "DECASPER", "CARLSON (D)": "DECASPER", "CARLSON LAW FIRM": "DECASPER", "CHAPA LAW": "DECASPER",
    "CHERIKA EDWARDS": "DECASPER", "CROCKETT LAW": "DECASPER", "D’ANN HINKLE": "OURS", "DAG LAW": "OURS",
    "DALY AND BLACK": "OURS", "DEHOYOS": "OURS", "EDWARDS": "OURS", "EDWARDS LAW": "OURS", "ELCO ZUBAITE": "DECASPER",
    "ESH LAW": "DECASPER", "FEL TABANGAY": "DECASPER", "FELDMAN LEE": "OURS", "FIELDING LAW": "OURS",
    "FIROZBAKHT": "DECASPER", "GALLO UWALAKA": "OURS", "GALVAN LAW": "DECASPER", "GIBSON HILL": "DECASPER",
    "GIBSON HILL (D)": "DECASPER", "GOLDENZWEIG": "OURS", "GRANADOS": "DECASPER", "HADI LAW": "OURS",
    "HAVINS": "DECASPER", "HERIBERTO RAMOS": "OURS", "HILDEBRAND": "DECASPER", "HORTON & GREGORY": "DECASPER",
    "JAMES PERKINS": "DECASPER", "JAS JORDAN": "OURS", "JD SILVA": "OURS", "JESUS DAVILA": "DECASPER",
    "JOHNSON GARCIA": "OURS", "K & P": "OURS", "KANNER PINTALUGA": "OURS", "KENNY PEREZ": "DECASPER",
    "KGS": "DECASPER", "KGS LAW": "DECASPER", "KHERKHER GARCIA": "DECASPER", "KLITSAS VERCHER": "DECASPER",
    "KV LAW": "DECASPER", "LANDER": "OURS", "LAW BOSS": "DECASPER", "LE BLANC LAW FIRM": "DECASPER",
    "LEO & OGINNI": "DECASPER", "LIDJI LAW": "DECASPER", "M&Y PERSONAL INJURY": "DECASPER", "MICHAEL WATSON": "DECASPER",
    "MOISES MORALES": "OURS", "MOKARAM": "OURS", "MORGAN BORQUE": "DECASPER", "MOUDGIL": "DECASPER",
    "MUKERJI": "DECASPER", "MUNOZ": "OURS", "MUNOZ & ASSOCIATES": "OURS", "MUNOZ ASSOC": "OURS",
    "NGUYEN AND DELCID": "OURS", "NMW": "OURS", "NMW LAW FIRM": "OURS", "NOYOLA LAW FIRM": "DECASPER",
    "ORIHUELA": "DECASPER", "OWSLEY": "DECASPER", "PARDO HOMAN": "DECASPER", "PARDO HOMAN(D)": "DECASPER",
    "PAYNE": "OURS", "PHIPPS GARZA": "DECASPER", "PM LAW": "OURS", "PMR": "OURS", "REYES BROWN": "DECASPER",
    "REYES BROWNE": "DECASPER", "ROBERTS MARKLAN": "OURS", "RODNEY JONES": "OURS", "RUTH RIVERA": "OURS",
    "RYAN SNIDER": "DECASPER", "SALAZAR LAW": "DECASPER", "SCOTT LANNIE": "DECASPER", "SDB LAW GRP": "DECASPER",
    "SERVOS": "DECASPER", "SERVOS LAW": "DECASPER", "SHARIFF LAW": "OURS", "SIMMONS FLETCHE": "DECASPER",
    "SIMON AND OROUK": "DECASPER", "SNEED MITCHELL": "DECASPER", "SOLIZ": "OURS", "SS&H": "DECASPER", "SSH": "DECASPER",
    "STEPHEN BOUTROS": "DECASPER", "STEPHENS JUREN": "DECASPER", "STEWART GUSS": "DECASPER", "STRICKLAND LAW": "DECASPER",
    "SVR LAW": "DECASPER", "TAKLA": "DECASPER", "TAKLA LAW": "DECASPER", "TALABI": "DECASPER", "TAYLOR LAW": "OURS",
    "THURLOW & ASSOC": "OURS", "UNIVERSAL": "OURS", "URIBE LAW FIRM": "DECASPER", "VILLAREAL LEGAL": "DECASPER",
    "WHITE LAW": "OURS", "Z & P LAW": "DECASPER", "Z AND P LAW": "DECASPER", "Z&P LAW": "DECASPER",
    "ZANE WEEKS": "DECASPER"
}

def clean_value(val):
    if pd.isna(val):
        return val
    try:
        return pd.to_numeric(val)
    except:
        try:
            return pd.to_datetime(val).date()
        except:
            return str(val).strip().upper().replace('="', '').replace('"', '')

def process_file(file_path):
    df = pd.read_csv(file_path, dtype=str)
    df.columns = df.columns.str.strip().str.upper()

    df = df.applymap(clean_value)

    if "PRESEMAIL" not in df.columns or "GROUPNO" not in df.columns:
        messagebox.showerror("Error", "Input CSV missing PRESEMAIL or GROUPNO columns.")
        return

    df['MAPVAL'] = df['GROUPNO'].map(law_office_map)
    df_filtered = df[(df['PRESEMAIL'] == "DECASPER@GMAIL") & (df['MAPVAL'] == "DECASPER")].copy()
    df_filtered.drop(columns=['MAPVAL'], inplace=True)

    if not df_filtered.empty:
        for col in ["DATEF", "PATDOB"]:
            if col in df_filtered.columns:
                parsed = pd.to_datetime(df_filtered[col], errors='coerce')
                df_filtered[col] = parsed.dt.date.astype(str).where(parsed.notna(), df_filtered[col].astype(str))

        cols_to_drop = ["PRESEMAIL", "PICKEDUP"]
        df_filtered = df_filtered.drop(columns=[c for c in cols_to_drop if c in df_filtered.columns])

        folder = os.path.dirname(file_path)
        save_path = os.path.join(folder, "filtered_decasper.xlsx")

        with pd.ExcelWriter(save_path, engine='xlsxwriter', datetime_format='yyyy-mm-dd') as writer:
            df_filtered.to_excel(writer, index=False, sheet_name="Filtered Data")
            workbook = writer.book
            worksheet = writer.sheets["Filtered Data"]

            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': False,
                'valign': 'center',
                'fg_color': '#DAF2D0',
                'font_color': '#000000',
                'border': 0
            })
            for col_num, value in enumerate(df_filtered.columns.values):
                worksheet.write(0, col_num, value, header_format)

            for i, col in enumerate(df_filtered.columns):
                column_len = max(df_filtered[col].astype(str).map(len).max(), len(col)) + 4
                worksheet.set_column(i, i, column_len)

            if "TOTALRXAMOUNT" in df_filtered.columns:
                col_idx = df_filtered.columns.get_loc("TOTALRXAMOUNT")
                total_series = pd.to_numeric(df_filtered["TOTALRXAMOUNT"], errors='coerce')

                for row_num, value in enumerate(total_series, start=1):
                    if pd.notna(value):
                        worksheet.write(row_num, col_idx, f"${value:,.2f}")
                    else:
                        worksheet.write(row_num, col_idx, "")

                total_row = len(df_filtered) + 1
                total_sum = total_series.sum()
                total_format = workbook.add_format({
                    'bold': True,
                    'fg_color': '#DAF2D0',
                    'font_color': '#000000',
                    'border': 0
                })
                for i_col in range(len(df_filtered.columns)):
                    if i_col == col_idx:
                        worksheet.write(total_row, i_col, f"${total_sum:,.2f}", total_format)
                    else:
                        worksheet.write(total_row, i_col, "", total_format)
            else:
                total_row = len(df_filtered) + 1
                total_format = workbook.add_format({
                    'bold': True,
                    'fg_color': '#DAF2D0',
                    'font_color': '#000000',
                    'border': 0
                })

            if "PATLNAME" in df_filtered.columns:
                lname_series = df_filtered["PATLNAME"].astype(str).str.strip()
                lname_series = lname_series.replace("", pd.NA)
                unique_people_count = lname_series.dropna().nunique()
            else:
                unique_people_count = len(df_filtered)

            unique_row = total_row + 2
            worksheet.write(unique_row, 0, f"# of Unique Names: {unique_people_count}", total_format)

        messagebox.showinfo("Success", f"Filtered file saved at:\n{save_path}")
    else:
        messagebox.showinfo("No Matches", "No matching rows found.")

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Decasper CSV Filter")
        self.geometry("450x220")
        self.resizable(False, False)

        self.label = tk.Label(self, text="Select a CSV file to filter Decasper rows", font=("Segoe UI", 11), wraplength=400)
        self.label.pack(pady=20)

        self.select_button = tk.Button(self, text="Select CSV File", width=20, font=("Segoe UI", 10), command=self.select_file)
        self.select_button.pack(pady=10)

        self.status_label = tk.Label(self, text="", fg="green", font=("Segoe UI", 10))
        self.status_label.pack(pady=10)

        self.quit_button = tk.Button(self, text="Quit", width=15, command=self.destroy)
        self.quit_button.pack(pady=10)

    def select_file(self):
        file_path = filedialog.askopenfilename(title="Select CSV File", filetypes=[("CSV files", "*.csv")])
        if file_path:
            self.status_label.config(text="Processing file...")
            threading.Thread(target=self.run_processing, args=(file_path,)).start()

    def run_processing(self, file_path):
        try:
            process_file(file_path)
            self.status_label.config(text="File processed successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred:\n{e}")
            self.status_label.config(text="Processing failed.")

app = App()
app.mainloop()

