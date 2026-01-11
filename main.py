import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import sys

def get_data():
    root = tk.Tk()
    root.withdraw()
    # Lift the dialog to the front on Mac
    root.attributes('-topmost', True) 
    
    file_path_str = filedialog.askopenfilename(
        title="Select Attendance File (CSV or Excel)",
        filetypes=[("Data files", "*.xlsx *.csv *.xls"), ("All files", "*.*")]
    )
    
    root.destroy() 
    if not file_path_str:
        sys.exit()
        
    p = Path(file_path_str)
    # Check extension and return DF
    if p.suffix.lower() == '.csv':
        return pd.read_csv(p)
    else:
        return pd.read_excel(p)

def main():
    # 1. Get the data from the user
    df = get_data()

    # 2. Cleanup and Process
    df.columns = [s.lower().replace(" ", "_") for s in df.columns]

    daily_df = (
        df
        .assign(
            full_timestamp=lambda d: pd.to_datetime(
                d["punch_date"].astype(str) + " " + d["attendance_record"].astype(str)
            )
        )
        .groupby(["person_id", "person_name", "punch_date"])
        .agg(
            punch_count=("full_timestamp", "size"),
            first_punch=("full_timestamp", lambda s: s.min().time()),
            last_punch=("full_timestamp", lambda s: s.max().time()),
            in_to_out_time=("full_timestamp", lambda s: (pd.to_datetime("1970-01-01") + (s.max() - s.min())).time()),
        )
        .reset_index()
        .sort_values(["person_name", "person_id", "punch_date"])
    )

    # 3. Save to Excel
    output_file_path = Path.home() / "Desktop" / f'report_{pd.Timestamp.today().date()}.xlsx'

    with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
        daily_df.to_excel(writer, sheet_name="daily", index=False)
        workbook = writer.book
        text_format = workbook.add_format({'num_format': '@'}) 
        money_format = workbook.add_format({'num_format': '#,##0.00'}) 

        worksheet = writer.sheets["daily"]
        for i, col in enumerate(daily_df.columns):
            width = 20
            col_str = str(col).lower()
            if any(x in col_str for x in ['id', 'no', 'code', 'iban', 'account']):
                worksheet.set_column(i, i, width, text_format)
            elif any(x in col_str for x in ['amount', 'charge', 'price', 'rate', 'cost', 'total']):
                worksheet.set_column(i, i, width, money_format)
            else:
                worksheet.set_column(i, i, width)

    # 4. Final Success Popup
    final_root = tk.Tk()
    final_root.withdraw()
    final_root.attributes('-topmost', True)
    messagebox.showinfo("Success", f"Report saved to Desktop:\n{output_file_path.name}")
    final_root.destroy()

if __name__ == "__main__":
    main()