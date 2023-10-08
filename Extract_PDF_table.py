import tabula
import pyperclip
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, simpledialog, messagebox

def extract_and_save():
    pdf_path = pdf_file_entry.get()
    excel_path = excel_file_entry.get()
    start_page = int(start_page_entry.get())
    end_page = int(end_page_entry.get())
    merge_adjacent = merge_adjacent_var.get()

    try:
        dfs = []
        for page_num in range(start_page, end_page+1):
            df = tabula.read_pdf(pdf_path, pages=page_num, multiple_tables=False)
            dfs.extend(df)
        
        final_df = pd.concat(dfs, ignore_index=True)

        if merge_adjacent:
            for col in final_df.columns:
                final_df[col] = final_df[col].fillna(method='ffill')

        final_df.to_excel(excel_path, index=False)
        messagebox.showinfo("Success", "Table extraction successful!")
    except Exception as e:
        exception_text = str(e)
        messagebox.showerror("Error", exception_text)
        pyperclip.copy(exception_text)

root = tk.Tk()
root.title("PDF to Excel Converter")

frame = ttk.Frame(root, padding="10")
frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

pdf_file_entry = ttk.Entry(frame, width=50)
pdf_file_entry.grid(row=0, column=1)
ttk.Button(frame, text="Browse PDF", command=lambda: pdf_file_entry.insert(0, filedialog.askopenfilename())).grid(row=0, column=2)

excel_file_entry = ttk.Entry(frame, width=50)
excel_file_entry.grid(row=1, column=1)
ttk.Button(frame, text="Save Excel As", command=lambda: excel_file_entry.insert(0, filedialog.asksaveasfilename(defaultextension=".xlsx"))).grid(row=1, column=2)

start_page_entry = ttk.Entry(frame, width=5)
start_page_entry.grid(row=2, column=1, sticky=tk.W)
ttk.Label(frame, text="Start Page").grid(row=2, column=0, sticky=tk.E)

end_page_entry = ttk.Entry(frame, width=5)
end_page_entry.grid(row=3, column=1, sticky=tk.W)
ttk.Label(frame, text="End Page").grid(row=3, column=0, sticky=tk.E)

merge_adjacent_var = tk.BooleanVar()
merge_adjacent_check = ttk.Checkbutton(frame, text="Merge adjacent cells", variable=merge_adjacent_var)
merge_adjacent_check.grid(row=4, columnspan=3)

ttk.Button(frame, text="Convert", command=extract_and_save).grid(row=5, columnspan=3)

root.mainloop()
