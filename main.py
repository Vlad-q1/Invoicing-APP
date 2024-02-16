import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox, filedialog
import generate_invoice as inv_gen
import pandas as pd
import logging
import threading
import os

excel_path = None

required_fields = ['invoice_number', 'Month', 'DCU', 'print_value_ron', 'print_value_eur', 'total_in_ron_de_printat', 'print_value_eur_total', 'print_exchange_rate']

logging.basicConfig(filename='app.log', filemode='w', format='%(name)s - %(levelname)s - %(message)s', level=logging.INFO)

def load_excel_file():
    logging.info('load_excel_file started')
    global excel_path
    excel_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if excel_path:
        try:
            data = pd.read_excel(excel_path)
            data.columns = data.columns.str.replace(' ', '_')
            missing_fields = [field for field in required_fields if field not in data.columns]
            if missing_fields:
                messagebox.showerror("Error", f"The following fields are missing in the Excel file: {', '.join(missing_fields)}")
                excel_path = None
            else:
                generate_button.pack(pady=10)
        except Exception as err:
            messagebox.showerror("Error", f"Error loading Excel file: {str(err)}")
            excel_path = None
    logging.info('load_excel_file ended')

def generate_invoice():
    logging.info('generate_invoice started')
    if excel_path is not None:
        inv_gen.generate_invoice(excel_path, close_window)
        if messagebox.showinfo("Task Finished", "The Invoices have been generated successfully") == "ok":
            os._exit(0)
    logging.info('generate_invoice closed')

def close_window():
    logging.info('close_window function started')
    window.destroy()
    logging.info('close_window function ended')

if __name__ == "__main__":
    logging.info('main function started')
    window = tk.Tk()
    window.title("PDF Generator")
    window.geometry("300x200")

    tk.Label(window, text=" ").pack()

    load_button = tk.Button(window, text="Load Excel file", command=lambda: threading.Thread(target=load_excel_file).start(), width=20, height=2, bg="blue", fg="white")
    load_button.pack(pady=10)

    generate_button = tk.Button(window, text="Generate PDF", command=lambda: threading.Thread(target=generate_invoice).start(), width=20, height=2, bg="green", fg="white")

    exit_button = tk.Button(window, text="Exit", command=close_window, width=20, height=2, bg="red", fg="white")
    exit_button.pack(pady=10)

    window.protocol("WM_DELETE_WINDOW", close_window)

    window.mainloop()
    logging.info('main function ended')