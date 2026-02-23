import tkinter as tk
from tkinter import ttk, messagebox

from data_collection import get_sheet_data
from autofill import run_automation


def get_entry_data(url, start, num):
    # Basic validation
    if not url or not start or not num:
        messagebox.showwarning("Input Error", "Please fill in all fields.")
        return

    # Fetch the data from the google sheet
    sheet_data = get_sheet_data(url, start, num)
    
    if sheet_data:
        # Run the autofiller with the data passed to it
        run_automation(sheet_data)
        
    else:
        messagebox.showerror("Data Error", "Could not fetch sheet. Check URL/Permissions.")


def gui_display():
    root = tk.Tk()
    root.title("Refurbished Computers Entering Bot")
    root.geometry("400x350")
    root.resizable(False, False)

    container = ttk.Frame(root, padding="20")
    container.pack(fill="both", expand=True)

    # Sheet url: label and text box
    ttk.Label(container, text="Google Sheet URL:").pack(anchor="w", pady=(0, 5))
    sheet_var = tk.StringVar(value="https://docs.google.com/spreadsheets/d/1JN4Cq0aqU9mL1KndfEY4a7F_cSZv3NWILeqjl3UO4uw/edit?gid=292262298#gid=292262298")
    ttk.Entry(container, textvariable=sheet_var, width=50).pack(fill="x", pady=(0, 15))

    row_frame = ttk.Frame(container)
    row_frame.pack(fill="x")

    # Starting row: label and spin box
    left_col = ttk.Frame(row_frame)
    left_col.pack(side="left", expand=True)
    ttk.Label(left_col, text="Starting row:").pack(anchor="w")
    start_var = tk.IntVar(value=5)
    ttk.Spinbox(left_col, from_=1, to=10000, textvariable=start_var, width=8).pack(anchor="w")

    # Number of rows: label and spin box
    right_col = ttk.Frame(row_frame)
    right_col.pack(side="right", expand=True)
    ttk.Label(right_col, text="# of rows:").pack(anchor="w")
    num_var = tk.IntVar(value=5)
    ttk.Spinbox(right_col, from_=1, to=100, textvariable=num_var, width=8).pack(anchor="w")

    # Run button
    btn = ttk.Button(container, text='Run program',command=lambda: get_entry_data(sheet_var.get(), start_var.get(), num_var.get()))
    btn.pack(pady=30, fill="x")

    root.mainloop()


if __name__ == "__main__":
    gui_display()