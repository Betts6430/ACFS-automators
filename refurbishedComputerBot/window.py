import tkinter as tk
from tkinter import ttk, messagebox
import threading

from data_collection import get_sheet_data
from autofill import run_automation


def get_entry_data(url, start, num, root, order_entry, error_label_ref, btn, status_label):
    # Clear previous error message if it exists
    if error_label_ref[0] is not None:
        error_label_ref[0].destroy()
        error_label_ref[0] = None

    # Basic validation
    if not url or not start or not num:
        messagebox.showwarning("Input Error", "Please fill in all fields.")
        return

    # Disable button and show status while fetching
    btn.config(state="disabled")
    status_label.config(text="Fetching spreadsheet data...")

    def _run():
        # Fetch the data from the google sheet
        sheet_data, valid_data, error_message = get_sheet_data(url, start, num)

        if valid_data:
            root.after(0, lambda: status_label.config(text="Opening browser..."))
            run_automation(sheet_data, root, order_entry, status_label)
        else:
            # Add error message label to window
            def _show_errors():
                error_label_ref[0] = ttk.Label(root, text=error_message, foreground='red')
                error_label_ref[0].pack()
                btn.config(state="normal")
                status_label.config(text="")
            root.after(0, _show_errors)

    threading.Thread(target=_run, daemon=True).start()


def gui_display():
    root = tk.Tk()
    root.title("Refurbished Computers Entering Bot")
    root.resizable(False, False)

    container = ttk.Frame(root, padding="20")
    container.pack(fill="both", expand=True)

    # Sheet url: label and text box
    ttk.Label(container, text="Google Sheet URL:").pack(anchor="w", pady=(0, 5))
    sheet_var = tk.StringVar(value="https://docs.google.com/spreadsheets/d/1mARf98z1tTqTimLuweBU10VGqXJilK9cpAYh2CchmDo/edit?gid=209298323#gid=209298323")
    ttk.Entry(container, textvariable=sheet_var, width=50).pack(fill="x", pady=(0, 15))

    row_frame = ttk.Frame(container)
    row_frame.pack(fill="x")

    # Starting row: label and spin box
    left_col = ttk.Frame(row_frame)
    left_col.pack(side="left", expand=True)
    ttk.Label(left_col, text="Starting row:").pack(anchor="w")
    start_var = tk.IntVar(value=1)
    ttk.Spinbox(left_col, from_=1, to=10000, textvariable=start_var, width=8).pack(anchor="w")

    # Number of rows: label and spin box
    right_col = ttk.Frame(row_frame)
    right_col.pack(side="right", expand=True)
    ttk.Label(right_col, text="# of rows:").pack(anchor="w")
    num_var = tk.IntVar(value=10)
    ttk.Spinbox(right_col, from_=1, to=100, textvariable=num_var, width=8).pack(anchor="w")

    # Order entry check box
    order_entry_var = tk.IntVar()
    order_entry_var.set(1)
    check_btn = ttk.Checkbutton(container, text="Auto-enter into order?", variable=order_entry_var)
    check_btn.pack(pady=25)

    error_label_ref = [None]

    # Status label
    status_label = ttk.Label(container, text="", foreground='gray')
    status_label.pack(pady=(0, 5))

    # Run button
    btn = ttk.Button(container, text='Run program',
        command=lambda: get_entry_data(
            sheet_var.get(), start_var.get(), num_var.get(),
            root, order_entry_var.get(), error_label_ref, btn, status_label
        ))
    btn.pack(pady=5, fill="x")

    root.mainloop()


if __name__ == "__main__":
    gui_display()