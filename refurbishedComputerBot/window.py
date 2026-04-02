import tkinter as tk
from tkinter import ttk, messagebox
import threading

from data_collection import get_sheet_data
from autofill import run_automation

from config import DEFAULT_SHEET_URL


class AnimatedStatus:
    """Cycles dots on a ttk.Label to create a loading animation."""

    def __init__(self, root, label, interval=400):
        self.root = root
        self.label = label
        self.interval = interval
        self.base_text = ""
        self.dot_count = 0
        self._job = None

    def start(self, text):
        """Start animating with the given base text."""
        self.stop()
        self.base_text = text
        self.dot_count = 0
        self._animate()

    def stop(self, final_text=""):
        """Stop animating and optionally set a final static message."""
        if self._job is not None:
            self.root.after_cancel(self._job)
            self._job = None
        self.label.config(text=final_text)

    def _animate(self):
        dots = "." * (self.dot_count % 4)
        self.label.config(text=f"{self.base_text}{dots}")
        self.dot_count += 1
        self._job = self.root.after(self.interval, self._animate)


def get_entry_data(url, start, num, root, order_entry, error_label_ref, btn, status):
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
    status.start("Fetching spreadsheet data")

    def _run():
        # Fetch the data from the google sheet
        sheet_data, valid_data, error_message = get_sheet_data(url, start, num)

        if valid_data:
            root.after(0, lambda: status.start("Opening browser"))
            run_automation(sheet_data, root, order_entry, status)
        else:
            # Add error message label to window
            def _show_errors():
                error_label_ref[0] = ttk.Label(root, text=error_message, foreground='red')
                error_label_ref[0].pack()
                btn.config(state="normal")
                status.stop()
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
    sheet_var = tk.StringVar(value=DEFAULT_SHEET_URL)
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

    # Status label with animated dots
    status_label = ttk.Label(container, text="", foreground='gray')
    status_label.pack(pady=(0, 5))
    status = AnimatedStatus(root, status_label)

    # Run button
    btn = ttk.Button(container, text='Run program',
        command=lambda: get_entry_data(
            sheet_var.get(), start_var.get(), num_var.get(),
            root, order_entry_var.get(), error_label_ref, btn, status
        ))
    btn.pack(pady=5, fill="x")

    root.mainloop()


if __name__ == "__main__":
    gui_display()