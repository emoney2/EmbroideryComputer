import tkinter as tk
from tkinter import simpledialog
import gspread
from google.oauth2.service_account import Credentials
import pygetwindow as gw
import keyboard
import re

def prompt_cut_type():
    """Prompt the user to select a cut type using buttons."""
    def set_cut_type(selection):
        nonlocal cut_type
        cut_type = selection
        root.destroy()

    root = tk.Tk()
    root.title("Cut Type Selection")
    root.geometry("300x150")
    root.attributes('-topmost', True)

    cut_type = None

    label = tk.Label(root, text="Does the job need to be custom cut?", pady=10)
    label.pack()

    button_die_cut = tk.Button(root, text="Die Cut", command=lambda: set_cut_type("Die Cut"))
    button_die_cut.pack(pady=5)

    button_custom_cut = tk.Button(root, text="Custom Cut", command=lambda: set_cut_type("Custom Cut"))
    button_custom_cut.pack(pady=5)

    button_both = tk.Button(root, text="Both", command=lambda: set_cut_type("Both"))
    button_both.pack(pady=5)

    root.mainloop()
    return cut_type

def write_cut_type_to_sheet(order_number, cut_type):
    """Write the selected cut type to the Google Sheet."""
    try:
        # Load credentials and connect to Google Sheets
        creds_path = r"C:\Users\eckar\Desktop\OrderEntry,Inventory,QR,PrintPDF\Keys\poetic-logic-454717-h2-3dd1bedb673d.json"
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        creds = Credentials.from_service_account_file(creds_path, scopes=scopes)
        client = gspread.authorize(creds)

        # Open the spreadsheet and the Production Orders tab
        sheet_name = "JR and Co."
        tab_name = "Production Orders"
        spreadsheet = client.open(sheet_name)
        worksheet = spreadsheet.worksheet(tab_name)

        # Get all values in the first row (headers)
        headers = worksheet.row_values(1)

        # Debugging: Log all headers to verify column names
        print(f"Headers in the sheet: {headers}")

        # Find the "Cut Type" column
        cut_type_col = next((i + 1 for i, h in enumerate(headers) if h.strip().lower() == "cut type"), None)
        if not cut_type_col:
            print("Cut Type column not found in the sheet.")
            return

        # Find the row for the given order number
        order_col = next((i + 1 for i, h in enumerate(headers) if h.strip().lower() in ["order number", "order #"]), None)
        if not order_col:
            print("Order # column not found in the sheet. Check the column headers.")
            return

        order_numbers = worksheet.col_values(order_col)
        order_row = next((i + 1 for i, num in enumerate(order_numbers) if str(num).strip() == str(order_number)), None)
        if not order_row:
            print(f"Order number {order_number} not found in the sheet.")
            return

        # Write the cut type to the sheet
        worksheet.update_cell(order_row, cut_type_col, cut_type)
        print(f"Cut type '{cut_type}' written to order {order_number}.")
    except Exception as e:
        print(f"Error writing to Google Sheet: {e}")

def get_active_job_number():
    """Extract the job number from the active window title."""
    active_window = gw.getActiveWindow()
    if active_window:
        title = active_window.title
        print(f"Active window title: {title}")  # Debugging log

        # Extract the numeric part of the job number, ensuring it captures digits within brackets
        match = re.search(r"\[(\d+)\]", title)
        if not match:
            match = re.search(r"\[.*?(\d+).*?\]", title)  # Fallback for any digits within brackets
        if match:
            job_number = match.group(1)
            print(f"Extracted job number: {job_number}")  # Debugging log
            return job_number.strip()

        print("No job number found in the window title.")  # Debugging log
        return None
    print("No active window found.")  # Debugging log
    return None

def on_control_alt_t():
    """Handle the Control+Alt+T key combination."""
    job_number = get_active_job_number()
    if not job_number:
        print("No active job found.")
        return

    cut_type = prompt_cut_type()
    if cut_type:
        write_cut_type_to_sheet(job_number, cut_type)

def test_control_alt_t():
    """Test function for Control+Alt+T."""
    order_number = "12345"  # Replace with a valid order number for testing
    cut_type = prompt_cut_type()
    if cut_type:
        write_cut_type_to_sheet(order_number, cut_type)

# Listen for the Control+Alt+T key combination
keyboard.add_hotkey("ctrl+alt+t", on_control_alt_t)

print("Listening for Control+Alt+T... Press ESC to exit.")
keyboard.wait("esc")
