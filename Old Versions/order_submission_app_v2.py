
import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog
from tkinter import ttk
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
import re
import datetime
import threading

# For Google Drive operations:
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# For formatting cells in Sheets:
from gspread_formatting import (CellFormat, format_cell_range,
                                DataValidationRule, BooleanCondition,
                                set_data_validation_for_cell_range)

# For image previews:
from PIL import Image, ImageTk, ImageOps

# For PDF rendering (PyMuPDF)
import fitz

# ===================== Global Variables & Helper Functions =====================

def update_image_preview(filepath, target_widget, width, height):
    """
    Opens the file at 'filepath.'
      - If it's a PDF, uses PyMuPDF (fitz) to render the first page.
      - Otherwise, uses Pillow (with exif_transpose to account for orientation).
    Then uses ImageOps.contain so that the entire image fits within (width, height) without cropping.
    The resized image is then centered on a white background.
    Returns an ImageTk.PhotoImage.
    """
    try:
        if filepath.lower().endswith('.pdf'):
            doc = fitz.open(filepath)
            page = doc.load_page(0)
            matrix = fitz.Matrix(2, 2)  # Adjust scale as needed.
            pix = page.get_pixmap(matrix=matrix)
            image = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        else:
            image = Image.open(filepath)
            image = ImageOps.exif_transpose(image)
        resized = ImageOps.contain(image, (width, height))
        background = Image.new("RGB", (width, height), "white")
        paste_x = (width - resized.width) // 2
        paste_y = (height - resized.height) // 2
        background.paste(resized, (paste_x, paste_y))
        return ImageTk.PhotoImage(background)
    except Exception as e:
        print("Error in update_image_preview:", e)
        target_widget.delete("all")
        target_widget.create_text(width // 2, height // 2, text="Preview not available", fill="black")
        return None

SPREADSHEET = None
MATERIAL_INVENTORY_WS = None  # Cache for "Material Inventory" worksheet

MATERIALS_CACHE = None
FUR_COLORS_CACHE = None
COMPANIES_CACHE = None
PRODUCTS_CACHE = None
THREADS_CACHE = None

ORDERED_INVENTORY = []
ORDERED_TAB = None  # Will be set later

def find_exact_header_index(header_list, target):
    for i, header in enumerate(header_list):
        if header == target:
            return i
    return None

def get_next_business_day(date_obj):
    next_day = date_obj + datetime.timedelta(days=1)
    while next_day.weekday() >= 5:
        next_day += datetime.timedelta(days=1)
    return next_day

def column_letter(n):
    result = ""
    while n:
        n, rem = divmod(n-1, 26)
        result = chr(65 + rem) + result
    return result

# ----- Cache Clearing Functions -----
def clear_materials_cache():
    global MATERIALS_CACHE
    MATERIALS_CACHE = None

def clear_fur_colors_cache():
    global FUR_COLORS_CACHE
    FUR_COLORS_CACHE = None

def clear_companies_cache():
    global COMPANIES_CACHE
    COMPANIES_CACHE = None

def clear_products_cache():
    global PRODUCTS_CACHE
    PRODUCTS_CACHE = None

def clear_threads_cache():
    global THREADS_CACHE
    THREADS_CACHE = None

# ===================== FilterableCombobox =====================

class FilterableCombobox(ttk.Combobox):
    def __init__(self, master=None, **kw):
        kw.setdefault("state", "normal")
        super().__init__(master, **kw)
        self._completion_list = []
        self._after_id = None
        self.bind("<KeyRelease>", self._on_keyrelease)
    
    def set_completion_list(self, completion_list):
        self._completion_list = sorted(completion_list, key=str.lower)
        self["values"] = self._completion_list

    def _on_keyrelease(self, event):
        if event.keysym in ("BackSpace", "Left", "Right", "Delete", "Shift_L", "Shift_R", "Control_L", "Control_R"):
            return
        if self._after_id:
            self.after_cancel(self._after_id)
        self._after_id = self.after(300, self._autocomplete)

    def _autocomplete(self):
        typed = self.get()
        if not typed:
            return
        candidates = [item for item in self._completion_list if item.lower().startswith(typed.lower())]
        if candidates:
            completion = candidates[0]
            if completion.lower() != typed.lower():
                self.delete(0, tk.END)
                self.insert(0, completion)
                self.select_range(len(typed), tk.END)

# ===================== Data Loading Functions =====================

def load_list_from_sheet(sheet, header_name):
    try:
        records = sheet.get_all_records()
        values = [record[header_name] for record in records if record.get(header_name)]
        return sorted(list(set(values)))
    except Exception as e:
        messagebox.showerror("Error", f"Could not load data for {header_name}: {e}")
        return []

def get_companies():
    global COMPANIES_CACHE
    if COMPANIES_CACHE is not None:
        return COMPANIES_CACHE
    ws = open_sheet().worksheet("Directory")
    COMPANIES_CACHE = load_list_from_sheet(ws, "Company Name")
    return COMPANIES_CACHE

def get_products():
    global PRODUCTS_CACHE
    if PRODUCTS_CACHE is not None:
        return PRODUCTS_CACHE
    ws = open_sheet().worksheet("Table")
    PRODUCTS_CACHE = load_list_from_sheet(ws, "Products")
    return PRODUCTS_CACHE

def get_materials():
    global MATERIALS_CACHE
    if MATERIALS_CACHE is not None:
        return MATERIALS_CACHE
    ws = get_material_inventory_ws()
    headers = ws.row_values(1)
    if "Materials" in headers:
        col_index = headers.index("Materials") + 1
        MATERIALS_CACHE = [m for m in ws.col_values(col_index)[1:] if m]
        return MATERIALS_CACHE
    else:
        messagebox.showerror("Error", "Header 'Materials' not found in Material Inventory")
        return []

def get_fur_colors():
    global FUR_COLORS_CACHE
    if FUR_COLORS_CACHE is not None:
        return FUR_COLORS_CACHE
    ws = get_material_inventory_ws()
    headers = ws.row_values(1)
    if "Fur Color" in headers:
        col_index = headers.index("Fur Color") + 1
        FUR_COLORS_CACHE = [f for f in ws.col_values(col_index)[1:] if f]
        return FUR_COLORS_CACHE
    else:
        messagebox.showerror("Error", "Header 'Fur Color' not found in Material Inventory")
        return []

def get_threads_inventory():
    global THREADS_CACHE
    if THREADS_CACHE is not None:
        return THREADS_CACHE
    ws = get_material_inventory_ws()
    headers = ws.row_values(1)
    if "Thread Colors" in headers:
        col_index = headers.index("Thread Colors") + 1
        THREADS_CACHE = [t for t in ws.col_values(col_index)[1:] if t]
        return THREADS_CACHE
    else:
        messagebox.showerror("Error", "Header 'Thread Colors' not found in Material Inventory")
        return []

# ===================== Google Drive Functions =====================

def get_drive_service():
    drive_scope = ['https://spreadsheets.google.com/feeds',
                   'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', drive_scope)
    service = build('drive', 'v3', credentials=creds)
    return service

def make_file_public(file_id):
    service = get_drive_service()
    permission = {'type': 'anyone', 'role': 'reader'}
    service.permissions().create(fileId=file_id, body=permission).execute()

def upload_file_to_drive(filepath, folder_id):
    service = get_drive_service()
    file_metadata = {'name': os.path.basename(filepath), 'parents': [folder_id]}
    media = MediaFileUpload(filepath, resumable=True)
    file = service.files().create(body=file_metadata, media_body=media, fields='id, webViewLink').execute()
    make_file_public(file['id'])
    return file['id'], file['webViewLink']

def create_drive_folder(folder_name, parent_id=None):
    service = get_drive_service()
    file_metadata = {'name': folder_name, 'mimeType': 'application/vnd.google-apps.folder'}
    if parent_id:
        file_metadata['parents'] = [parent_id]
    folder = service.files().create(body=file_metadata, fields='id, webViewLink').execute()
    make_file_public(folder['id'])
    return folder['id'], folder['webViewLink']

# ===================== Google Sheets Setup =====================

def init_google_client():
    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    client = gspread.authorize(creds)
    return client

def open_sheet():
    global SPREADSHEET
    if SPREADSHEET is None:
        client = init_google_client()
        SPREADSHEET = client.open("JR and Co.")  # Updated sheet name here.
    return SPREADSHEET

def get_material_inventory_ws():
    global MATERIAL_INVENTORY_WS
    if MATERIAL_INVENTORY_WS is None:
        MATERIAL_INVENTORY_WS = open_sheet().worksheet("Material Inventory")
    return MATERIAL_INVENTORY_WS

# ===================== Order Entry App =====================

class OrderEntryApp(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.company_var = tk.StringVar()
        self.referral_var = tk.StringVar()   # Referral field (not submitted)
        self.design_var = tk.StringVar()
        self.quantity_var = tk.StringVar()
        self.product_var = tk.StringVar()
        self.due_date_var = tk.StringVar()
        self.price_var = tk.StringVar()
        self.date_type_var = tk.StringVar()
        self.material1_var = tk.StringVar()
        self.material2_var = tk.StringVar()
        self.material3_var = tk.StringVar()
        self.material4_var = tk.StringVar()
        self.material5_var = tk.StringVar()
        self.back_material_var = tk.StringVar()
        self.fur_color_var = tk.StringVar()
        self.backing_type_var = tk.StringVar()
        self.notes_var = tk.StringVar()
        # Use lists for multiple files:
        self.prod_file_paths = []
        self.print_file_paths = []

        # Embedded progress/status area
        self.loading_frame = tk.Frame(self)
        self.loading_label = tk.Label(self.loading_frame, text="Submitting order, please wait...")
        self.loading_progress = ttk.Progressbar(self.loading_frame, mode='determinate', maximum=100, value=0)
        self.loading_label.pack(side="left", padx=5, pady=5)
        self.loading_progress.pack(side="left", padx=5, pady=5)
        self.loading_frame.grid(row=5, column=0, columnspan=2, sticky="ew")
        self.loading_frame.grid_remove()

        self.grid_columnconfigure(0, weight=3)
        self.grid_columnconfigure(1, weight=1)

        self.left_frame = tk.Frame(self)
        self.left_frame.grid(row=0, column=0, sticky="nsew")
        self.build_order_details_frame(self.left_frame)
        self.build_materials_frame(self.left_frame)
        self.build_additional_info_frame(self.left_frame)

        self.right_frame = tk.Frame(self)
        self.right_frame.grid(row=0, column=1, sticky="nsew")
        self.build_file_previews(self.right_frame)

    def submit_order(self):
        # Order entry only submits to the Production Orders sheet now, no material logging
        self.show_loading()
        try:
            self.loading_progress.config(value=10)
            self.update_idletasks()
            due_date_str = self.due_date_var.get().strip()
            if not due_date_str:
                raise Exception("Due Date is required.")
            try:
                current_year = datetime.date.today().year
                due_date = datetime.datetime.strptime(due_date_str, '%m/%d').date().replace(year=current_year)
                if due_date < datetime.date.today():
                    due_date = due_date.replace(year=current_year + 1)
            except ValueError:
                raise Exception("Invalid due date format. Use m/dd or mm/dd.")
            
            for field, value in [("Company Name", self.company_var.get()),
                                 ("Design Name", self.design_var.get().strip()),
                                 ("Quantity", self.quantity_var.get().strip()),
                                 ("Product", self.product_var.get()),
                                 ("Due Date", self.due_date_var.get().strip()),
                                 ("Price", self.price_var.get())]:
                if not value:
                    raise Exception(f"Provide a valid {field}.")
            
            # Now, we only submit the order to the Production Orders sheet
            prod_orders_ws = open_sheet().worksheet("Production Orders")
            existing_rows = prod_orders_ws.col_values(1)
            next_empty_row = len(existing_rows) + 1
            order_folder_name = str(next_empty_row)
            order_folder_id, _ = create_drive_folder(order_folder_name)
            self.loading_progress.config(value=50)

            # Production file upload handling
            prod_file_link = ""
            if self.prod_file_paths:
                if len(self.prod_file_paths) == 1:
                    _, prod_file_link = upload_file_to_drive(self.prod_file_paths[0], order_folder_id)
                else:
                    prod_folder_id, prod_folder_link = create_drive_folder("Production Files", parent_id=order_folder_id)
                    for fp in self.prod_file_paths:
                        upload_file_to_drive(fp, prod_folder_id)
                    prod_file_link = prod_folder_link

            print_file_link = ""
            if self.print_file_paths:
                print_folder_id, print_folder_link = create_drive_folder("Print Files", parent_id=order_folder_id)
                for fp in self.print_file_paths:
                    upload_file_to_drive(fp, print_folder_id)
                print_file_link = print_folder_link

            headers = prod_orders_ws.row_values(1)
            mapping = {
                "Company Name": "Company Name",
                "Design Name": "Design",
                "Due Date": "Due Date",
                "Quantity": "Quantity",
                "Product": "Product",
                "Price": "Price",
            }

            for order_field, sheet_header in mapping.items():
                col_index = find_exact_header_index(headers, sheet_header)
                if col_index is not None:
                    prod_orders_ws.update_cell(next_empty_row, col_index + 1, self.company_var.get())
                else:
                    messagebox.showwarning("Warning", f"Header '{sheet_header}' not found in Production Orders tab.")

            messagebox.showinfo("Success", "Order submitted successfully!")
            self.clear_fields()

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")
        finally:
            self.hide_loading()
