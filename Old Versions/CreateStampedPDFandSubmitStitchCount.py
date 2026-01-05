import os
import time
import re
import urllib.parse
from io import BytesIO
from datetime import datetime  # For timestamp

import pdfplumber
import gspread
import qrcode
from fpdf import FPDF
from PyPDF2 import PdfFileReader, PdfFileWriter
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from google.oauth2.service_account import Credentials
from plyer import notification
from PIL import Image
import pyshorteners

# Import matplotlib for dynamic color conversion.
from matplotlib import colors as mcolors

# ----------------------------
# Google Sheet & PDF Processing Functions
# ----------------------------

def shorten_url(url):
    s = pyshorteners.Shortener()
    try:
        short_url = s.tinyurl.short(url)
        print(f"Shortened URL: {short_url}")
        return short_url
    except Exception as e:
        print("Error shortening URL:", e)
        return url

def clean_value(value):
    return str(value).lstrip("'")

def connect_google_sheet(sheet_name, tab_name="Thread Data"):
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_file(
        r"C:\Users\eckar\Desktop\Working Folder for Python Scripts\credentials.json", scopes=scope
    )
    client = gspread.authorize(creds)
    spreadsheet = client.open(sheet_name)
    tab = spreadsheet.worksheet(tab_name)
    return tab, spreadsheet

def update_sheet(sheet, data):
    if not data:
        return

    order_to_update = clean_value(data[0][0]).strip().lower().replace(".0", "")
    print(f"Updating rows for order number: {order_to_update}")

    header_row = sheet.row_values(1)
    header_mapping = {header.strip().lower(): idx for idx, header in enumerate(header_row)}
    
    if "order number" not in header_mapping:
        print("No 'Order Number' header found in the sheet.")
        return

    target_col = header_mapping["order number"] + 1  # 1-based index
    print(f"Target column for Order Number is: {target_col}")
    print(f"Sheet ID is: {sheet.id}")

    col_vals = sheet.col_values(target_col)
    rows_to_delete = []
    for i, cell_val in enumerate(col_vals[1:], start=2):
        cell_str = str(cell_val).strip().lower().replace(".0", "")
        if cell_str == order_to_update:
            rows_to_delete.append(i)
            print(f"Found matching row {i} with value: {cell_val}")

    replaced = len(rows_to_delete) > 0

    delete_requests = []
    for row_num in sorted(rows_to_delete, reverse=True):
        req = {
            "deleteDimension": {
                "range": {
                    "sheetId": sheet.id,
                    "dimension": "ROWS",
                    "startIndex": row_num - 1,
                    "endIndex": row_num
                }
            }
        }
        delete_requests.append(req)
        print(f"Scheduling deletion for row {row_num}")
    
    if delete_requests:
        try:
            print("Attempting batch deletion using spreadsheet.batch_update()")
            sheet.spreadsheet.batch_update({"requests": delete_requests})
            time.sleep(1)
            print(f"Batch deletion succeeded; deleted {len(delete_requests)} matching row(s).")
        except Exception as e:
            print("Error during batch deletion:", e)
    
    timestamp = datetime.now().strftime("%m/%d/%Y %H:%M:%S")
    for row in data:
        new_row = ["" for _ in range(len(header_row))]
        if "date" in header_mapping:
            new_row[header_mapping["date"]] = timestamp
        if "order number" in header_mapping:
            new_row[header_mapping["order number"]] = str(row[0]).strip()
        if "color" in header_mapping:
            new_row[header_mapping["color"]] = row[1]
        if "color name" in header_mapping:
            new_row[header_mapping["color name"]] = row[2]
        if "length (ft)" in header_mapping:
            new_row[header_mapping["length (ft)"]] = row[3]
        if "stitch count" in header_mapping:
            new_row[header_mapping["stitch count"]] = row[4]
        if "in/out" in header_mapping:
            new_row[header_mapping["in/out"]] = "OUT"
        if "o/r" in header_mapping:
            new_row[header_mapping["o/r"]] = ""
        sheet.append_row(new_row, value_input_option="USER_ENTERED")
    
    msg = "Data Replaced" if replaced else "Data Recorded"
    try:
        sheet.update_acell("Z1", msg)
    except Exception as e:
        print("Error updating status cell:", e)
    notification.notify(title="PDF Data Extraction", message=msg, timeout=5)

def extract_thread_usage(pdf_path):
    data = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if not pdf.pages:
                return data
            page = pdf.pages[0]
            text = page.extract_text() or ""
            m_stitch = re.search(r"Stitches:\s*([\d,]+)", text)
            stitch_count = float(m_stitch.group(1).replace(",", "")) if m_stitch else 0.0
            lines = text.splitlines()
            header_idx = next((i for i, line in enumerate(lines) if "N# Color Name Length" in line), None)
            if header_idx is None:
                return data
            for line in lines[header_idx+1:]:
                if not line or not line[0].isdigit():
                    continue
                m = re.search(r"(\d+)\.\s+(\d+)\s+([\w\s]+?)\s+([\d\.]+ft)", line)
                if m:
                    data.append([
                        clean_value(m.group(1)),
                        clean_value(m.group(2)),
                        clean_value(m.group(3).strip()),
                        float(m.group(4).replace("ft", "")),
                        stitch_count
                    ])
    except Exception as e:
        print("Error in extract_thread_usage:", e)
    return data

def get_header_aa_to_ae(order_number, spreadsheet):
    order_norm = order_number.strip().lower()
    sheet = spreadsheet.worksheet("Production Orders")
    vals = sheet.get_all_values()
    headers = vals[0][26:31] if len(vals[0]) >= 31 else []
    data_row = next((row for row in vals[1:] if row and row[0].strip().lower() == order_norm), None)
    values = data_row[26:31] if data_row and len(data_row) >= 31 else []
    return headers, values

def get_order_quantity(order_number, spreadsheet):
    order_norm = order_number.strip().lower()
    try:
        prod_sheet = spreadsheet.worksheet("Production Orders")
        all_vals = prod_sheet.get_all_values()
        if not all_vals:
            return 1.0
        header = all_vals[0]
        qty_col_idx = next((i for i, h in enumerate(header) if h.strip().lower() == "quantity"), None)
        if qty_col_idx is None:
            return 1.0
        for row in all_vals[1:]:
            if row and row[0].strip().lower() == order_norm:
                return float(row[qty_col_idx])
    except Exception as e:
        print("Error in get_order_quantity:", e)
    return 1.0

def get_fur_color(order_number, spreadsheet):
    order_norm = order_number.strip().lower()
    try:
        prod_sheet = spreadsheet.worksheet("Production Orders")
        all_vals = prod_sheet.get_all_values()
        if not all_vals:
            return ""
        header = all_vals[0]
        fur_col_idx = next((i for i, h in enumerate(header) if "fur color" in h.strip().lower()), None)
        if fur_col_idx is None:
            return ""
        for row in all_vals[1:]:
            if row and row[0].strip().lower() == order_norm:
                return row[fur_col_idx]
    except Exception as e:
        print("Error in get_fur_color:", e)
    return ""

def get_due_date(order_number, spreadsheet):
    order_norm = order_number.strip().lower()
    try:
        prod_sheet = spreadsheet.worksheet("Production Orders")
        all_vals = prod_sheet.get_all_values()
        if not all_vals:
            return ""
        header = all_vals[0]
        due_col_idx = next((i for i, h in enumerate(header) if "ship date" in h.strip().lower()), None)
        if due_col_idx is None:
            return ""
        for row in all_vals[1:]:
            if row and row[0].strip().lower() == order_norm:
                return row[due_col_idx]
    except Exception as e:
        print("Error in get_due_date:", e)
    return ""

def get_product(order_number, spreadsheet):
    order_norm = order_number.strip().lower()
    try:
        prod_sheet = spreadsheet.worksheet("Production Orders")
        all_vals = prod_sheet.get_all_values()
        if not all_vals:
            return ""
        header = all_vals[0]
        product_col_idx = next((i for i, h in enumerate(header) if "product" in h.strip().lower()), None)
        if product_col_idx is None:
            return ""
        for row in all_vals[1:]:
            if row and row[0].strip().lower() == order_norm:
                return row[product_col_idx]
    except Exception as e:
        print("Error in get_product:", e)
    return ""

def color_to_rgb(color):
    """
    Dynamically converts a color string into an (R, G, B) tuple using matplotlib.
    """
    c = color.lower().strip()
    c = re.sub(r'\s*fur\s*$', '', c).strip()
    c_css = c.replace(" ", "")
    try:
        rgb_frac = mcolors.to_rgb(c_css)
        rgb = tuple(int(255 * comp) for comp in rgb_frac)
        return rgb
    except Exception as e:
        print(f"Error converting '{color}' to RGB:", e)
        return (255, 255, 255)

def get_contrast_color(color):
    r, g, b = color_to_rgb(color)
    brightness = (r * 299 + g * 587 + b * 114) / 1000
    return (0, 0, 0) if brightness > 127 else (255, 255, 255)

def draw_top_right_box(pdf, order_number, spreadsheet, page_width, page_height, scale):
    bounding_box_width = 60   # Outer bounding box width = 60 pts
    content_box_width = 58    # Inner content width (2 pts inset total)
    desired_box_height = 120  # Total height remains 120 pts

    x = page_width - bounding_box_width - 10
    y = 10

    pdf.set_line_width(1)
    pdf.set_draw_color(0, 0, 0)
    pdf.rect(x, y, bounding_box_width, desired_box_height)

    content_x = x + 2
    inner_width = content_box_width

    row_heights = [20, 20, 20, 14, 24, 22]
    row_offsets = [5, 6, 6, 3, 6, 6]

    row_starts = [y]
    for h in row_heights[:-1]:
        row_starts.append(row_starts[-1] + h)
    text_y = [row_starts[i] + row_offsets[i] for i in range(len(row_heights))]

    pdf.set_font("Arial", "BU", 10)
    pdf.set_text_color(0, 0, 0)
    pdf.set_xy(content_x, text_y[0])
    pdf.cell(w=inner_width, h=row_heights[0], txt="Sewing", border=0, ln=2, align="C")

    pdf.set_font("Arial", "B", 9)
    pdf.set_text_color(0, 0, 0)
    pdf.set_xy(content_x, text_y[1])
    qty = int(get_order_quantity(order_number, spreadsheet))
    pdf.cell(w=inner_width, h=row_heights[1], txt=f"QTY: {qty}", border=0, ln=2, align="C")

    pdf.set_font("Arial", "", 8)
    fur_full = get_fur_color(order_number, spreadsheet)
    fur_text = re.sub(r'\s*fur\s*$', '', fur_full, flags=re.IGNORECASE).strip()
    print(f"Fur color read from sheet: '{fur_text}' -> {color_to_rgb(fur_text)}")
    pdf.set_fill_color(*color_to_rgb(fur_text) if fur_text else (255, 255, 255))
    pdf.rect(content_x, row_starts[2], inner_width, row_heights[2], style="F")
    contrast = get_contrast_color(fur_text) if fur_text else (0, 0, 0)
    pdf.set_text_color(*contrast)
    pdf.set_xy(content_x, text_y[2])
    pdf.cell(w=inner_width, h=row_heights[2], txt=fur_text, border=0, ln=2, align="C")

    pdf.set_font("Arial", "", 8)
    pdf.set_fill_color(*color_to_rgb(fur_text) if fur_text else (255, 255, 255))
    pdf.rect(content_x, row_starts[3], inner_width, row_heights[3], style="F")
    pdf.set_text_color(*contrast)
    pdf.set_xy(content_x, text_y[3])
    pdf.cell(w=inner_width, h=row_heights[3], txt="Fur", border=0, ln=2, align="C")

    pdf.set_font("Arial", "B", 12)
    pdf.set_text_color(0, 0, 0)
    product = get_product(order_number, spreadsheet)
    pdf.set_xy(content_x, text_y[4])
    pdf.cell(w=inner_width, h=row_heights[4], txt=f"{product}", border=0, ln=2, align="C")

    pdf.set_font("Arial", "B", 10)
    pdf.set_text_color(0, 0, 0)
    ship_date = get_due_date(order_number, spreadsheet)
    pdf.set_xy(content_x, text_y[5])
    pdf.cell(w=inner_width, h=row_heights[5], txt=f"Ship: {ship_date}", border=0, ln=2, align="C")

def get_product(order_number, spreadsheet):
    order_norm = order_number.strip().lower()
    try:
        prod_sheet = spreadsheet.worksheet("Production Orders")
        all_vals = prod_sheet.get_all_values()
        if not all_vals:
            return ""
        header = all_vals[0]
        product_col_idx = next((i for i, h in enumerate(header) if "product" in h.strip().lower()), None)
        if product_col_idx is None:
            return ""
        for row in all_vals[1:]:
            if row and row[0].strip().lower() == order_norm:
                return row[product_col_idx]
    except Exception as e:
        print("Error in get_product:", e)
    return ""

def shrink_page_and_stamp_horizontal_qrs(original_pdf_path, order_number, spreadsheet, bottom_margin_pts=72):
    qty = get_order_quantity(order_number, spreadsheet)
    base_url = ("https://script.google.com/macros/s/AKfycbwcw3yzL2HSCbwSmykvP8PqsfTIwm8Rmm5MdKSDoFA8yZiszd4AFzX6HijxUj84EZC8Zg/"
                "exec?data=")
    try:
        extra_bottom = 18
        new_bottom_margin = bottom_margin_pts + extra_bottom
        with open(original_pdf_path, "rb") as f:
            reader = PdfFileReader(f)
            writer = PdfFileWriter()
            page = reader.getPage(0)
            w, h = float(page.mediaBox.getWidth()), float(page.mediaBox.getHeight())
            print("Page width:", w, "Page height:", h)
            print("New bottom margin (pts):", new_bottom_margin)
            s = (h - new_bottom_margin) / h
            tx = 0
            ty = new_bottom_margin
            page.addTransformation([s, 0, 0, s, tx, ty])
            writer.addPage(page)
            for i in range(1, reader.getNumPages()):
                writer.addPage(reader.getPage(i))
            pdf = FPDF(unit="pt", format=(w, h))
            pdf.add_page()
            pdf.set_font("Arial", style="B", size=8)
            code_size = 54
            labels = ["Fur List", "Cut List", "Print List", "Embroidery List", "Shipping"]
            gap = (w - (5 * code_size)) / 6
            bottom = 10
            order_clean = clean_value(order_number.strip().replace(" ", ""))
            shift_qr = 19
            for i, label in enumerate(labels):
                if label.lower() == "shipping":
                    qr_data_raw = f"Production Orders_{order_clean}_{qty}"
                else:
                    qr_data_raw = f"{label.replace(' ', '_')}_{order_clean}_{qty}"
                qr_encoded = urllib.parse.quote(qr_data_raw)
                full_url = f"{base_url}{qr_encoded}"
                short_url = shorten_url(full_url)
                qr = qrcode.QRCode(
                    version=1,
                    error_correction=qrcode.constants.ERROR_CORRECT_L,
                    box_size=6,
                    border=0
                )
                qr.add_data(short_url)
                qr.make(fit=True)
                qr_img = qr.make_image(fill_color="black", back_color="white").convert('1', dither=Image.NONE)
                temp = f"qr_{i}.png"
                qr_img.save(temp)
                x = gap * (i + 1) + code_size * i
                pdf.set_xy(x=x, y=h - (bottom + code_size + 10 + shift_qr))
                pdf.cell(w=code_size, h=10, txt=label, align='C')
                pdf.image(temp, x=x, y=h - (bottom + code_size + shift_qr), w=code_size, h=code_size)
                os.remove(temp)
            draw_top_right_box(pdf, order_number, sheet_spreadsheet, w, h, s)
            overlay_bytes = pdf.output(dest="S").encode("latin-1")
            overlay = PdfFileReader(BytesIO(overlay_bytes))
            writer.getPage(0).mergePage(overlay.getPage(0))
            output_folder = r"C:\Users\eckar\Desktop\Embroidery Sheets\PrintPDF"
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)
            base = os.path.splitext(os.path.basename(original_pdf_path))[0]
            output_path = os.path.join(output_folder, f"{base}_stamped.pdf")
            with open(output_path, "wb") as out_f:
                writer.write(out_f)
        os.remove(original_pdf_path)
    except Exception as e:
        print("Error in stamping:", e)

def wait_for_stable_file(path, max_wait=10):
    last_size = -1
    stable = 0
    start = time.time()
    while time.time() - start < max_wait:
        if not os.path.exists(path):
            time.sleep(0.5)
            continue
        now = os.path.getsize(path)
        if now == last_size:
            stable += 1
            if stable >= 2:
                return True
        else:
            stable = 0
            last_size = now
        time.sleep(0.5)
    return False

class PDFHandler(FileSystemEventHandler):
    def on_created(self, event):
        if event.is_directory or not event.src_path.lower().endswith(".pdf"):
            return
        path = event.src_path
        if not wait_for_stable_file(path):
            return
        order = clean_value(os.path.splitext(os.path.basename(path))[0])
        print(f"\nNew PDF => {path}")
        usage = extract_thread_usage(path)
        if not usage:
            print("No usage data found or PDF read error. Not recording data or stamping.")
            return
        qty = get_order_quantity(order, sheet_spreadsheet)
        new_rows = [[order, r[1], r[2], r[3] * qty, r[4]] for r in usage if len(r) >= 5]
        if new_rows:
            update_sheet(sheet_thread, new_rows)
            shrink_page_and_stamp_horizontal_qrs(path, order, sheet_spreadsheet)

def monitor_folder(folder_path, sheet, spreadsheet):
    global sheet_thread, sheet_spreadsheet
    sheet_thread, sheet_spreadsheet = sheet, spreadsheet
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
    obs = Observer()
    obs.schedule(PDFHandler(), folder_path, recursive=False)
    obs.start()
    print(f"Monitoring folder: {folder_path}")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        obs.stop()
    obs.join()

if __name__ == "__main__":
    # Update the sheet name here from "Ascend Golf Co." to "JR and Co."
    sheet_name = "JR and Co."
    tab_name = "Thread Data"
    sheet_thread, sheet_spreadsheet = connect_google_sheet(sheet_name, tab_name)
    monitor_folder(r"C:\Users\eckar\Desktop\Embroidery Sheets", sheet_thread, sheet_spreadsheet)
