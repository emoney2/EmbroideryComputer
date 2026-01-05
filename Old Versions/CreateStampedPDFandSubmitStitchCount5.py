# CreateStampedPDFandSubmitStitchCount4.py

# ================================
# SECTION 1: Imports
# ================================
import os
import time
import re
import urllib.parse
from io import BytesIO
from datetime import datetime
import threading

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
from matplotlib import colors as mcolors
from functools import lru_cache

import keyboard
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

# ================================
# SECTION 2: Configuration & Constants
# ================================
CREDENTIALS_PATH   = r"C:\Users\eckar\Desktop\OrderEntry,Inventory,QR,PrintPDF\credentials.json"
MONITOR_FOLDER     = r"C:\Users\eckar\Desktop\Embroidery Sheets"
OUTPUT_FOLDER      = r"C:\Users\eckar\Desktop\Embroidery Sheets\PrintPDF"
SHORTENER_SERVICE  = pyshorteners.Shortener()

SHEET_NAME         = "JR and Co."
TAB_NAME           = "Thread Data"

# ←— your latest deployment URL:
WEBAPP_URL         = "https://script.google.com/macros/s/AKfycbyeDUIpCqiicPfmh9hLkPBxxXt9o0aMfezUj8-jCcsrAXay6c2ZyZJgs3IKHmWC8oSdGA/exec"
EMB_START_URL      = f"{WEBAPP_URL}?event=machine_start&order="
DATA_URL           = f"{WEBAPP_URL}?data="

# scanner timing threshold (seconds)
TIME_THRESHOLD     = 0.1
_accumulated       = ""
_last_time         = 0.0

# ================================
# SECTION 3: Utility Functions
# ================================
def shorten_url(url):
    try:
        return SHORTENER_SERVICE.tinyurl.short(url)
    except Exception:
        return url

def clean_value(value):
    return str(value).lstrip("'")

def wait_for_stable_file(path, max_wait=10):
    last = -1
    stable = 0
    start = time.time()
    while time.time() - start < max_wait:
        if not os.path.exists(path):
            time.sleep(0.5)
            continue
        now = os.path.getsize(path)
        if now == last:
            stable += 1
            if stable >= 2:
                return True
        else:
            last = now
            stable = 0
        time.sleep(0.5)
    return False

def color_to_rgb(color):
    c = re.sub(r"\s*fur\s*$", '', color.lower()).strip()
    try:
        frac = mcolors.to_rgb(c.replace(' ', ''))
        return tuple(int(255 * comp) for comp in frac)
    except Exception:
        return (255, 255, 255)

def get_contrast_color(color):
    r, g, b = color_to_rgb(color)
    lum = (r * 299 + g * 587 + b * 114) / 1000
    return (0, 0, 0) if lum > 127 else (255, 255, 255)

# ================================
# SECTION 4: Google Sheets Functions
# ================================
def connect_google_sheet(sheet_name, tab_name=TAB_NAME):
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_file(CREDENTIALS_PATH, scopes=scope)
    client = gspread.authorize(creds)
    ss = client.open(sheet_name)
    return ss.worksheet(tab_name), ss

def update_sheet(sheet, data):
    if not data:
        return
    order_key = clean_value(data[0][0]).strip().lower().replace('.0', '')
    headers = sheet.row_values(1)
    hmap = {h.strip().lower(): i for i, h in enumerate(headers)}
    if 'order number' not in hmap:
        return

    # delete prior rows
    col = hmap['order number'] + 1
    vals = sheet.col_values(col)
    to_delete = [
        i for i, v in enumerate(vals[1:], start=2)
        if str(v).strip().lower().replace('.0', '') == order_key
    ]
    reqs = []
    for r in sorted(to_delete, reverse=True):
        reqs.append({
            'deleteDimension': {
                'range': {
                    'sheetId': sheet.id,
                    'dimension': 'ROWS',
                    'startIndex': r - 1,
                    'endIndex': r
                }
            }
        })
    replaced = bool(reqs)
    if reqs:
        try:
            sheet.spreadsheet.batch_update({'requests': reqs})
            time.sleep(1)
        except Exception as e:
            print('Batch delete error:', e)

    # append new data
    ts = datetime.now().strftime('%m/%d/%Y %H:%M:%S')
    for row in data:
        new = [''] * len(headers)
        if 'date' in hmap:
            new[hmap['date']] = ts
        if 'order number' in hmap:
            new[hmap['order number']] = str(row[0]).strip()
        if 'color' in hmap:
            new[hmap['color']] = row[1]
        if 'color name' in hmap:
            new[hmap['color name']] = row[2]
        if 'length (ft)' in hmap:
            new[hmap['length (ft)']] = row[3]
        if 'stitch count' in hmap:
            new[hmap['stitch count']] = row[4]
        if 'in/out' in hmap:
            new[hmap['in/out']] = 'OUT'
        if 'o/r' in hmap:
            new[hmap['o/r']] = ''
        sheet.append_row(new, value_input_option='USER_ENTERED')

    msg = 'Data Replaced' if replaced else 'Data Recorded'
    try:
        sheet.update_acell('Z1', msg)
    except Exception as e:
        print('Status cell error:', e)
    notification.notify(title='PDF Data Extraction', message=msg, timeout=5)

# ================================
# SECTION 5: Cached Getters
# ================================
@lru_cache()
def get_order_quantity(order_number, spreadsheet):
    key = clean_value(order_number).strip().lower()
    try:
        ws   = spreadsheet.worksheet('Production Orders')
        vals = ws.get_all_values()
        hdr  = vals[0]
        idx  = next((i for i,h in enumerate(hdr) if h.strip().lower()=='quantity'), None)
        if idx is None: return 1.0
        for row in vals[1:]:
            if row and clean_value(row[0]).strip().lower() == key:
                return float(row[idx])
    except Exception as e:
        print('get_order_quantity error:', e)
    return 1.0

@lru_cache()
def get_fur_color(order_number, spreadsheet):
    key = clean_value(order_number).strip().lower()
    try:
        ws   = spreadsheet.worksheet('Production Orders')
        vals = ws.get_all_values()
        hdr  = vals[0]
        idx  = next((i for i,h in enumerate(hdr) if 'fur color' in h.strip().lower()), None)
        if idx is None: return ''
        for row in vals[1:]:
            if row and clean_value(row[0]).strip().lower() == key:
                return str(row[idx])
    except Exception as e:
        print('get_fur_color error:', e)
    return ''

@lru_cache()
def get_due_date(order_number, spreadsheet):
    key = clean_value(order_number).strip().lower()
    try:
        ws   = spreadsheet.worksheet('Production Orders')
        vals = ws.get_all_values()
        hdr  = vals[0]
        idx  = next((i for i,h in enumerate(hdr) if 'ship date' in h.strip().lower()), None)
        if idx is None: return ''
        for row in vals[1:]:
            if row and clean_value(row[0]).strip().lower() == key:
                return str(row[idx])
    except Exception as e:
        print('get_due_date error:', e)
    return ''

@lru_cache()
def get_product(order_number, spreadsheet):
    key = clean_value(order_number).strip().lower()
    try:
        ws   = spreadsheet.worksheet('Production Orders')
        vals = ws.get_all_values()
        hdr  = vals[0]
        idx  = next((i for i,h in enumerate(hdr) if h.strip().lower()=='product'), None)
        if idx is None: return ''
        for row in vals[1:]:
            if row and clean_value(row[0]).strip().lower() == key:
                return str(row[idx])
    except Exception as e:
        print('get_product error:', e)
    return ''

@lru_cache()
def get_notes(order_number, spreadsheet):
    """
    Return the Notes cell as a string for the given order.
    """
    key = clean_value(order_number).strip().lower()
    try:
        ws   = spreadsheet.worksheet('Production Orders')
        vals = ws.get_all_values()
        hdr  = [h.strip().lower() for h in vals[0]]
        idx  = next((i for i,h in enumerate(hdr) if h == 'notes'), None)
        if idx is None: return ''
        for row in vals[1:]:
            if row and clean_value(row[0]).strip().lower() == key:
                return str(row[idx] or '')
    except Exception as e:
        print('get_notes error:', e)
    return ''

# ================================
# SECTION 6: PDF Extraction
# ================================
def extract_thread_usage(pdf_path):
    data = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if not pdf.pages:
                return data
            text = pdf.pages[0].extract_text() or ""
            m = re.search(r"Stitches:\s*([\d,]+)", text)
            stitch_count = float(m.group(1).replace(",", "")) if m else 0.0
            lines = text.splitlines()
            hdr = next((i for i, l in enumerate(lines) if "N# Color Name Length" in l), None)
            if hdr is None:
                return data
            for line in lines[hdr+1:]:
                if not line or not line[0].isdigit():
                    continue
                m2 = re.search(r"(\d+)\.\s+(\d+)\s+([\w\s]+?)\s+([\d\.]+)ft", line)
                if m2:
                    data.append([
                        clean_value(m2.group(1)),
                        clean_value(m2.group(2)),
                        clean_value(m2.group(3).strip()),
                        float(m2.group(4)),
                        stitch_count
                    ])
    except Exception as e:
        print("extract_thread_usage error:", e)
    return data

# ================================
# SECTION 7: PDF Stamping & QR Codes
from io import BytesIO

def draw_top_right_box(pdf, order_number, spreadsheet, w, h, s):
    box_w, box_h = 60, 120
    x0, y0 = w - box_w - 10, 10
    pdf.set_line_width(1)
    pdf.set_draw_color(0,0,0)
    pdf.rect(x0, y0, box_w, box_h)

    content_x, inner_w = x0 + 2, box_w - 4
    rows = [20, 20, 20, 14, 24, 22]
    offs = [5, 6, 6, 3, 6, 6]
    starts = [y0]
    for r in rows[:-1]:
        starts.append(starts[-1] + r)
    ys = [starts[i] + offs[i] for i in range(len(rows))]

    pdf.set_font('Arial','BU',10)
    pdf.set_text_color(0,0,0)
    pdf.set_xy(content_x, ys[0])
    pdf.cell(inner_w, rows[0], 'Sewing', 0, 2, 'C')

    pdf.set_font('Arial','B',9)
    pdf.set_xy(content_x, ys[1])
    pdf.cell(inner_w, rows[1], f"QTY: {int(get_order_quantity(order_number, spreadsheet))}", 0, 2, 'C')

    fur = re.sub(r"\s*fur\s*$", '', get_fur_color(order_number, spreadsheet), flags=re.IGNORECASE).strip()
    rgb, contrast = color_to_rgb(fur), get_contrast_color(fur)

    pdf.set_font('Arial','',8)
    pdf.set_fill_color(*rgb)
    pdf.rect(content_x, starts[2], inner_w, rows[2], style='F')
    pdf.set_text_color(*contrast)
    pdf.set_xy(content_x, ys[2])
    pdf.cell(inner_w, rows[2], fur, 0, 2, 'C')

    pdf.set_fill_color(*rgb)
    pdf.rect(content_x, starts[3], inner_w, rows[3], style='F')
    pdf.set_text_color(*contrast)
    pdf.set_xy(content_x, ys[3])
    pdf.cell(inner_w, rows[3], 'Fur', 0, 2, 'C')

    pdf.set_font('Arial','B',12)
    pdf.set_text_color(0,0,0)
    pdf.set_xy(content_x, ys[4])
    pdf.cell(inner_w, rows[4], get_product(order_number, spreadsheet), 0, 2, 'C')

    pdf.set_font('Arial','B',10)
    pdf.set_xy(content_x, ys[5])
    pdf.cell(inner_w, rows[5], f"Ship: {get_due_date(order_number, spreadsheet)}", 0, 2, 'C')


def shrink_page_and_stamp_horizontal_qrs(original_pdf_path, order_number, spreadsheet, margin_pts=72):
    qty = get_order_quantity(order_number, spreadsheet)
    try:
        # 1) Load PDF into memory
        with open(original_pdf_path, 'rb') as f:
            raw = f.read()
        reader = PdfFileReader(BytesIO(raw))
        writer = PdfFileWriter()

        # original dimensions
        page0 = reader.getPage(0)
        w, h = float(page0.mediaBox.getWidth()), float(page0.mediaBox.getHeight())

        # stamping footprints
        pad     = 10
        size    = 54
        label_h = 10
        box_w   = 60
        box_h   = 120

        left_margin   = pad
        right_margin  = box_w + pad
        top_margin    = pad + label_h + int(size * 0.75)
        bottom_margin = margin_pts

        # compute max area and scale
        max_w = w - left_margin - right_margin
        max_h = h - top_margin - bottom_margin
        scale = min(0.90, max_w / w, max_h / h)

        # transform original page
        tx = left_margin
        ty = bottom_margin
        page0.addTransformation([scale, 0, 0, scale, tx, ty])

        writer.addPage(page0)
        for i in range(1, reader.getNumPages()):
            writer.addPage(reader.getPage(i))

        # 2) Build overlay
        pdf = FPDF(unit='pt', format=(w, h))
        pdf.set_auto_page_break(False)
        pdf.add_page()
        pdf.set_font('Arial','B',8)

        # Top QR row
        top_size = int(size * 0.75)
        y_lbl, y_img = pad, pad + label_h
        for idx in range(5):
            x = ((w - 5 * top_size) / 6) * (idx + 1) + top_size * idx - (36 if idx == 0 else 0)
            if idx == 0:
                url = shorten_url(EMB_START_URL + urllib.parse.quote(order_number))
                qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_L,
                                   box_size=6, border=0)
                qr.add_data(url + '\r'); qr.make(fit=True)
                img = qr.make_image().convert('1', dither=Image.NONE)
                tmp = f"qr_top_{idx}.png"; img.save(tmp)
                pdf.set_xy(x, y_lbl)
                pdf.cell(top_size, label_h, "Machine Start", 0, 2, 'C')
                pdf.image(tmp, x, y_img, w=top_size, h=top_size)
                os.remove(tmp)
            else:
                pdf.set_draw_color(200,200,200)
                pdf.rect(x, y_img, top_size, top_size)

        # Info box
        draw_top_right_box(pdf, order_number, spreadsheet, w, h, scale)

        # Right-side QR column (start below info box)
        x0 = w - box_w - pad
        y_start = pad + box_h + pad
        labels = ['Fur List','Cut List','Print List','Embroidery List','Shipping']
        for idx, label in enumerate(labels):
            raw  = f"{label.replace(' ', '_')}_{order_number}_{qty}"
            url2 = shorten_url(DATA_URL + urllib.parse.quote(raw))
            qr2 = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_L,
                                box_size=6, border=0)
            qr2.add_data(url2 + '\r'); qr2.make(fit=True)
            img2 = qr2.make_image().convert('1', dither=Image.NONE)
            tmp2 = f"qr_bot_{idx}.png"; img2.save(tmp2)

            y_lab = y_start + idx * (label_h + size + pad + 5)
            y_img2 = y_lab + label_h
            pdf.set_xy(x0, y_lab)
            pdf.cell(size, label_h, label, 0, 2, 'C')
            pdf.image(tmp2, x0, y_img2, w=size, h=size)
            os.remove(tmp2)

        # 3) Bottom notes box
        note_pad = 10
        note_h   = bottom_margin - 2 * note_pad
        note_y   = h - bottom_margin + note_pad
        note_x   = pad
        note_w   = w - 2 * pad

        # fill and border
        pdf.set_fill_color(255,255,255)
        pdf.rect(note_x, note_y, note_w, note_h, style='F')
        pdf.set_draw_color(0,0,0)
        pdf.set_line_width(1)
        pdf.rect(note_x, note_y, note_w, note_h, style='D')

        # fetch and debug
        note_text = get_notes(order_number, spreadsheet)
        print(f"[DEBUG] Notes for {order_number}: {note_text!r}")

        # draw text on top
        display = note_text or '-'
        pdf.set_text_color(0,0,0)
        pdf.set_font('Arial','',10)
        pdf.set_xy(note_x + 5, note_y + 5)
        pdf.multi_cell(note_w - 10, 12, f"Notes: {display}")

        # 4) Merge and write
        overlay_stream = BytesIO(pdf.output(dest='S').encode('latin-1'))
        overlay_reader = PdfFileReader(overlay_stream)
        overlay_page   = overlay_reader.getPage(0)
        target         = writer.getPage(0)
        target.merge_page(overlay_page)

        os.makedirs(OUTPUT_FOLDER, exist_ok=True)
        outp = os.path.join(
            OUTPUT_FOLDER,
            os.path.splitext(os.path.basename(original_pdf_path))[0] + '_Stamped.pdf'
        )
        with open(outp, 'wb') as of:
            writer.write(of)

        os.remove(original_pdf_path)

    except Exception as e:
        print('stamping error:', e)



# ================================
# SECTION 8: File Monitoring
# ================================
processed = set()
class PDFHandler(FileSystemEventHandler):
    def on_created(self, event):
        if event.is_directory or not event.src_path.lower().endswith('.pdf'):
            return
        p = event.src_path
        if p in processed:
            return
        if not wait_for_stable_file(p):
            return
        processed.add(p)
        order = clean_value(os.path.splitext(os.path.basename(p))[0])
        usage = extract_thread_usage(p)
        if not usage:
            return
        rows = [
            [order, r[1], r[2], r[3] * get_order_quantity(order, sheet_spreadsheet), r[4]]
            for r in usage
        ]
        update_sheet(sheet_thread, rows)
        shrink_page_and_stamp_horizontal_qrs(p, order, sheet_spreadsheet)

def monitor_folder(folder, sh, ss):
    global sheet_thread, sheet_spreadsheet
    sheet_thread, sheet_spreadsheet = sh, ss
    os.makedirs(folder, exist_ok=True)
    obs = Observer()
    obs.schedule(PDFHandler(), folder, recursive=False)
    obs.start()
    print(f"Monitoring folder: {folder}")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        obs.stop()
        obs.join()

# ================================
# SECTION 9: Scanner Listener
# ================================
def flash_color_for_duration(driver, color, duration, interval=0.3):
    start = time.time()
    while time.time() - start < duration:
        driver.execute_script(
            "document.body.style.transition='background-	color 0.2s';"
            "document.body.style.backgroundColor = arguments[0];", color
        )
        time.sleep(interval)
        driver.execute_script("document.body.style.backgroundColor = 'white';")
        time.sleep(interval)

def flash_color_continuous(driver, color, interval=0.3):
    try:
        while True:
            driver.execute_script(
                "document.body.style.transition='background-	color 0.2s';"
                "document.body.style.backgroundColor = arguments[0];", color
            )
            time.sleep(interval)
            driver.execute_script("document.body.style.backgroundColor = 'white';")
            time.sleep(interval)
    except:
        pass

def open_and_handle_url(url):
    options = Options()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=options)
    driver.get(url)
    driver.maximize_window()
    driver.execute_script("window.focus();")

    success = False
    elapsed = 0
    while elapsed < 5:
        if "Success" in driver.page_source:
            success = True
            break
        time.sleep(0.5)
        elapsed += 0.5

    if success:
        flash_color_for_duration(driver, "green", duration=4, interval=0.3)
        driver.quit()
    else:
        flash_color_continuous(driver, "red", interval=0.3)

def on_key_event(event):
    global _accumulated, _last_time
    if event.event_type != "down":
        return
    now = time.time()
    if now - _last_time > TIME_THRESHOLD:
        _accumulated = ""
    _last_time = now

    if event.name == "enter":
        candidate = _accumulated.strip()
        _accumulated = ""
        if candidate.startswith(WEBAPP_URL) or candidate.startswith("https://tinyurl.com/"):
            print("Launching URL:", candidate)
            open_and_handle_url(candidate)
        else:
            keyboard.send("enter")
    else:
        if len(event.name) == 1:
            _accumulated += event.name
        elif event.name == "space":
            _accumulated += " "

# ================================
# SECTION 10: Main
# ================================
if __name__ == "__main__":
    sheet_thread, sheet_spreadsheet = connect_google_sheet(SHEET_NAME, TAB_NAME)
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    folder_thread = threading.Thread(
        target=monitor_folder,
        args=(MONITOR_FOLDER, sheet_thread, sheet_spreadsheet),
        daemon=True
    )
    folder_thread.start()

    keyboard.hook(on_key_event)
    print("QR scanner listener started. Scan a QR code…")
    keyboard.wait()
