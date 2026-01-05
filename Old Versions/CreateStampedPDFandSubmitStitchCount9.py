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
import textwrap

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

import requests
import pyperclip

# scanner timing threshold (seconds)
TIME_THRESHOLD = 0.1
_accumulated   = ""
_last_time     = 0.0

def get_company_name(order_number, spreadsheet):
    sheet = spreadsheet.worksheet("Production Orders")
    data = sheet.get_all_records()

    for row in data:
        if str(row.get("Order #")).strip() == str(order_number).strip():
            return row.get("Company Name")

    print(f"[Warning] Company not found for Order #: {order_number}")
    return "Unknown"



# ================================
# SECTION 2: Configuration & Constants
# ================================
CREDENTIALS_PATH   = r"C:\Users\eckar\Desktop\OrderEntry,Inventory,QR,PrintPDF\credentials.json"
MONITOR_FOLDER     = r"C:\Users\eckar\Desktop\Embroidery Sheets"
OUTPUT_FOLDER      = r"C:\Users\eckar\Desktop\Embroidery Sheets\PrintPDF"
SHORTENER_SERVICE  = pyshorteners.Shortener()

SHEET_NAME         = "JR and Co."
TAB_NAME           = "Thread Data"

WEBAPP_URL         = "https://script.google.com/macros/s/AKfycbyeDUIpCqiicPfmh9hLkPBxxXt9o0aMfezUj8-jCcsrAXay6c2ZyZJgs3IKHmWC8oSdGA/exec"
EMB_START_URL      = f"{WEBAPP_URL}?event=machine_start&order="
DATA_URL           = f"{WEBAPP_URL}?data="

# ─── Embedded Queue Service Configuration & Route ───
import subprocess
import os
import time
from flask import Flask, request
from pywinauto import Application

# Folder where your .emb files live:
EMB_FOLDER        = r"C:\Users\eckar\Desktop\EMB"
# Regex matching your EmbroideryStudio window title:
EMB_WINDOW_TITLE  = r"EmbroideryStudio\s*2025.*"
# Hide launched process window:
CREATE_NO_WINDOW  = 0x08000000

app = Flask(__name__)

@app.route("/queue")
def queue_design():
    # 1) Get the .emb filename
    file_base = request.args.get("file")
    if not file_base:
        return "Error: no file specified", 400

    # 2) Locate the .emb file
    emb_path = os.path.join(EMB_FOLDER, f"{file_base}.emb")
    if not os.path.exists(emb_path):
        return f"Error: {emb_path} not found", 404

    # 3) Launch hidden
    subprocess.Popen([emb_path], creationflags=CREATE_NO_WINDOW)

    # 4) Wait up to 5 min for the window title to include the filename
    max_wait, interval = 300, 2
    start = time.time()
    while time.time() - start < max_wait:
        try:
            app = Application(backend="uia").connect(title_re=EMB_WINDOW_TITLE, timeout=1)
            win = app.window(title_re=EMB_WINDOW_TITLE)
            if file_base.lower() in (win.window_text() or "").lower():
                break
        except:
            pass
        time.sleep(interval)
    else:
        return f"Error: timed out waiting for {file_base}.emb window", 504

    # 5) Minimize then queue
    try:
        win.minimize()
    except:
        pass

    try:
        win.type_keys("+!q", set_foreground=False)   # Shift+Alt+Q
        time.sleep(1)                                # give the queue dialog a moment
        win.type_keys("{ENTER}")                     # confirm
        time.sleep(0.5)

        # ⬇ NEW: close just the design window, leave the app running
        win.close()

        return "Queued ✅ and closed design window", 200
    except Exception as e:
        return f"Error during queue/close: {e}", 500


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
import time
from datetime import datetime
import gspread
from plyer import notification  # or your notification library

def connect_google_sheet(sheet_name, tab_name=TAB_NAME):
    # ▶ Explicitly load the JSON key so there’s no ambiguity
    creds_path = r"C:\Users\eckar\Desktop\OrderEntry,Inventory,QR,PrintPDF\Keys\poetic-logic-454717-h2-3dd1bedb673d.json"
    print(f"[GoogleSheets] Loading credentials from: {creds_path}")
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    # use gspread's built-in service_account helper
    client = gspread.service_account(filename=creds_path, scopes=scopes)

    print(f"[GoogleSheets] Opening spreadsheet: {sheet_name!r}")
    ss = client.open(sheet_name)
    worksheet = ss.worksheet(tab_name)
    print(f"[GoogleSheets] Successfully connected to tab: {tab_name!r}")
    return worksheet, ss

def update_sheet(sheet, data):
    if not data:
        return
    order_key = clean_value(data[0][0]).strip().lower().replace('.0', '')
    headers = sheet.row_values(1)
    hmap = {h.strip().lower(): i for i, h in enumerate(headers)}
    if 'order number' not in hmap:
        return

    # delete any existing rows for this order
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

    # Draw product name, wrapped into multiple lines but bound to box height
    product = get_product(order_number, spreadsheet)

    pdf.set_font('Arial', 'B', 9)
    pdf.set_text_color(0, 0, 0)

    # Position at top‐left of the box
    pdf.set_xy(content_x, ys[4])

    # Compute a good line height based on your font size
    line_h = pdf.font_size * 1.01

    # Draw wrapped text inside the box with proper line spacing
    pdf.multi_cell(
        w      = inner_w,
        h      = line_h,
        txt    = product,
        border = 0,
        align  = 'C',
        fill   = False
    )

    pdf.set_font('Arial','B',10)
    pdf.set_xy(content_x, ys[5])
    pdf.cell(inner_w, rows[5], f"Ship: {get_due_date(order_number, spreadsheet)}", 0, 2, 'C')


def shrink_page_and_stamp_horizontal_qrs(original_pdf_path, order_number, spreadsheet, margin_pts=72):
    qty = get_order_quantity(order_number, spreadsheet)
    company = get_company_name(order_number, spreadsheet)
    try:
        with open(original_pdf_path, 'rb') as f:
            raw = f.read()
        reader = PdfFileReader(BytesIO(raw))
        writer = PdfFileWriter()

        page0 = reader.getPage(0)
        w, h = float(page0.mediaBox.getWidth()), float(page0.mediaBox.getHeight())

        pad = 10
        size = 54
        label_h = 10
        box_w = 60
        box_h = 120

        left_margin = pad
        right_margin = box_w + pad
        top_margin = pad + label_h + int(size * 0.75)
        bottom_margin = margin_pts

        max_w = w - left_margin - right_margin
        max_h = h - top_margin - bottom_margin
        scale = min(0.90, max_w / w, max_h / h)

        tx = left_margin
        ty = bottom_margin
        page0.addTransformation([scale, 0, 0, scale, tx, ty])

        writer.addPage(page0)
        for i in range(1, reader.getNumPages()):
            writer.addPage(reader.getPage(i))

        pdf = FPDF(unit='pt', format=(w, h))
        pdf.set_auto_page_break(False)
        pdf.add_page()
        pdf.set_font('Arial','B',8)

        top_size = int(size * 0.75)
        box_w = 60
        left_bound = pad
        right_bound = w - (pad + box_w)
        available = right_bound - left_bound
        gap = (available - (5 * top_size)) / 6

        y_lbl = pad
        y_img = pad + label_h

        for idx in range(5):
            x = left_bound + gap * (idx + 1) + top_size * idx

            if idx == 0:
                url = shorten_url(EMB_START_URL + urllib.parse.quote(order_number))
                qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_L, box_size=6, border=0)
                qr.add_data(url + '\r'); qr.make(fit=True)
                img = qr.make_image().convert('1', dither=Image.NONE)
                tmp = f"qr_top_{idx}.png"; img.save(tmp)

                pdf.set_xy(x, y_lbl)
                pdf.cell(top_size, label_h, "Machine Start", 0, 2, 'C')
                pdf.image(tmp, x, y_img, w=top_size, h=top_size)
                os.remove(tmp)

            elif idx == 1:
                queue_url = f"http://192.168.1.24:5001/queue?file={urllib.parse.quote(order_number)}"
                qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_L, box_size=6, border=0)
                qr.add_data(queue_url + '\r'); qr.make(fit=True)
                img = qr.make_image().convert('1', dither=Image.NONE)
                tmp = f"qr_top_{idx}.png"; img.save(tmp)

                pdf.set_xy(x, y_lbl)
                pdf.cell(top_size, label_h, "Queue Design", 0, 2, 'C')
                pdf.image(tmp, x, y_img, w=top_size, h=top_size)
                os.remove(tmp)

            else:
                pdf.set_draw_color(200,200,200)
                pdf.rect(x, y_img, top_size, top_size)

        draw_top_right_box(pdf, order_number, spreadsheet, w, h, scale)

        x0 = w - box_w - pad
        y_start = pad + box_h + pad
        labels = ['Fur List','Cut List','Print List','Embroidery List','Shipping']
        for idx, label in enumerate(labels):
            if label == 'Shipping':
                raw = f"https://machineschedule.netlify.app/ship?company={urllib.parse.quote(company)}&order={urllib.parse.quote(order_number)}"
            else:
                raw = f"{label.replace(' ', '_')}_{order_number}_{qty}"
                raw = shorten_url(DATA_URL + urllib.parse.quote(raw))

            qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_L, box_size=6, border=0)
            qr.add_data(raw + '\r'); qr.make(fit=True)
            img = qr.make_image().convert('1', dither=Image.NONE)
            tmp = f"qr_bot_{idx}.png"; img.save(tmp)

            y_lab = y_start + idx * (label_h + size + pad + 5)
            y_img = y_lab + label_h
            pdf.set_xy(x0, y_lab)
            pdf.cell(size, label_h, label, 0, 2, 'C')
            pdf.image(tmp, x0, y_img, w=size, h=size)
            os.remove(tmp)

        note_pad = 10
        note_h = bottom_margin - 2 * note_pad
        note_y = h - bottom_margin + note_pad
        note_x = pad
        note_w = w - 2 * pad

        pdf.set_fill_color(255,255,255)
        pdf.rect(note_x, note_y, note_w, note_h, style='F')
        pdf.set_draw_color(0,0,0)
        pdf.set_line_width(1)
        pdf.rect(note_x, note_y, note_w, note_h, style='D')

        note_text = get_notes(order_number, spreadsheet)
        pdf.set_text_color(0,0,0)
        pdf.set_font('Arial','',10)
        pdf.set_xy(note_x + 5, note_y + 5)
        pdf.multi_cell(note_w - 10, 12, f"Notes: {note_text or '-'}")

        overlay_stream = BytesIO(pdf.output(dest='S').encode('latin-1'))
        overlay_reader = PdfFileReader(overlay_stream)
        overlay_page = overlay_reader.getPage(0)
        target = writer.getPage(0)
        target.merge_page(overlay_page)

        os.makedirs(OUTPUT_FOLDER, exist_ok=True)
        outp = os.path.join(OUTPUT_FOLDER, os.path.splitext(os.path.basename(original_pdf_path))[0] + '_Stamped.pdf')
        with open(outp, 'wb') as of:
            writer.write(of)

        os.remove(original_pdf_path)

    except Exception as e:
        print('stamping error:', e)





# ================================
# SECTION 8: File Monitoring
# ================================
processed = set()

class PDFHandler(FileSystemEventHandler):
    def on_created(self, event):
        # Only handle newly created PDF files
        if event.is_directory or not event.src_path.lower().endswith('.pdf'):
            return
        p = event.src_path
        # Avoid reprocessing the same file
        if p in processed:
            return
        # Wait until the file is fully written
        if not wait_for_stable_file(p):
            return
        processed.add(p)

        # Extract order number and thread usage
        order = clean_value(os.path.splitext(os.path.basename(p))[0])
        usage = extract_thread_usage(p)
        if not usage:
            return

        # Update Google Sheet
        rows = [
            [order,
             r[1],
             r[2],
             r[3] * get_order_quantity(order, sheet_spreadsheet),
             r[4]]
            for r in usage
        ]
        update_sheet(sheet_thread, rows)

        # Stamp PDF and add QR codes
        shrink_page_and_stamp_horizontal_qrs(p, order, sheet_spreadsheet)


def monitor_folder(folder, sh, ss):
    global sheet_thread, sheet_spreadsheet
    sheet_thread, sheet_spreadsheet = sh, ss
    os.makedirs(folder, exist_ok=True)

    obs = Observer()
    obs.schedule(PDFHandler(), folder, recursive=False)
    obs.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        obs.stop()
        obs.join()


# ================================
# SECTION 9: Scanner Listener (Hybrid QR actions, no notifications)
# ================================
import time
import os
import subprocess
import urllib.parse
import keyboard
from pywinauto import Application
from pywinauto.keyboard import send_keys
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

# scanner timing threshold (seconds)
TIME_THRESHOLD = 0.08
_accumulated = ""
_last_time   = 0.0

# Embedded file folder and window title
EMB_FOLDER       = r"C:\Users\eckar\Desktop\EMB"
EMB_WINDOW_TITLE = r"EmbroideryStudio\s*2025.*"
CREATE_NO_WINDOW = 0x08000000

# Function for non-queue QR: open URL and flash success via browser
from selenium import webdriver
import time
import webbrowser
import win32gui
import win32con

def bring_chrome_to_front():
    def enum_handler(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if "Chrome" in title:
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(hwnd)
                return False  # Stop enumeration after first match
        return True

    win32gui.EnumWindows(enum_handler, None)

def open_and_handle_url(url):
    # If it's the Shipping QR, use browser with focus
    if url.startswith("https://machineschedule.netlify.app/ship"):
        print(f"[QR Scanner] Opening Shipping link: {url}")
        webbrowser.open_new_tab(url)

        # Wait a moment for browser tab to open
        time.sleep(1)

        # Try to bring Chrome window to front
        for window in gw.getWindowsWithTitle('Chrome'):
            try:
                window.activate()
                window.maximize()
                break
            except Exception as e:
                print("Couldn't activate Chrome window:", e)

        # Move mouse and click to ensure browser has focus (optional)
        pyautogui.moveTo(100, 100)
        pyautogui.click()
        return

    # Handle all other QR URLs the normal way (with ChromeDriver)
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--no-first-run")
    options.add_argument("--password-store=basic")

    driver = webdriver.Chrome(options=options)

    # Try to avoid opening duplicates
    for handle in driver.window_handles:
        driver.switch_to.window(handle)
        if url in driver.current_url:
            return

    driver.get(url)
    driver.execute_script("window.focus();")

    start = time.time()
    success = False
    while time.time() - start < 5:
        if "Success" in driver.page_source:
            success = True
            break
        time.sleep(0.5)

    if success:
        for _ in range(4):
            driver.execute_script("document.body.style.backgroundColor='lightgreen';")
            time.sleep(0.3)
            driver.execute_script("document.body.style.backgroundColor='white';")
            time.sleep(0.3)
        driver.quit()
    else:
        while True:
            driver.execute_script("document.body.style.backgroundColor='red';")
            time.sleep(0.3)
            driver.execute_script("document.body.style.backgroundColor='white';")
            time.sleep(0.3)



# Function to queue the EMB file

def queue_emb_file(file_base):
    emb_path = os.path.join(EMB_FOLDER, f"{file_base}.EMB")
    print(f"[Queue Service] Looking for: {emb_path}")
    if not os.path.exists(emb_path):
        print(f"[Queue Service] File not found: {file_base}.EMB")
        return
    print(f"[Queue Service] Launching: {emb_path}")
    cmd = f'start "" /min "{emb_path}"'
    subprocess.Popen(cmd, shell=True, creationflags=CREATE_NO_WINDOW)

    # wait for EmbroideryStudio title includes file_base
    start = time.time()
    win_handle = None
    while time.time() - start < 60:
        try:
            app = Application(backend="uia").connect(title_re=EMB_WINDOW_TITLE, timeout=1)
            win_handle = app.window(title_re=EMB_WINDOW_TITLE)
            title = win_handle.window_text() or ""
            print(f"DEBUG ▶ Window title: {title}")
            if file_base.lower() in title.lower():
                print("[Queue Service] Window loaded")
                break
        except Exception:
            pass
        time.sleep(1)
    else:
        print(f"[Queue Service] Window load timeout for {file_base}")
        return
    # focus and send SHIFT+ALT+Q, then ENTER, then minimize
    try:
        win_handle.set_focus()
        time.sleep(0.5)
        send_keys('+%q')
        time.sleep(2)
        send_keys('{ENTER}')
        win_handle.minimize()
        print(f"[Queue Service] Queued {file_base}")
    except Exception as e:
        print(f"[Queue Service] Error sending keys: {e}")

# Keyboard hook: route scan based on QR content

def on_key_event(event):
    global _accumulated, _last_time
    if event.event_type != 'down':
        return
    now = time.time()
    if now - _last_time > TIME_THRESHOLD:
        _accumulated = ''
    _last_time = now

    key = event.name
    if len(key) == 1:
        _accumulated += key
        return
    if key.lower() in ('enter','return'):
        candidate = _accumulated.strip()
        print(f"DEBUG ▶ Full scan buffer: {candidate!r}")

        # only proceed if this looks like a scanner payload:
        # either a URL or a short order-code (e.g. ph-12345) of at least 5 chars
        if candidate.lower().startswith(('http://','https://')):
            parsed = urllib.parse.urlparse(candidate)
            params = urllib.parse.parse_qs(parsed.query)
            if parsed.path.endswith('/queue') and 'file' in params:
                queue_emb_file(params['file'][0])
            else:
                open_and_handle_url(candidate)
        elif re.match(r'^[A-Za-z0-9\-_]{5,}$', candidate):
            # looks like a job code, not random typing
            queue_emb_file(candidate)
        _accumulated = ''
    else:
        pass

print('DEBUG ▶ Registering keyboard hook')
keyboard.hook(on_key_event)

# ================================
# SECTION 10: Main
# ================================
import logging
import time
import threading

if __name__ == "__main__":
    # 1) Connect to Google Sheets & ensure output folder exists
    sheet_thread, sheet_spreadsheet = connect_google_sheet(SHEET_NAME, TAB_NAME)
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    # 2) Silence Flask/Werkzeug dev-server warning
    log = logging.getLogger('werkzeug')
    log.setLevel(logging.ERROR)

    # 3) Start embedded Flask queue-service on port 5001
    flask_thread = threading.Thread(
        target=lambda: app.run(
            host="0.0.0.0",
            port=5001,
            debug=False,
            use_reloader=False
        ),
        daemon=True
    )
    flask_thread.start()

    # 4) Begin watching the folder for new PDFs
    print(f"Monitoring folder: {MONITOR_FOLDER}")
    folder_thread = threading.Thread(
        target=monitor_folder,
        args=(MONITOR_FOLDER, sheet_thread, sheet_spreadsheet),
        daemon=True
    )
    folder_thread.start()

    # 5) Start clipboard listener to catch scan-as-paste events
    def clipboard_listener():
        last = ""
        while True:
            try:
                clip = pyperclip.paste().strip()
                if clip and clip != last and clip.lower().startswith("http://192.168.1.24:5001/queue?file="):
                    last = clip
                    print("Clipboard scan detected URL:", clip)
                    open_and_handle_url(clip)
                    pyperclip.copy("")  # clear to avoid repeats
            except Exception as e:
                print("Clipboard listener error:", e)
            time.sleep(0.2)

    threading.Thread(target=clipboard_listener, daemon=True).start()

    # 6) Block forever so threads stay alive
    print("Service running. Press Ctrl+C to exit.")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("Shutting down…")
