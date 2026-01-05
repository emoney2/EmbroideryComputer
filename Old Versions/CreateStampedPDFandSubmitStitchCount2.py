# ================================
# CreateStampedPDFandSubmitStitchCount2.py (Complete)
# ================================

# ================================
# SECTION 1: Imports
# ================================
import os
import time
import re
import urllib.parse
from io import BytesIO
from datetime import datetime

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

# ================================
# SECTION 2: Configuration & Constants
# ================================
CREDENTIALS_PATH  = r"C:\Users\eckar\Desktop\OrderEntry,Inventory,QR,PrintPDF\credentials.json"
MONITOR_FOLDER    = r"C:\Users\eckar\Desktop\Embroidery Sheets"
OUTPUT_FOLDER     = r"C:\Users\eckar\Desktop\Embroidery Sheets\PrintPDF"
SHORTENER_SERVICE = pyshorteners.Shortener()

SHEET_NAME        = "JR and Co."
TAB_NAME          = "Thread Data"
WEBAPP_URL        = "https://script.google.com/macros/s/11s5QahOgGsDRFWFX6diXvonG5pESRE1ak79V-8uEbb4/exec"
EMB_START_URL     = f"{WEBAPP_URL}?event=machine_start&order="  # Machine Start QR endpoint
DATA_URL          = f"{WEBAPP_URL}?data="                      # Generic data QR endpoint

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
            stable = 0
            last = now
        time.sleep(0.5)
    return False


def color_to_rgb(color):
    c = re.sub(r"\s*fur\s*$", '', color.lower()).strip()
    try:
        frac = mcolors.to_rgb(c.replace(' ', ''))
        return tuple(int(255*comp) for comp in frac)
    except Exception:
        return (255,255,255)


def get_contrast_color(color):
    r,g,b = color_to_rgb(color)
    lum = (r*299 + g*587 + b*114)/1000
    return (0,0,0) if lum>127 else (255,255,255)

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
    order_key = clean_value(data[0][0]).strip().lower().replace('.0','')
    headers = sheet.row_values(1)
    hmap = {h.strip().lower():i for i,h in enumerate(headers)}
    if 'order number' not in hmap:
        return
    col = hmap['order number']+1
    vals = sheet.col_values(col)
    to_delete = [i for i,v in enumerate(vals[1:], start=2)
                 if str(v).strip().lower().replace('.0','')==order_key]
    reqs=[]
    for r in sorted(to_delete, reverse=True):
        reqs.append({'deleteDimension':{
            'range':{'sheetId':sheet.id,'dimension':'ROWS','startIndex':r-1,'endIndex':r}
        }})
    replaced = bool(reqs)
    if reqs:
        try:
            sheet.spreadsheet.batch_update({'requests':reqs})
            time.sleep(1)
        except Exception as e:
            print('Batch delete error:', e)
    ts = datetime.now().strftime('%m/%d/%Y %H:%M:%S')
    for row in data:
        new = ['']*len(headers)
        if 'date' in hmap:         new[hmap['date']] = ts
        if 'order number' in hmap: new[hmap['order number']] = str(row[0]).strip()
        if 'color' in hmap:        new[hmap['color']] = row[1]
        if 'color name' in hmap:   new[hmap['color name']] = row[2]
        if 'length (ft)' in hmap:  new[hmap['length (ft)']] = row[3]
        if 'stitch count' in hmap: new[hmap['stitch count']] = row[4]
        if 'in/out' in hmap:       new[hmap['in/out']] = 'OUT'
        if 'o/r' in hmap:          new[hmap['o/r']] = ''
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
    key = order_number.strip().lower()
    try:
        ws = spreadsheet.worksheet('Production Orders')
        vals = ws.get_all_values()
        hdr = vals[0]
        idx = next((i for i,h in enumerate(hdr) if h.strip().lower()=='quantity'), None)
        if idx is None: return 1.0
        for row in vals[1:]:
            if row and row[0].strip().lower()==key:
                return float(row[idx])
    except Exception as e:
        print('get_order_quantity error:', e)
    return 1.0

@lru_cache()
def get_fur_color(order_number, spreadsheet):
    key = order_number.strip().lower()
    try:
        ws = spreadsheet.worksheet('Production Orders')
        vals = ws.get_all_values()
        hdr = vals[0]
        idx = next((i for i,h in enumerate(hdr) if 'fur color' in h.strip().lower()), None)
        if idx is None: return ''
        for row in vals[1:]:
            if row and row[0].strip().lower()==key:
                return row[idx]
    except Exception as e:
        print('get_fur_color error:', e)
    return ''

@lru_cache()
def get_due_date(order_number, spreadsheet):
    key = order_number.strip().lower()
    try:
        ws = spreadsheet.worksheet('Production Orders')
        vals = ws.get_all_values()
        hdr = vals[0]
        idx = next((i for i,h in enumerate(hdr) if 'ship date' in h.strip().lower()), None)
        if idx is None: return ''
        for row in vals[1:]:
            if row and row[0].strip().lower()==key:
                return row[idx]
    except Exception as e:
        print('get_due_date error:', e)
    return ''

@lru_cache()
def get_product(order_number, spreadsheet):
    key = order_number.strip().lower()
    try:
        ws = spreadsheet.worksheet('Production Orders')
        vals = ws.get_all_values()
        hdr = vals[0]
        idx = next((i for i,h in enumerate(hdr) if h.strip().lower()=='product'), None)
        if idx is None: return ''
        for row in vals[1:]:
            if row and row[0].strip().lower()==key:
                return row[idx]
    except Exception as e:
        print('get_product error:', e)
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
            stitch_count = float(m.group(1).replace(",","")) if m else 0.0
            lines = text.splitlines()
            hdr = next((i for i,l in enumerate(lines) if "N# Color Name Length" in l), None)
            if hdr is None:
                return data
            for line in lines[hdr+1:]:
                if not line or not line[0].isdigit():
                    continue
                m2 = re.search(r"(\d+)\.\s+(\d+)\s+([\w\s]+?)\s+([\d\.]+)ft", line)
                if m2:
                    data.append([
                        clean_value(m2.group(1)), clean_value(m2.group(2)),
                        clean_value(m2.group(3).strip()), float(m2.group(4)), stitch_count
                    ])
    except Exception as e:
        print("extract_thread_usage error:", e)
    return data

# ================================
# SECTION 7: PDF Stamping & QR Codes
# ================================
def draw_top_right_box(pdf, order_number, spreadsheet, w, h, s):
    # (same implementation as before)
    box_w, box_h = 60, 120
    x0, y0 = w - box_w - 10, 10
    pdf.set_line_width(1)
    pdf.set_draw_color(0,0,0)
    pdf.rect(x0, y0, box_w, box_h)
    content_x = x0 + 2
    inner_w = box_w - 4
    rows = [20,20,20,14,24,22]
    offs = [5,6,6,3,6,6]
    starts = [y0]
    for r in rows[:-1]:
        starts.append(starts[-1] + r)
    ys = [starts[i] + offs[i] for i in range(len(rows))]
    pdf.set_font("Arial","BU",10)
    pdf.set_text_color(0,0,0)
    pdf.set_xy(content_x, ys[0]); pdf.cell(inner_w, rows[0], "Sewing", 0,2,'C')
    pdf.set_font("Arial","B",9)
    pdf.set_xy(content_x, ys[1]); pdf.cell(inner_w, rows[1], f"QTY: {int(get_order_quantity(order_number, spreadsheet))}", 0,2,'C')
    fur = re.sub(r"\s*fur\s*$","", get_fur_color(order_number, spreadsheet), flags=re.IGNORECASE).strip()
    rgb = color_to_rgb(fur)
    contrast = get_contrast_color(fur)
    pdf.set_font("Arial","",8)
    pdf.set_fill_color(*rgb); pdf.rect(content_x, starts[2], inner_w, rows[2], style='F')
    pdf.set_text_color(*contrast); pdf.set_xy(content_x, ys[2]); pdf.cell(inner_w, rows[2], fur, 0,2,'C')
    pdf.set_fill_color(*rgb); pdf.rect(content_x, starts[3], inner_w, rows[3], style='F')
    pdf.set_text_color(*contrast); pdf.set_xy(content_x, ys[3]); pdf.cell(inner_w, rows[3], "Fur", 0,2,'C')
    pdf.set_font("Arial","B",12); pdf.set_text_color(0,0,0)
    pdf.set_xy(content_x, ys[4]); pdf.cell(inner_w, rows[4], get_product(order_number, spreadsheet), 0,2,'C')
    pdf.set_font("Arial","B",10)
    pdf.set_xy(content_x, ys[5]); pdf.cell(inner_w, rows[5], f"Ship: {get_due_date(order_number, spreadsheet)}", 0,2,'C')


def shrink_page_and_stamp_horizontal_qrs(original_pdf_path, order_number, spreadsheet, margin_pts=72):
    qty = get_order_quantity(order_number, spreadsheet)
    try:
        with open(original_pdf_path,'rb') as f:
            reader = PdfFileReader(f)
            writer = PdfFileWriter()
            page0 = reader.getPage(0)
            w = float(page0.mediaBox.getWidth()); h = float(page0.mediaBox.getHeight())
            size, label_h, pad = 54,10,10
            extra = 2*(label_h+size+pad); half = extra/2
            scale = (h-extra)/h
            page0.addTransformation([scale,0,0,scale,0,half])
            writer.addPage(page0)
            for i in range(1,reader.getNumPages()): writer.addPage(reader.getPage(i))
            pdf = FPDF(unit='pt',format=(w,h)); pdf.add_page(); pdf.set_font('Arial','B',8)
            # Top QR: only Machine Start, others placeholders
            y_lbl, y_img = pad, pad+label_h
            for idx in range(5):
                x = ((w-5*size)/6)*(idx+1) + size*idx - 36
                if idx==0:
                    short = shorten_url(EMB_START_URL + urllib.parse.quote(order_number))
                    qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_L,box_size=6,border=0)
                    qr.add_data(short); qr.make(fit=True)
                    img = qr.make_image().convert('1',dither=Image.NONE)
                    tmp=f"qr_top_{idx}.png"; img.save(tmp)
                    pdf.set_xy(x,y_lbl); pdf.cell(size,label_h,'Machine Start',0,2,'C')
                    pdf.image(tmp,x,y_img,w=size,h=size); os.remove(tmp)
                else:
                    pdf.set_draw_color(200,200,200); pdf.rect(x,y_img,size,size)
            # Bottom QR row
            bottom_labels=['Fur List','Cut List','Print List','Embroidery List','Shipping']
            y_img2 = h - pad - size; y_lbl2 = y_img2 - label_h
            for idx,label in enumerate(bottom_labels):
                x2 = ((w-5*size)/6)*(idx+1) + size*idx
                raw = f"{label.replace(' ','_')}_{order_number}_{qty}"
                short2 = shorten_url(DATA_URL + urllib.parse.quote(raw))
                qr2=qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_L,box_size=6,border=0)
                qr2.add_data(short2); qr2.make(fit=True)
                img2=qr2.make_image().convert('1',dither=Image.NONE)
                tmp2=f"qr_bot_{idx}.png"; img2.save(tmp2)
                pdf.set_xy(x2,y_lbl2); pdf.cell(size,label_h,label,0,2,'C')
                pdf.image(tmp2,x2,y_img2,w=size,h=size); os.remove(tmp2)
            draw_top_right_box(pdf,order_number,spreadsheet,w,h,scale)
            overlay=PdfFileReader(BytesIO(pdf.output(dest='S').encode('latin-1')))
            pg=writer.getPage(0); pg.merge_page(overlay.getPage(0))
            os.makedirs(OUTPUT_FOLDER,exist_ok=True)
            outp=os.path.join(OUTPUT_FOLDER,os.path.splitext(os.path.basename(original_pdf_path))[0]+'_Stamped.pdf')
            with open(outp,'wb') as of: writer.write(of)
        os.remove(original_pdf_path)
    except Exception as e:
        print('stamping error:',e)

# ================================
# SECTION 8: File Monitoring & Main
# ================================
processed=set()
class PDFHandler(FileSystemEventHandler):
    def on_created(self,event):
        if event.is_directory or not event.src_path.lower().endswith('.pdf'): return
        p=event.src_path
        if p in processed: return
        if not wait_for_stable_file(p): return
        processed.add(p)
        order=clean_value(os.path.splitext(os.path.basename(p))[0])
        usage=extract_thread_usage(p)
        if not usage: return
        rows=[[order,r[1],r[2],r[3]*get_order_quantity(order,sheet_spreadsheet),r[4]] for r in usage]
        update_sheet(sheet_thread,rows)
        shrink_page_and_stamp_horizontal_qrs(p,order,sheet_spreadsheet)

def monitor_folder(folder,sh,ss):
    global sheet_thread,sheet_spreadsheet
    sheet_thread,sheet_spreadsheet=sh,ss
    os.makedirs(folder,exist_ok=True)
    obs=Observer(); obs.schedule(PDFHandler(),folder,recursive=False); obs.start()
    print(f"Monitoring folder: {folder}")
    try:
        while True: time.sleep(1)
    except KeyboardInterrupt: obs.stop()
    obs.join()

if __name__=='__main__':
    sheet_thread,sheet_spreadsheet=connect_google_sheet(SHEET_NAME,TAB_NAME)
    os.makedirs(OUTPUT_FOLDER,exist_ok=True)
    monitor_folder(MONITOR_FOLDER,sheet_thread,sheet_spreadsheet)
