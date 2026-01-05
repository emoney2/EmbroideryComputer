import os
import glob
import subprocess
import time
import threading

from flask import Flask, request, jsonify
from flask_cors import CORS

# ----------------------------
# LABEL WATCHER IMPORTS
# ----------------------------
import win32print
import win32api
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# ============================================================
#                EXISTING PRINT SERVER SETTINGS
# ============================================================

BASE_PATH = r"G:\My Drive\Orders"
SUMATRA_PATH = r"C:\Users\eckar\AppData\Local\SumatraPDF\SumatraPDF.exe"

app = Flask(__name__)

CORS(app, resources={
    r"/*": {
        "origins": [
            "http://localhost:3000",
            "http://127.0.0.1:3000",
            "https://machineschedule.netlify.app"
        ],
        "methods": ["GET", "POST", "OPTIONS"],
        "allow_headers": ["Content-Type"]
    }
})

def find_latest(order, suffix):
    pattern = os.path.join(BASE_PATH, order, f"{order}_{suffix}*.pdf")
    matches = glob.glob(pattern)
    if not matches:
        return None
    return max(matches, key=os.path.getmtime)


def silent_print(filepath):
    subprocess.run([
        SUMATRA_PATH,
        "-print-to-default",
        "-silent",
        filepath
    ], check=False)

@app.route("/print", methods=["OPTIONS"])
def print_options():
    return jsonify({"ok": True, "message": "Print service online"}), 200


@app.post("/print")
def print_files():
    data = request.get_json(silent=True) or {}
    order = str(data.get("order"))
    mode = data.get("mode")

    printed = []

    if mode in ("both", "process"):
        stamped = find_latest(order, "Stamped")
        if stamped:
            silent_print(stamped)
            printed.append(stamped)

    if mode in ("both", "binsheet"):
        binsheet = find_latest(order, "BINSHEET")
        if binsheet:
            silent_print(binsheet)
            printed.append(binsheet)

    return jsonify({
        "ok": True,
        "order": order,
        "mode": mode,
        "printed": printed
    })


@app.post("/queue-emb")
def queue_emb():
    from pywinauto.application import Application
    from pywinauto.keyboard import send_keys
    import time
    import subprocess
    import os

    data = request.get_json(silent=True) or {}
    emb_path = data.get("path")

    if not emb_path or not os.path.exists(emb_path):
        return jsonify({"error": f"EMB file not found: {emb_path}"}), 400

    subprocess.Popen(['start', '', emb_path], shell=True)
    time.sleep(2.0)

    try:
        app_w = Application(backend="uia").connect(title_re=".*EmbroideryStudio.*")
        win = app_w.window(title_re=".*EmbroideryStudio.*")
    except Exception as e:
        return jsonify({"error": f"Could not connect to Wilcom window: {e}"}), 500

    win.set_focus()
    send_keys('+%q')
    time.sleep(1.0)
    send_keys('^%{F4}')

    return jsonify({"ok": True})

# ============================================================
#              UPS LABEL WATCHER (ADDED HERE)
# ============================================================

WATCH_FOLDER = r"G:\My Drive\Label Printer"
PRINTER_NAME = "PL80E"

from PIL import Image, ImageWin, ImageChops
import fitz
import win32print
import win32ui
import win32con
import os

def crop_whitespace(img):
    bg = Image.new(img.mode, img.size, img.getpixel((0,0)))
    diff = ImageChops.difference(img, bg)
    bbox = diff.getbbox()
    return img.crop(bbox) if bbox else img

def print_label(file_path):
    try:
        print("\n==============================")
        print("📄 START CROPPED PRINT (FINAL MODE)")
        print("==============================")
        print(f"Incoming file: {file_path}")

        folder = os.path.dirname(file_path)
        base = os.path.splitext(os.path.basename(file_path))[0]

        # STEP 1: Convert PDF → PNG
        if file_path.lower().endswith(".pdf"):
            pdf = fitz.open(file_path)
            page = pdf[0]
            pix = page.get_pixmap(matrix=fitz.Matrix(4, 4))
            png_path = os.path.join(folder, f"{base}.png")
            pix.save(png_path)
            file_path = png_path
            print("Converted PDF → PNG:", png_path)

        # Load the image
        img = Image.open(file_path).convert("RGB")
        print("Original size:", img.size)

        # STEP 2: Crop whitespace
        cropped = crop_whitespace(img)
        print("Cropped size:", cropped.size)

        # STEP 3: Rotate to portrait
        rotated = cropped.rotate(90, expand=True)
        print("Rotated size:", rotated.size)

        # STEP 4: Resize to printer drawable area (minus 1/8” all around)
        dc = win32ui.CreateDC()
        dc.CreatePrinterDC(PRINTER_NAME)
        PW = dc.GetDeviceCaps(win32con.HORZRES)     # e.g., 799
        PH = dc.GetDeviceCaps(win32con.VERTRES)     # e.g., 1199

        print(f"Printer drawable area: {PW} × {PH}")

        # 1/8 inch border = 25px each side → subtract 50px width & height
        BORDER = 25
        SAFE_W = PW - (BORDER * 2)
        SAFE_H = PH - (BORDER * 2)

        print(f"Shrinking for 1/8 inch margins → final size: {SAFE_W} × {SAFE_H}")

        final_img = rotated.resize((SAFE_W, SAFE_H), Image.LANCZOS)

        final_path = os.path.join(folder, f"{base}_final.png")
        final_img.save(final_path)
        print("Saved final print image:", final_path)

        # STEP 5: PRINT (centered on page)
        hDC = win32ui.CreateDC()
        hDC.CreatePrinterDC(PRINTER_NAME)

        hDC.StartDoc("UPS Label - Final Shrink")
        hDC.StartPage()

        dib = ImageWin.Dib(final_img)

        # Center the image inside PW × PH
        left = BORDER
        top = BORDER
        right = SAFE_W + BORDER
        bottom = SAFE_H + BORDER

        dib.draw(hDC.GetHandleOutput(), (left, top, right, bottom))

        hDC.EndPage()
        hDC.EndDoc()

        print("✅ PRINT COMPLETE — Sent final shrunk label to printer")

        # STEP 6: CLEAN OUT THE ENTIRE FOLDER
        print("🧹 Deleting all files in Label Printer folder...")

        for fname in os.listdir(folder):
            try:
                os.remove(os.path.join(folder, fname))
            except Exception:
                pass

        print("🧹 Folder cleaned!")

        print("======================================\n")

    except Exception as e:
        print(f"❌ Error:", e)


import time
import os

last_print_time = {}

class LabelHandler(FileSystemEventHandler):
    def on_created(self, event):
        if event.is_directory:
            return

        file_path = event.src_path

        # Ignore temp/sync files
        if not file_path.lower().endswith(".pdf"):
            return
        if ".tmp" in file_path.lower() or "~" in file_path.lower():
            return

        # Debounce (4 seconds)
        now = time.time()
        last = last_print_time.get(file_path, 0)
        if now - last < 4:
            return
        last_print_time[file_path] = now

        # Make sure file is fully written
        time.sleep(2)

        print_label(file_path)



def start_label_watcher():
    print(f"👀 Watching UPS Label folder: {WATCH_FOLDER}")
    print(f"🖨️ UPS Label Printer: {PRINTER_NAME}")

    event_handler = LabelHandler()
    observer = Observer()
    observer.schedule(event_handler, WATCH_FOLDER, recursive=False)
    observer.start()

    while True:
        time.sleep(1)

# ============================================================
#                     COMBINED STARTUP
# ============================================================

if __name__ == "__main__":

    # Start label watcher in background thread
    watcher_thread = threading.Thread(target=start_label_watcher, daemon=True)
    watcher_thread.start()

    print("🚀 JRCO Print Server + UPS Label Watcher running...")
    app.run(host="127.0.0.1", port=5009)
