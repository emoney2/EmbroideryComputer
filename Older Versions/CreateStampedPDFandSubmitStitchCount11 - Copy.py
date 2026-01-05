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
import shutil
import subprocess
import keyboard
import traceback

import pdfplumber
import gspread
import qrcode
from fpdf import FPDF
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
import win32gui
import pythoncom
import win32com.client
from win32com.client import gencache
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from google.oauth2.service_account import Credentials
from plyer import notification
from PIL import Image, ImageTk, ImageDraw

from PIL import Image, ImageTk, ImageDraw

# Prefer psd-tools for real PSD/PSB previews
_PSD_OK = False
try:
    from psd_tools import PSDImage  # pip install psd-tools
    _PSD_OK = True
except Exception as _e:
    _PSD_OK = False
    # Optional: print once so you know why PSDs are gray
    print("[Thumbs] psd-tools not available → PSD/PSB previews will use placeholders.", _e)

import pyshorteners
from matplotlib import colors as mcolors
from functools import lru_cache
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
import requests
import pyperclip
import tkinter as tk
from tkinter import simpledialog, messagebox
from controlaltp2 import listen_for_hotkey
import pygetwindow as gw

import win32process
try:
    import psutil
except Exception:
    psutil = None

# Scanner/clipboard gating
SCANNER_ENABLED = True
PDF_BUILD_ACTIVE = False

# Low-level key injection (forces exact Alt+F, G, Enter)
import ctypes, time
from ctypes import wintypes

_user32 = ctypes.windll.user32
KEYEVENTF_KEYUP = 0x0002
VK_MENU   = 0x12   # Alt
VK_F      = 0x46
VK_G      = 0x47
VK_RETURN = 0x0D

# ---- Tk UI dispatcher (run Tk calls on one thread) ----
import queue
import tkinter as tk

_tk_cmds = queue.Queue()

def _tk_thread_main():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    def pump():
        try:
            while True:
                fn = _tk_cmds.get_nowait()
                try:
                    fn(root)
                except Exception:
                    traceback.print_exc()  # see any errors in the Tk thread
                finally:
                    _tk_cmds.task_done()
        except queue.Empty:
            pass
        root.after(50, pump)

    pump()
    root.mainloop()

def run_on_tk_thread(fn):
    """Schedule fn(root) to run on the Tk thread."""
    _tk_cmds.put(fn)

# Start the Tk thread once at startup
threading.Thread(target=_tk_thread_main, daemon=True).start()

# ---- Percent Progress Modal (determinate) ----
_progress_modal = {"win": None, "bar": None, "txt": None}

def _progress_modal_open(root, title="Working…", text="Starting…"):
    import tkinter as tk
    from tkinter import ttk
    win = tk.Toplevel(root)
    win.title(title)
    win.transient(root)
    win.lift()
    win.focus_force()
    win.attributes("-topmost", True)
    win.resizable(False, False)

    frm = ttk.Frame(win); frm.pack(padx=14, pady=12)
    var = tk.StringVar(value=text)
    ttk.Label(frm, textvariable=var).pack(pady=(0,8))
    bar = ttk.Progressbar(frm, mode="determinate", length=300, maximum=100)
    bar.pack()
    bar["value"] = 0
    win.update_idletasks()

    _progress_modal["win"] = win
    _progress_modal["bar"] = bar
    _progress_modal["txt"] = var

def _progress_modal_update(root, percent: float, text: str | None = None):
    # percent: 0..100
    win = _progress_modal["win"]
    bar = _progress_modal["bar"]
    var = _progress_modal["txt"]
    if not win or not bar or not var:
        return
    try:
        bar["value"] = max(0, min(100, percent))
        if text is not None:
            var.set(text)
        win.update_idletasks()
    except Exception:
        pass

def _progress_modal_close(root):
    win = _progress_modal.get("win")
    if not win:
        return
    try:
        win.destroy()
    except Exception:
        pass
    _progress_modal["win"] = None
    _progress_modal["bar"] = None
    _progress_modal["txt"] = None

# ---- Tiny "Loading previews…" modal (for long UI work) ----
def _open_loading_modal(root, total_count:int|None=None, text="Loading previews…"):
    import tkinter as tk
    from tkinter import ttk
    win = tk.Toplevel(root)
    win.title("Please wait")
    win.transient(root)
    win.lift()
    win.focus_force()
    win.attributes("-topmost", True)
    win.resizable(False, False)
    frm = ttk.Frame(win); frm.pack(padx=12, pady=12)
    var = tk.StringVar(value=text if total_count is None else f"{text} 0/{total_count}")
    ttk.Label(frm, textvariable=var).pack(pady=(0,6))
    pb = ttk.Progressbar(frm, mode="indeterminate", length=260)
    pb.pack()
    pb.start(10)
    win.update_idletasks()
    def set_progress(n:int):
        if total_count is not None:
            var.set(f"{text} {n}/{total_count}")
        try:
            win.update()
        except:
            pass
    return win, pb, set_progress

def _close_loading_modal(win, pb=None):
    try:
        if pb: pb.stop()
    except:
        pass
    try:
        win.destroy()
    except:
        pass



def _vk_down(vk): _user32.keybd_event(vk, 0, 0, 0)
def _vk_up(vk):   _user32.keybd_event(vk, 0, KEYEVENTF_KEYUP, 0)

def _chord_alt_f_g_enter():
    # Alt+F (hold Alt, tap F, release Alt)
    _vk_down(VK_MENU)
    _vk_down(VK_F); _vk_up(VK_F)
    _vk_up(VK_MENU)

    time.sleep(1.0)     # let File backstage render

    # G
    _vk_down(VK_G); _vk_up(VK_G)

    time.sleep(0.5)

    # Enter
    _vk_down(VK_RETURN); _vk_up(VK_RETURN)


BASE_PATH = r"G:\My Drive\Orders"
HOTKEY_O  = "ctrl+alt+o"

# scanner timing threshold (seconds)
TIME_THRESHOLD = 0.1
_accumulated   = ""
_last_time     = 0.0
_scan_token    = 0

# --- Hotkey helpers ---
def _run_async(target, *args, **kwargs):
    t = threading.Thread(target=target, args=args, kwargs=kwargs, daemon=True)
    t.start()
    return t

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

WEBAPP_URL      = "https://script.google.com/macros/s/AKfycbyeDUIpCqiicPfmh9hLkPBxxXt9o0aMfezUj8-jCcsrAXay6c2ZyZJgs3IKHmWC8oSdGA/exec"
EMB_START_URL   = f"{WEBAPP_URL}?event=machine_start&order="
EMB_QUEUE_URL   = "http://127.0.0.1:5001/queue?file="
FOLDER_URL      = "https://drive.google.com/drive/search?q="    # search orders by number
EMB_LIST_URL    = f"{WEBAPP_URL}?event=embroidery_list&order="
PRINT_LIST_URL  = f"{WEBAPP_URL}?event=print_list&order="
SHIP_URL        = "https://machineschedule.netlify.app/ship?order="
FUR_URL         = f"{WEBAPP_URL}?event=fur_list&order="
CUT_URL         = f"{WEBAPP_URL}?event=cut_list&order="
THREAD_DATA_URL = f"{WEBAPP_URL}?event=thread_data&order="



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

import base64

@app.route("/success")
def success_screen():
    """
    Light-green success page that shows the order image, order number, and company.
    Example: http://127.0.0.1:5001/success?order=12345
    Auto-closes after ~7 seconds (best effort).
    """
    order = (request.args.get("order") or "").strip()
    if not order:
        order = "Unknown"

    # Find the oldest image in the order’s folder
    job_folder = get_job_folder(order)
    img_path, _ = find_oldest_image(job_folder)

    # Base64 encode the image (optional: if none, we’ll just hide the <img>)
    img_b64 = ""
    if img_path and os.path.exists(img_path):
        try:
            with open(img_path, "rb") as f:
                img_b64 = base64.b64encode(f.read()).decode("ascii")
        except Exception:
            img_b64 = ""

    # Pull metadata
    try:
        qty = int(get_order_quantity(order, sheet_spreadsheet))
    except Exception:
        qty = 0
    company = get_company_name(order, sheet_spreadsheet) or "Unknown"
    design  = get_product(order, sheet_spreadsheet) or ""

    html = f"""
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8" />
  <title>Success</title>
  <style>
    html, body {{
      margin: 0; padding: 0; height: 100%; width: 100%;
      background: #cfeeda; color: #0b3d1f; font-family: Arial, sans-serif;
      display: flex; align-items: center; justify-content: center;
    }}
    .wrap {{
      display: flex; align-items: center; gap: 2rem;
      max-width: 1200px; width: 92%; padding: 2rem;
    }}
    .img-col img {{
      display: block;
      max-width: 40vw; max-height: 70vh;
      border-radius: 10px; box-shadow: 0 10px 30px rgba(0,0,0,0.35);
    }}
    .info-col {{
      display: flex; flex-direction: column; text-align: left;
    }}
    .big {{ font-size: 4rem; font-weight: 800; line-height: 1.1; margin: 0.25em 0; }}
    .mid {{ font-size: 1.75rem; font-weight: 700; opacity: 0.95; margin: 0.25em 0; }}

    /* Stack on very small screens */
    @media (max-width: 800px) {{
      .wrap {{ flex-direction: column; text-align: center; }}
      .info-col {{ text-align: center; }}
      .big {{ font-size: 3rem; }}
      .mid {{ font-size: 1.4rem; }}
      .img-col img {{ max-width: 80vw; max-height: 50vh; }}
    }}
  </style>
</head>
<body>
  <div class="wrap">
    {("<div class='img-col'><img src='data:image/png;base64," + img_b64 + "' alt='Order image' /></div>") if img_b64 else ""}
    <div class="info-col">
      <div class="big">ORDER #{order}</div>
      <div class="big">{company}</div>
      <div class="mid">{design}{(" • Qty: " + str(qty)) if qty else ""}</div>
    </div>
  </div>
  <script>
    // Best-effort auto-close after ~7 sec
    setTimeout(() => {{
      window.close();
      window.location.href = "about:blank";
    }}, 7000);
  </script>
</body>
</html>
"""

    return html

def _fail_page(msg="Update failed"):
    return f"""<!doctype html>
<html><head><meta charset="utf-8"><title>Failed</title>
<style>
  html,body{{margin:0;height:100%;}} body{{display:flex;align-items:center;justify-content:center;background:#000;color:#fff;font-family:Arial}}
  .flash{{width:100%;height:100%;animation:flash 0.5s infinite;display:flex;align-items:center;justify-content:center;font-size:4rem;font-weight:800}}
  @keyframes flash{{0%{{background:#900}}50%{{background:#f00}}100%{{background:#900}}}}
</style></head>
<body><div class="flash">{msg}</div>
<script>setTimeout(()=>{{window.close();window.location.href='about:blank'}}, 7000);</script>
</body></html>"""

@app.route("/fail")
def fail_page():
    from flask import request
    msg = request.args.get("msg", "Update failed")
    return _fail_page(msg)

# ================================
# SECTION 3: Utility Functions
# ================================
def shorten_url(url):
    try:
        return SHORTENER_SERVICE.tinyurl.short(url)
    except Exception:
        return url

# ── START: Ctrl+Alt+O (open .emb) routines ───────────────────────────────────
_busy = False

def prompt_open_order_number():
    """Prompt for the order number to open the .emb file."""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    order = simpledialog.askstring("Open Order File", "Enter Order Number:", parent=root)
    root.destroy()
    return order.strip() if order else None

def open_order_file():
    """Open .emb if it exists, otherwise show an error popup."""
    global _busy
    if _busy:
        return
    _busy = True
    try:
        order = prompt_open_order_number()
        if not order:
            return
        emb_path = os.path.join(BASE_PATH, order, f"{order}.emb")
        if os.path.exists(emb_path):
            os.startfile(emb_path)
        else:
            err = tk.Tk()
            err.withdraw()
            err.attributes('-topmost', True)
            messagebox.showerror(
                "File Not Found",
                f"Cannot find .emb for order {order}\nExpected:\n{emb_path}",
                parent=err
            )
    finally:
        # give the system a moment to clear modifier keys
        time.sleep(0.5)
        _busy = False

def listen_for_ctrl_alt_o():
    """Thread‑target: wait for Ctrl+Alt+O and call open_order_file."""
    while True:
        keyboard.wait(HOTKEY_O)
        open_order_file()


# ── START: Ctrl+Alt+N (Save As PDF) routines ─────────────────────────────────
def prompt_save_order_number():
    """Prompt for the order number when saving a PDF."""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    order = simpledialog.askstring("Enter Order Number", "Order Number:", parent=root)
    root.destroy()
    return order.strip() if order else None

import glob
from pywinauto import Desktop
from pywinauto.keyboard import send_keys

def get_job_folder(order_number: str) -> str:
    """G:\\My Drive\\Orders\\<order#>"""
    return os.path.join(BASE_PATH, order_number)

def find_oldest_image(job_folder: str) -> tuple[str, str] | tuple[None, None]:
    """
    Return (fullpath, basename) for the oldest image in the job folder
    among PNG/JPG/JPEG. If none exist, returns (None, None).
    """
    patterns = ["*.png", "*.jpg", "*.jpeg", "*.PNG", "*.JPG", "*.JPEG"]
    files = []
    for pat in patterns:
        files.extend(glob.glob(os.path.join(job_folder, pat)))
    if not files:
        return None, None
    files.sort(key=lambda p: (os.path.getmtime(p), p))  # oldest first, stable tiebreaker
    oldest = files[0]
    return oldest, os.path.basename(oldest)


def _wait_for_file_dialog(timeout: int = 15):
    """
    Find a common file dialog (UIA or Win32) by heuristics:
    - visible top-level window
    - has at least one Edit control
    - has an 'Open'/'Import'/'Select'/'Choose'/'OK' button
    Returns the window wrapper or None.
    """
    end = time.time() + timeout
    dialog_btn_names = {"open", "import", "select", "choose", "ok"}
    while time.time() < end:
        # Try both backends
        for backend in ("uia", "win32"):
            try:
                desk = Desktop(backend=backend)
                for w in desk.windows():
                    try:
                        if hasattr(w, "is_visible") and not w.is_visible():
                            continue
                        # edits
                        if backend == "uia":
                            edits = w.descendants(control_type="Edit")
                            btns  = w.descendants(control_type="Button")
                        else:
                            edits = w.descendants(class_name="Edit")
                            btns  = w.descendants(class_name="Button")
                        if not edits:
                            continue
                        has_action_btn = any(
                            (b.window_text() or "").strip().lower() in dialog_btn_names
                            for b in btns
                        )
                        if has_action_btn:
                            return w
                    except Exception:
                        continue
            except Exception:
                pass
        time.sleep(0.2)
    return None

def _navigate_dialog_to_folder(dlg, folder_path: str) -> bool:
    """
    In the file dialog, navigate to folder_path via address bar (Ctrl+L).
    Fallback: type into first Edit control.
    """
    try:
        send_keys('^l')  # focus address/location bar
        time.sleep(0.2)
        send_keys(folder_path, with_spaces=True)
        time.sleep(0.1)
        send_keys('{ENTER}')
        return True
    except Exception:
        pass
    try:
        # Fallback to filename edit
        edit = dlg.child_window(control_type="Edit")
    except Exception:
        # Win32 fallback
        try:
            edit = next(e for e in dlg.descendants(class_name="Edit"))
        except Exception:
            return False
    try:
        edit.set_focus()
    except Exception:
        pass
    try:
        try:
            edit.set_edit_text(folder_path)
        except Exception:
            send_keys(folder_path, with_spaces=True)
        send_keys('{ENTER}')
        return True
    except Exception:
        return False

def _type_filename_and_open(dlg, filename: str) -> bool:
    """Type just the filename (not path) and press Enter."""
    if not filename:
        return False
    try:
        edit = dlg.child_window(control_type="Edit")
    except Exception:
        # Win32 fallback
        try:
            edit = next(e for e in dlg.descendants(class_name="Edit"))
        except Exception:
            edit = None
    try:
        if edit:
            edit.set_focus()
            try:
                edit.set_edit_text(filename)
            except Exception:
                send_keys(filename, with_spaces=True)
        else:
            send_keys(filename, with_spaces=True)
        time.sleep(0.1)
        send_keys('{ENTER}')
        return True
    except Exception:
        return False

# === Product Template Insert (after Import Graphic) ===
PRODUCT_TEMPLATE_FOLDER = r"G:\My Drive\Embroidery Templates"  # Blade.emb, Mallet.emb, ...

def try_save_current_doc():
    """Best-effort Ctrl+S on the active EmbroideryStudio window."""
    try:
        win = Desktop(backend="uia").window(title_re=EMB_WINDOW_TITLE)
        win.wait('visible', timeout=10)
        win.set_focus()
        time.sleep(0.1)
        win.type_keys('^s', set_foreground=True)
    except Exception:
        pass

def insert_product_template(order_number: str) -> bool:
    """
    1) Look up Product from Google Sheet (Production Orders -> Product)
    2) If PRODUCT_TEMPLATE_FOLDER/<Product>.emb exists:
       - open it, Ctrl+A/Ctrl+C, close template (don't save), paste into current doc, save.
    3) If no template, notify and just save current doc.
    """
    # 1) Pull Product from your sheet (util you already have)
    try:
        product = (get_product(order_number, sheet_spreadsheet) or "").strip()
    except Exception:
        product = ""

    if not product:
        try:
            notification.notify(title="JR & Co", message="No Product found for this order.", timeout=4)
        except Exception:
            pass
        try_save_current_doc()
        return False

    # 2) Build expected template path like Blade.emb, Mallet.emb, ...
    tpl_path = os.path.join(PRODUCT_TEMPLATE_FOLDER, f"{product}.emb")
    if not os.path.exists(tpl_path):
        try:
            notification.notify(title="JR & Co", message="No Template File is set up.", timeout=4)
        except Exception:
            pass
        try_save_current_doc()
        return False

    desk = Desktop(backend="uia")

    # Remember current ES window (your just-saved doc with imported PNG)
    try:
        main_win = desk.window(title_re=EMB_WINDOW_TITLE)
        main_win.wait('visible', timeout=15)
    except Exception:
        main_win = None

    # Open the template file (new ES window)
    try:
        os.startfile(tpl_path)
    except Exception as e:
        logging.error(f"Failed to open template: {tpl_path} ({e})", exc_info=True)
        try_save_current_doc()
        return False

    # Wait for template window (title contains the template base name)
    base = os.path.splitext(os.path.basename(tpl_path))[0]
    prod_win = None
    start = time.time()
    while time.time() - start < 30:
        try:
            for w in desk.windows():
                title = (w.window_text() or "")
                if re.search(EMB_WINDOW_TITLE, title or "", re.I) and base.lower() in title.lower():
                    prod_win = w
                    break
            if prod_win:
                break
        except Exception:
            pass
        time.sleep(0.3)

    if not prod_win:
        logging.error("Template window not detected in time.")
        try_save_current_doc()
        return False

    # Select all & copy from template
    try:
        prod_win.set_focus(); time.sleep(0.1)
        prod_win.type_keys('^a', set_foreground=True)  # Select All
        time.sleep(0.15)
        prod_win.type_keys('^c', set_foreground=True)  # Copy
        time.sleep(0.15)
    except Exception as e:
        logging.error(f"Copy from template failed: {e}", exc_info=True)
        try_save_current_doc()
        return False

    # 7) Close only the template document (not the whole app)
    try:
        prod_win.set_focus()
        time.sleep(0.1)
        close_active_document_no_save()
    except Exception as e:
        logging.error(f"Close template doc failed: {e}", exc_info=True)


    # Paste into your original document and save
    try:
        if main_win:
            main_win.set_focus()
        else:
            main_win = desk.window(title_re=EMB_WINDOW_TITLE)
            main_win.set_focus()
        time.sleep(0.2)
        main_win.type_keys('^v', set_foreground=True)  # Paste
        time.sleep(0.15)
        main_win.type_keys('^s', set_foreground=True)  # Save
        try:
            notification.notify(title="JR & Co", message=f"Inserted template: {product}", timeout=3)
        except Exception:
            pass
        return True
    except Exception as e:
        logging.error(f"Paste into main doc failed: {e}", exc_info=True)
        try_save_current_doc()
        return False

def close_active_document_no_save():
    """
    Close only the active document in EmbroideryStudio (MDI child), not the whole app.
    Sends Ctrl+F4 and dismisses any save prompt with 'No' / 'Don't Save'.
    """
    desk = Desktop(backend="uia")
    win = desk.window(title_re=EMB_WINDOW_TITLE)
    try:
        win.set_focus()
    except Exception:
        pass

    # Close active document (MDI child)
    try:
        # Ctrl+F4
        win.type_keys('^{F4}', set_foreground=True)
    except Exception:
        send_keys('^{F4}')

    # If a "Save changes?" prompt appears, pick "No" / "Don't Save"
    time.sleep(0.3)
    try:
        dlg = desk.window(title_re=r"(Save|Confirm|Close|Wilcom|EmbroideryStudio)", visible_only=True)
        dlg.wait('visible', timeout=2)
        # Prefer invoking a "No"/"Don't Save" style button
        for pat in [r"(?i)don'?t save", r"(?i)no", r"(?i)discard"]:
            try:
                btn = dlg.child_window(title_re=pat, control_type="Button")
                if btn.exists():
                    btn.invoke()
                    break
            except Exception:
                continue
        else:
            # Fallback: try pressing 'N', then Esc
            send_keys('n')
            time.sleep(0.1)
            send_keys('{ESC}')
    except Exception:
        pass



def _open_import_graphic_via_keystrokes_and_select(
    *, 
    img_full: str | None = None, 
    img_name: str | None = None, 
    png_full: str | None = None, 
    png_name: str | None = None, 
    folder: str
) -> bool:
    # Back-compat with older callers that used png_full/png_name
    if not img_full and png_full:
        img_full = png_full
    if not img_name and png_name:
        img_name = png_name
    if not img_name and img_full:
        img_name = os.path.basename(img_full)

    """
    EXACT KEY CHORD: Alt+F → wait → G → wait → Enter (via low-level VKs),
    then: wait for file dialog → navigate to folder → type the selected image filename → Enter.
    Retries once if the dialog doesn't appear.
    """
    desk = Desktop(backend="uia")
    win = desk.window(title_re=r".*EmbroideryStudio.*")
    win.wait('visible', timeout=15)
    try:
        win.set_focus()
    except Exception:
        pass

    time.sleep(2.0)
    _chord_alt_f_g_enter()

    dlg = _wait_for_file_dialog(timeout=15)
    if not dlg:
        try:
            win.set_focus()
        except Exception:
            pass
        time.sleep(0.6)
        _chord_alt_f_g_enter()
        dlg = _wait_for_file_dialog(timeout=15)
        if not dlg:
            return False

    if not _navigate_dialog_to_folder(dlg, folder):
        return False

    time.sleep(0.4)
    # Type the chosen image filename and open
    return _type_filename_and_open(dlg, img_name)



def save_as_order_pdf():
    """
    Ctrl+Alt+N: Save into the order folder (G:\\My Drive\\Orders\\<order>\\<order>),
    then File→Import Graphic with the oldest PNG in that folder, then insert Product template if available.
    """
    order_number = prompt_save_order_number()
    if not order_number:
        logging.info("No order number provided, aborting.")
        return

    # Build and ensure the correct folder path for saving
    target_no_ext = os.path.join(BASE_PATH, order_number, order_number)
    os.makedirs(os.path.dirname(target_no_ext), exist_ok=True)

    # Save As (as in your current flow)
    keyboard.send('ctrl+s')
    time.sleep(2)
    keyboard.send('alt+n')
    time.sleep(0.1)
    keyboard.write(target_no_ext)
    time.sleep(0.1)
    keyboard.send('enter')

    # Non-blocking toast
    try:
        notification.notify(title="JR & Co", message=f"Saved to {target_no_ext}", timeout=3)
    except Exception:
        pass

    # Compute oldest image (PNG/JPG)
    job_folder = get_job_folder(order_number)
    img_full, img_name = find_oldest_image(job_folder)
    if not img_full:
        try:
            notification.notify(
                title="JR & Co",
                message=f"No images (PNG/JPG) found in {job_folder}. Import skipped.",
                timeout=4
            )
        except Exception:
            pass
        # Even if no image, still try to insert product template
        try:
            insert_product_template(order_number)
        except Exception as e:
            logging.error(f"insert_product_template error: {e}", exc_info=True)
        return

    # Run the exact key sequence and complete the import
    try:
        ok = _open_import_graphic_via_keystrokes_and_select(
            img_full=img_full, img_name=img_name, folder=job_folder
        )

        if not ok:
            messagebox.showwarning(
                "Import Failed",
                "Could not open the Import dialog or select the image automatically.\n"
                "Try again, or open File → Import manually."
            )
            # Still try template insert so the doc is usable
            try:
                insert_product_template(order_number)
            except Exception as e:
                logging.error(f"insert_product_template error: {e}", exc_info=True)
            return
    except Exception as e:
        logging.error(f"Import Graphic error: {e}", exc_info=True)
        messagebox.showwarning(
            "Import Failed",
            "There was an error importing the graphic. Check logs and try again."
        )
        # Still try template insert
        try:
            insert_product_template(order_number)
        except Exception as e2:
            logging.error(f"insert_product_template error: {e2}", exc_info=True)
        return

    # If import succeeded, insert the Product template
    try:
        insert_product_template(order_number)
    except Exception as e:
        logging.error(f"insert_product_template error: {e}", exc_info=True)

def listen_for_ctrl_alt_n():
    """Thread-target: register Ctrl+Alt+N and then sit idle."""
    logging.info("Registering Ctrl+Alt+N hotkey")
    keyboard.add_hotkey('ctrl+alt+n', save_as_order_pdf)
    keyboard.wait()
# ── END: Ctrl+Alt+N (Save As PDF) routines ────────────────────────────────────


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

# --- Write 'Cut Type' to Production Orders for the given order ---
def write_cut_type_to_sheet(order_number: str, cut_type: str) -> bool:
    """
    Writes the selected Cut Type into the 'Production Orders' row for this order.
    - Finds the 'Order' column by header name: tries 'Order #', 'Order Number', 'Order'
    - Finds/uses a 'Cut Type' column (case-insensitive). If not present, fails gracefully.
    Returns True on success, False on any failure.
    """
    try:
        ws = sheet_spreadsheet.worksheet("Production Orders")
        header = ws.row_values(1)
        header_lc = [h.strip().lower() for h in header]

        # Locate order column
        order_candidates = ["order #", "order number", "order"]
        order_col_idx = None
        for name in order_candidates:
            if name in header_lc:
                order_col_idx = header_lc.index(name) + 1
                break
        if not order_col_idx:
            print("[CutType] Could not find 'Order' column")
            return False

        # Locate cut type column
        cut_type_col_idx = None
        for i, h in enumerate(header_lc, start=1):
            if h == "cut type":
                cut_type_col_idx = i
                break
        if not cut_type_col_idx:
            print("[CutType] Could not find 'Cut Type' column")
            return False

        # Find the row for this order
        ord_norm = order_number.lstrip("$").strip()
        col_vals = ws.col_values(order_col_idx)
        target_row = None
        for r in range(2, len(col_vals) + 1):
            v = (col_vals[r-1] or "").strip().lstrip("$")
            if v == ord_norm:
                target_row = r
                break

        if not target_row:
            print(f"[CutType] Order not found in sheet: {order_number}")
            return False

        ws.update_cell(target_row, cut_type_col_idx, cut_type)
        print(f"[CutType] Wrote '{cut_type}' for order {order_number} (row {target_row})")
        return True
    except Exception as e:
        print(f"[CutType] Error writing cut type: {e}")
        return False

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

# SECTION 7: PDF Stamping & QR Codes
from io import BytesIO
import logging

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


# Configure logging at the top of your module
logging.basicConfig(
    level=logging.DEBUG,
    filename='stamping.log',   # send all DEBUG logs into this file
    filemode='w',              # overwrite the file every time you run
    format="%(asctime)s %(levelname)s: %(message)s",
    datefmt="%H:%M:%S"
)

def get_cut_type_for_order(order_number, spreadsheet):
    """Fetch the Cut Type for the given order from the Google Sheet."""
    worksheet = spreadsheet.worksheet("Production Orders")
    headers = worksheet.row_values(1)
    cut_type_col = next((i + 1 for i, h in enumerate(headers) if h.strip().lower() == "cut type"), None)
    order_col    = next((i + 1 for i, h in enumerate(headers) if h.strip().lower() in ["order number", "order #"]), None)
    if not cut_type_col or not order_col:
        return None
    order_numbers = worksheet.col_values(order_col)
    order_row = next((i + 1 for i, num in enumerate(order_numbers) if str(num).strip() == str(order_number)), None)
    if not order_row:
        return None
    return worksheet.cell(order_row, cut_type_col).value

def shrink_page_and_stamp_horizontal_qrs(original_pdf_path, order_number, spreadsheet, margin_pts=72):
    logging.debug("Opening original PDF: %s", original_pdf_path)

    # --- define 'base' early (filename without extension) ---
    base = os.path.splitext(os.path.basename(original_pdf_path))[0]

    # Optional: quantity (some QR payloads may want it later)
    try:
        qty = int(get_order_quantity(order_number, spreadsheet))
    except Exception:
        qty = 1

    with open(original_pdf_path, 'rb') as f:
        raw = f.read()

    reader = PdfReader(BytesIO(raw))
    writer = PdfWriter()
    if not reader.pages:
        logging.debug("No pages found in PDF.")
        return

    page0 = reader.pages[0]
    w = float(page0.mediabox.width)
    h = float(page0.mediabox.height)

    # layout constants
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

    # move/shrink original first page
    tx = left_margin
    ty = bottom_margin
    page0.add_transformation([scale, 0, 0, scale, tx, ty])
    writer.add_page(page0)
    for p in reader.pages[1:]:
        writer.add_page(p)

    # Build overlay page
    pdf = FPDF(unit='pt', format=(w, h))
    pdf.set_auto_page_break(False)
    pdf.add_page()
    pdf.set_font('Arial', 'B', 8)

    # ---------- Top 5 QRs ----------
    top_size = int(size * 0.75)
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
            txt = "Machine Start"
        elif idx == 1:
            url = shorten_url(EMB_QUEUE_URL + urllib.parse.quote(order_number))
            txt = "Queue Design"
        elif idx == 2:
            # Order Folder on Drive (kept your original approach)
            base_folder = "https://drive.google.com/drive/folders/1n6RX0SumEipD5Nb3pUIgO5OtQFfyQXYz"
            url = f"{base_folder}/{urllib.parse.quote(order_number)}"
            txt = "Order Folder"
        elif idx == 3:
            url = shorten_url(THREAD_DATA_URL + urllib.parse.quote(order_number))
            txt = "Thread Data"
        else:
            url = SHIP_URL + urllib.parse.quote(order_number)
            txt = "Ship PDF"

        qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_L, box_size=6, border=0)
        qr.add_data(url)
        qr.make(fit=True)
        img = qr.make_image().convert('1', dither=Image.NONE)
        tmp = f"qr_top_{idx}.png"; img.save(tmp)
        pdf.set_xy(x, y_lbl)
        pdf.cell(top_size, label_h, txt, 0, 2, 'C')
        pdf.image(tmp, x, y_img, w=top_size, h=top_size)
        os.remove(tmp)

    # ---------- Right-side 5 QRs ----------
    x0 = w - box_w - pad
    y_start = pad + box_h + pad

    side_links = [
        ("Fur List",        FUR_URL),
        ("Cut List",        CUT_URL),
        ("Print List",      PRINT_LIST_URL),
        ("Embroidery List", EMB_LIST_URL),
        ("Shipping",        SHIP_URL),  # handled specially below
    ]

    # Fetch Cut Type for this order (for display above "Cut List")
    cut_type = None
    try:
        cut_type = get_cut_type_for_order(order_number, spreadsheet)
    except Exception as e:
        logging.debug("Could not fetch Cut Type for order %s: %s", order_number, e)

    for idx, (label, base_url) in enumerate(side_links):
        # Use the raw app URL so the scanner parser sees the event/&order= params
        url = base_url + urllib.parse.quote(order_number)
        # (Shipping stays the same, but this line already produces the same result)


        qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_L, box_size=6, border=0)
        qr.add_data(url)
        qr.make(fit=True)
        img = qr.make_image().convert('1', dither=Image.NONE)
        tmp = f"qr_bot_{idx}.png"; img.save(tmp)

        y_lab = y_start + idx * (label_h + size + pad + 5)
        y_img2 = y_lab + label_h

        # Add Cut Type label above Cut List QR code and label
        if label == "Cut List" and cut_type:
            pdf.set_font('Arial', 'B', label_h)
            pdf.set_xy(x0, y_lab - label_h - 2)
            pdf.cell(size, label_h, f"Cut Type: {cut_type}", 0, 2, 'C')

        pdf.set_font('Arial', 'B', label_h)
        pdf.set_xy(x0, y_lab)
        pdf.cell(size, label_h, label, 0, 2, 'C')
        pdf.image(tmp, x0, y_img2, w=size, h=size)
        os.remove(tmp)

    # ---------- Top-right info box ----------
    box_w, box_h = 60, 120
    x_box, y_box = w - box_w - pad, pad
    pdf.set_line_width(1)
    pdf.set_draw_color(0,0,0)
    pdf.rect(x_box, y_box, box_w, box_h)

    content_x, inner_w = x_box + 2, box_w - 4
    rows = [20, 20, 20, 14, 24, 22]
    offs = [5, 6, 6, 3, 6, 6]
    starts = [y_box]
    for r in rows[:-1]:
        starts.append(starts[-1] + r)
    ys = [starts[i] + offs[i] for i in range(len(rows))]

    pdf.set_font('Arial','BU',10)
    pdf.set_text_color(0,0,0)
    pdf.set_xy(content_x, ys[0]); pdf.cell(inner_w, rows[0], 'Sewing', 0, 2, 'C')

    pdf.set_font('Arial','B',9)
    pdf.set_xy(content_x, ys[1]); pdf.cell(inner_w, rows[1], f"QTY: {qty}", 0, 2, 'C')

    fur = re.sub(r"\s*fur\s*$", '', get_fur_color(order_number, spreadsheet), flags=re.IGNORECASE).strip()
    rgb, contrast = color_to_rgb(fur), get_contrast_color(fur)

    pdf.set_font('Arial','',8)
    pdf.set_fill_color(*rgb); pdf.rect(content_x, starts[2], inner_w, rows[2], style='F')
    pdf.set_text_color(*contrast); pdf.set_xy(content_x, ys[2]); pdf.cell(inner_w, rows[2], fur, 0, 2, 'C')

    pdf.set_fill_color(*rgb); pdf.rect(content_x, starts[3], inner_w, rows[3], style='F')
    pdf.set_text_color(*contrast); pdf.set_xy(content_x, ys[3]); pdf.cell(inner_w, rows[3], 'Fur', 0, 2, 'C')

    product = get_product(order_number, spreadsheet)
    pdf.set_font('Arial', 'B', 9); pdf.set_text_color(0, 0, 0)
    pdf.set_xy(content_x, ys[4])
    line_h = pdf.font_size * 1.01
    pdf.multi_cell(w=inner_w, h=line_h, txt=product, border=0, align='C', fill=False)

    pdf.set_font('Arial','B',10)
    pdf.set_xy(content_x, ys[5])
    pdf.cell(inner_w, rows[5], f"Ship: {get_due_date(order_number, spreadsheet)}", 0, 2, 'C')

    # ---------- Notes box ----------
    note_pad = 10
    note_h   = bottom_margin - 2 * note_pad
    note_y   = h - bottom_margin + note_pad
    note_x   = pad
    note_w   = w - 2 * pad

    pdf.set_fill_color(255, 255, 255)
    pdf.rect(note_x, note_y, note_w, note_h, style='F')
    pdf.set_draw_color(0, 0, 0)
    pdf.set_line_width(1)
    pdf.rect(note_x, note_y, note_w, note_h, style='D')

    note_text = get_notes(order_number, spreadsheet)
    pdf.set_text_color(0, 0, 0)
    pdf.set_font('Arial', '', 10)
    pdf.set_xy(note_x + 5, note_y + 5)
    pdf.multi_cell(note_w - 10, 12, f"Notes: {note_text or '-'}")

    # ---------- Merge overlay ----------
    overlay_stream = BytesIO(pdf.output(dest='S').encode('latin-1'))
    overlay_reader = PdfReader(overlay_stream)
    overlay_page   = overlay_reader.pages[0]

    stamped_path = os.path.join(OUTPUT_FOLDER, f"{base}_Stamped.pdf")
    final_writer = PdfWriter()

    first_page = writer.pages[0]
    first_page.merge_page(overlay_page)
    final_writer.add_page(first_page)
    for p in writer.pages[1:]:
        final_writer.add_page(p)

    with open(stamped_path, 'wb') as of:
        final_writer.write(of)
    logging.debug("Stamped PDF saved: %s", stamped_path)

    # cleanup temp (source) file
    try:
        os.remove(original_pdf_path)
        logging.debug("Source PDF removed: %s", original_pdf_path)
    except Exception as e:
        logging.debug("Couldn't remove source PDF (%s): %s", original_pdf_path, e)

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
        # MOVE the just‑processed PDF into “processed/” so Watchdog won’t see it again
        import shutil, os
        processed_dir = os.path.join(os.path.dirname(event.src_path), 'processed')
        os.makedirs(processed_dir, exist_ok=True)
        shutil.move(event.src_path, os.path.join(processed_dir, os.path.basename(event.src_path)))


from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

def monitor_drive_folder(folder_id, sh, ss):
    global sheet_thread, sheet_spreadsheet
    sheet_thread, sheet_spreadsheet = sh, ss

    creds_path = r"C:\Users\eckar\Desktop\OrderEntry,Inventory,QR,PrintPDF\Keys\poetic-logic-454717-h2-3dd1bedb673d.json"
    creds = Credentials.from_service_account_file(
        creds_path,
        scopes=['https://www.googleapis.com/auth/drive']
    )
    service = build('drive', 'v3', credentials=creds)

    stamped_folder_id = "193Mhu2QUb4tQyncBsc46lKIAF65Sukkw"  # Stamped PS
    printed_folder_id = "1HmEUhJI1RRsZXAqL13_bgJEzNYmp6x-1"  # Printed PS
    archive_folder_id = "1TOvAxuqKmiTBtptqi3oTj6GS5WuvWAMh"  # Archived PS
    orders_root_folder_id = "1n6RX0SumEipD5Nb3pUIgO5OtQFfyQXYz"  # Orders folder

    processed_ids = set()
    os.makedirs("temp_downloads", exist_ok=True)

    while True:
        try:
            results = service.files().list(
                q=f"'{folder_id}' in parents and mimeType='application/pdf'",
                spaces='drive',
                fields='files(id, name, modifiedTime)',
                orderBy='modifiedTime desc'
            ).execute()
            items = results.get('files', [])

            for file in items:
                file_id = file['id']
                name = file['name']

                if file_id in processed_ids:
                    continue
                if name.endswith('_Stamped.pdf') or name.startswith('~$') or name.startswith('.') or name.startswith('_'):
                    continue  # Skip already processed or temp files

                print(f"⬇️ New PDF found on Drive: {name}")
                request = service.files().get_media(fileId=file_id)
                local_path = os.path.join("temp_downloads", name)

                with open(local_path, 'wb') as f:
                    downloader = MediaIoBaseDownload(f, request)
                    done = False
                    while not done:
                        _, done = downloader.next_chunk()

                order = clean_value(os.path.splitext(name)[0])
                usage = extract_thread_usage(local_path)
                if not usage:
                    print(f"⚠️ No thread usage found in {name}, skipping.")
                    continue

                rows = [
                    [order,
                     r[1],
                     r[2],
                     r[3] * get_order_quantity(order, sheet_spreadsheet),
                     r[4]]
                    for r in usage
                ]
                update_sheet(sheet_thread, rows)
                shrink_page_and_stamp_horizontal_qrs(local_path, order, sheet_spreadsheet)

                # Upload stamped version to Stamped PS folder (ensure it exists)
                stamped_path = os.path.join(
                    OUTPUT_FOLDER, os.path.splitext(name)[0] + '_Stamped.pdf'
                )

                if not os.path.exists(stamped_path):
                    # brief grace period (in case of slow disk) then recheck once
                    time.sleep(0.5)
                    if not os.path.exists(stamped_path):
                        print(f"⚠️ Stamped file not found yet for {name}. Skipping this cycle.")
                        continue  # do not error or mark processed; we'll see it next loop

                file_metadata = {
                    'name': os.path.basename(stamped_path),
                    'parents': [stamped_folder_id]
                }
                media = MediaFileUpload(stamped_path, mimetype='application/pdf')

                uploaded = service.files().create(
                    body=file_metadata,
                    media_body=media,
                    fields='id',
                    supportsAllDrives=True
                ).execute()


                print(f"✅ Uploaded stamped version to Stamped PS: {uploaded['id']}")

                # Simulate sending to printer
                print("🖨️ Simulating print... (placeholder)")
                time.sleep(2)
                print("✅ Dummy print complete.")

                # Move to Printed PS folder
                service.files().update(
                    fileId=uploaded['id'],
                    addParents=printed_folder_id,
                    removeParents=stamped_folder_id,
                    supportsAllDrives=True,
                    fields='id, parents'
                ).execute()
                print("📦 Moved file to Printed PS folder.")

                # Move stamped file to order subfolder in Orders
                order_folder_name = os.path.splitext(name)[0]

                # Search for subfolder
                query = f"'{orders_root_folder_id}' in parents and name = '{order_folder_name}' and mimeType = 'application/vnd.google-apps.folder'"
                response = service.files().list(
                    q=query,
                    spaces='drive',
                    fields='files(id, name)',
                    supportsAllDrives=True
                ).execute()
                order_folders = response.get('files', [])

                if order_folders:
                    order_subfolder_id = order_folders[0]['id']
                    service.files().update(
                        fileId=uploaded['id'],
                        addParents=order_subfolder_id,
                        removeParents=printed_folder_id,
                        supportsAllDrives=True,
                        fields='id, parents'
                    ).execute()
                    print(f"📂 Filed into Order subfolder: {order_folder_name}")
                else:
                    print(f"⚠️ No subfolder found for order '{order_folder_name}' in Orders folder.")

                # Move original file to Archive folder instead of deleting
                service.files().update(
                    fileId=file_id,
                    addParents=archive_folder_id,
                    removeParents=folder_id,
                    supportsAllDrives=True,
                    fields='id, parents'
                ).execute()
                print(f"📁 Moved original to Archived PS: {file_id}")

                processed_ids.add(file_id)

        except Exception as e:
            print("Drive polling error:", e)

        time.sleep(10)
# ================================
# SECTION X: PDF Presentation Builder (Ctrl+Alt+D)
# ================================
from pathlib import Path
import tempfile
import re
from collections import Counter

# --- small helpers ---
def _with_retries(fn, *, retries=40, delay=0.25, busy_hresults=(-2147417846,)):
    import time as _t
    last_err = None
    for _ in range(retries):
        try:
            return fn()
        except Exception as e:
            hr = None
            try:
                hr = int(getattr(e, 'hresult', None) or e.args[0] or 0)
            except Exception:
                pass
            if hr in busy_hresults:
                _t.sleep(delay)
                continue
            last_err = e
            break
    if last_err is None:
        last_err = RuntimeError("COM call failed after retries (unknown error).")
    raise last_err

def _ensure_photoshop():
    pythoncom.CoInitialize()
    return gencache.EnsureDispatch("Photoshop.Application")

def _ensure_folder(p: Path):
    p.mkdir(parents=True, exist_ok=True)
    return p

def _next_version_name(folder: Path, base: str) -> str:
    patt = re.compile(rf"^{re.escape(base)}(\d+)\.pdf$", re.IGNORECASE)
    max_n = 0
    for p in folder.glob("*.pdf"):
        m = patt.match(p.name)
        if m:
            try:
                n = int(m.group(1))
                if n > max_n:
                    max_n = n
            except ValueError:
                pass
    return f"{base}{max_n + 1}"

from tkinter import filedialog, messagebox, simpledialog, ttk  # keep near top if not present

def _pick_parent_folder_with_root(root):
    folder = filedialog.askdirectory(
        parent=root,
        title="Select the PARENT folder containing your designs"
    )
    return Path(folder) if folder else None

def _pick_files_in_folder_with_root(root, initialdir: Path):
    selected = filedialog.askopenfilenames(
        parent=root,
        title="Select files for PDF (order = selection order)",
        initialdir=str(initialdir),
        filetypes=[
            ("All files", "*.*"),
            ("Images & Photoshop", "*.psd *.psb *.png *.jpg *.jpeg *.tif *.tiff"),
            ("Photoshop", "*.psd *.psb"),
            ("Images", "*.png *.jpg *.jpeg *.tif *.tiff"),
        ],
    )

    # Normalize across Tk variants
    if isinstance(selected, str):
        selected = root.tk.splitlist(selected) if selected else ()

    files = [Path(f) for f in selected if f]
    print(f"[Picker] Selected {len(files)} file(s).")
    if not files:
        return []  # user canceled at OS dialog

    # Show a loading modal while thumbnails are prepared
    loading_win, loading_pb, set_progress = _open_loading_modal(root, total_count=len(files), text="Loading previews…")
    try:
        reordered = _reorder_files_dialog_with_root(root, files, on_progress=set_progress, loading_win=loading_win)
    finally:
        _close_loading_modal(loading_win, loading_pb)

    if not reordered and files:
        print("[Picker] Reorder dialog returned empty; keeping original selection.")
        return files
    return reordered




def _reorder_files_dialog_with_root(root, files, on_progress=None, loading_win=None):
    """
    Thumbnail reorder dialog:
      - Scrollable rows with 96x96 preview + filename
      - Click row to select; Up/Down buttons reorder
      - Returns reordered list of Path objects
    """
    import tkinter as tk
    from tkinter import ttk
    from PIL import Image, ImageTk, ImageDraw

    # Try PSD support; fall back to placeholders if psd-tools not installed
    try:
        from psd_tools import PSDImage
        _PSD_OK = True
    except Exception:
        _PSD_OK = False

    THUMB = (96, 96)

    def _make_placeholder(ext_text="PSD"):
        img = Image.new("RGB", THUMB, "gray")
        drw = ImageDraw.Draw(img)
        txt = (ext_text or "").upper()[:4]
        w = drw.textlength(txt)
        drw.text(((THUMB[0]-w)//2, (THUMB[1]-12)//2), txt, fill="white")
        return img

    def _load_thumb(p):
        ext = p.suffix.lower()
        im = None
        try:
            if ext in {".png", ".jpg", ".jpeg", ".tif", ".tiff"}:
                with Image.open(p) as src:
                    src = src.convert("RGB") if src.mode not in ("RGB","RGBA") else src
                    im = src.copy()
            elif ext in {".psd", ".psb"} and _PSD_OK:
                psd = PSDImage.open(str(p))
                try:
                    pil = psd.composite()
                except AttributeError:
                    pil = psd.compose()
                pil = pil.convert("RGB") if pil.mode not in ("RGB","RGBA") else pil
                im = pil

        except Exception:
            im = None
        if im is None:
            tag = "PSD" if ext in {".psd",".psb"} else (ext[1:] or "?")
            im = _make_placeholder(tag)
        im.thumbnail(THUMB, Image.LANCZOS)
        return ImageTk.PhotoImage(im, master=win)

    win = tk.Toplevel(root)
    win.title("Reorder Files")
    win.geometry("720x520")
    win.attributes("-topmost", True)

    _cancelled = {"value": False}
    def _early_close():
        _cancelled["value"] = True
        try:
            win.destroy()
        except:
            pass
    win.protocol("WM_DELETE_WINDOW", _early_close)

    outer = ttk.Frame(win)
    outer.pack(fill=tk.BOTH, expand=True)

    canvas = tk.Canvas(outer, highlightthickness=0)
    vbar   = ttk.Scrollbar(outer, orient=tk.VERTICAL, command=canvas.yview)
    canvas.configure(yscrollcommand=vbar.set)
    vbar.pack(side=tk.RIGHT, fill=tk.Y)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    inner = ttk.Frame(canvas)
    inner_id = canvas.create_window((0,0), window=inner, anchor="nw")

    def _on_config(_e=None):
        canvas.configure(scrollregion=canvas.bbox("all"))
        bbox = canvas.bbox(inner_id)
        if bbox:
            canvas.itemconfig(inner_id, width=canvas.winfo_width())

    inner.bind("<Configure>", _on_config)
    canvas.bind("<Configure>", _on_config)

    entries = []                 # [{path, frame, img_label, name_label, photo}]
    selected = {"set": set()}    # multiselect indices

    def _update_highlights():
        sel = selected["set"]
        for j, e in enumerate(entries):
            bg = "#d6ebff" if j in sel else "#f8f8f8"
            e["frame"].configure(background=bg)
            e["name_label"].configure(background=bg)

    def _select_only(i: int):
        selected["set"] = {i}
        _update_highlights()

    def _toggle_select(i: int):
        sel = selected["set"]
        if i in sel:
            sel.remove(i)
            if not sel:
                sel.add(i)  # keep at least one selected
        else:
            sel.add(i)
        _update_highlights()

    def _scroll_to_selected():
        """Ensure at least one selected row is visible inside the canvas viewport."""
        try:
            if not entries or not selected["set"]:
                return
            i = min(selected["set"])  # bring the first selected into view
            row = entries[i]["frame"]

            canvas.update_idletasks()

            row_y = row.winfo_y()
            row_h = row.winfo_height()

            view_top = canvas.canvasy(0)
            view_h   = canvas.winfo_height()
            view_bot = view_top + view_h

            new_top = None
            if row_y < view_top:
                new_top = max(0, row_y - 8)
            elif (row_y + row_h) > view_bot:
                new_top = max(0, row_y + row_h - view_h + 8)

            if new_top is not None:
                bbox = canvas.bbox("all")
                total_h = (bbox[3] - bbox[1]) if bbox else 1
                frac = 0.0 if total_h <= 0 else min(1.0, max(0.0, new_top / total_h))
                canvas.yview_moveto(frac)
        except Exception:
            pass

    def _row_click(i, event=None):
        # Ctrl-click toggles; plain click selects one
        if event and (event.state & 0x4):  # 0x4 == Control key
            _toggle_select(i)
        else:
            _select_only(i)
        _scroll_to_selected()

    def _render_rows():
        for e in entries:
            e["frame"].pack_forget()
        for i, e in enumerate(entries):
            e["frame"].pack(fill=tk.X, padx=8, pady=4)
            e["frame"].bind("<Button-1>",      lambda ev, k=i: _row_click(k, ev))
            e["img_label"].bind("<Button-1>",  lambda ev, k=i: _row_click(k, ev))
            e["name_label"].bind("<Button-1>", lambda ev, k=i: _row_click(k, ev))

        if not selected["set"] and entries:
            selected["set"] = {0}
        _update_highlights()
        _on_config()
        _scroll_to_selected()


    built = 0
    for p in files:
        row = tk.Frame(inner, background="#f8f8f8")
        photo = _load_thumb(p)
        img_lbl  = tk.Label(row, image=photo, background="#f8f8f8")
        name_lbl = tk.Label(row, text=p.name, anchor="w", background="#f8f8f8")
        img_lbl.pack(side=tk.LEFT, padx=8, pady=6)
        name_lbl.pack(side=tk.LEFT, padx=10, pady=6)
        entries.append({"path": p, "frame": row, "img_label": img_lbl, "name_label": name_lbl, "photo": photo})

        # keep the loading spinner alive + update the counter
        built += 1
        if on_progress:
            try:
                on_progress(built)
            except:
                pass
    _render_rows()

    # Now that the dialog is fully ready, close the loading modal
    if loading_win:
        try: loading_win.destroy()
        except: pass

    btns = ttk.Frame(win)
    btns.pack(fill=tk.X, padx=10, pady=10)

    def move_up():
        sel = selected["set"]
        if not sel or 0 in sel:
            return
        i = 0
        while i < len(entries) - 1:
            if i not in sel and (i + 1) in sel:
                entries[i], entries[i + 1] = entries[i + 1], entries[i]
                sel.remove(i + 1); sel.add(i)
                i += 2
            else:
                i += 1
        _render_rows()
        _scroll_to_selected()

    def move_down():
        sel = selected["set"]
        if not sel or (len(entries) - 1) in sel:
            return
        i = len(entries) - 1
        while i > 0:
            if i not in sel and (i - 1) in sel:
                entries[i], entries[i - 1] = entries[i - 1], entries[i]
                sel.remove(i - 1); sel.add(i)
                i -= 2
            else:
                i -= 1
        _render_rows()
        _scroll_to_selected()


    done = {"ok": False}
    def on_ok():
        done["ok"] = True
        win.destroy()
    def on_cancel():
        done["ok"] = False
        win.destroy()

    ttk.Button(btns, text="Up", command=move_up).pack(side=tk.LEFT, padx=5)
    ttk.Button(btns, text="Down", command=move_down).pack(side=tk.LEFT, padx=5)
    ttk.Button(btns, text="Cancel", command=on_cancel).pack(side=tk.RIGHT, padx=5)
    ttk.Button(btns, text="OK", command=on_ok).pack(side=tk.RIGHT, padx=5)

    win.bind("<Return>", lambda e: on_ok())
    win.bind("<Escape>", lambda e: on_cancel())
    win.protocol("WM_DELETE_WINDOW", on_cancel)

    win.grab_set()
    root.wait_window(win)

    if not done["ok"]:
        return files  # keep original selection
    return [e["path"] for e in entries]



def _run_pdf_presentation_picker_on_root(root):
    # Decide parent folder (Photoshop-aware)
    parent = None
    try:
        if _is_photoshop_foreground():
            parent = _get_active_ps_parent()
    except Exception:
        parent = None
    if parent is None:
        parent = _pick_parent_folder_with_root(root)
        if not parent:
            try:
                notification.notify(title="JR & Co", message="No parent folder selected.", timeout=3)
            except:
                pass
            return

    # Pick files with Tk
    files = _pick_files_in_folder_with_root(root, parent)

    # Final guard: if something still came back empty, bail with a clearer toast.
    if not files:
        try:
            notification.notify(title="JR & Co", message="No files selected. (OS dialog canceled)", timeout=3)
        except:
            pass
        return

    loading_win, loading_pb, set_progress = _open_loading_modal(root, total_count=len(files), text="Loading previews…")

    try:
        # Let the reorder dialog update the progress as it builds rows
        files = _reorder_files_dialog_with_root(root, files, on_progress=set_progress, loading_win=loading_win)
    finally:
        _close_loading_modal(loading_win, loading_pb)

    if not files:
        # (User canceled in reorder dialog — keep earlier guard behavior if you like)
        try:
            notification.notify(title="JR & Co", message="No files selected. (Reorder canceled)", timeout=3)
        except:
            pass
        return

    if not _all_within_parent(files, parent):
        try:
            notification.notify(title="JR & Co", message="All files must be inside the chosen parent folder.", timeout=5)
        except:
            pass
        return

    def build():
        global SCANNER_ENABLED, PDF_BUILD_ACTIVE
        if 'SCANNER_ENABLED' not in globals(): SCANNER_ENABLED = True
        if 'PDF_BUILD_ACTIVE' not in globals(): PDF_BUILD_ACTIVE = False

        SCANNER_ENABLED = False
        PDF_BUILD_ACTIVE = True

        # Open progress modal on the Tk thread
        run_on_tk_thread(lambda root: _progress_modal_open(root, title="Building PDF…", text="Preparing… 0%"))

        try:
            out = _build_pdf_presentation(parent, files)  # now emits progress updates (next step)
            try:
                notification.notify(title="JR & Co", message=f"PDF created: {out.name}", timeout=4)
            except:
                print(f"[PDF-Pres] Created: {out}")
        except Exception as e:
            try:
                notification.notify(title="JR & Co", message=f"PDF Presentation failed: {e}", timeout=6)
            except:
                print("[PDF-Pres] Error:", e)
            raise
        finally:
            # Close modal on Tk thread
            run_on_tk_thread(_progress_modal_close)
            PDF_BUILD_ACTIVE = False
            SCANNER_ENABLED = True

    threading.Thread(target=build, daemon=True).start()




def _all_within_parent(paths, parent: Path) -> bool:
    parent = parent.resolve()
    for p in paths:
        try:
            if parent not in p.resolve().parents and p.resolve().parent != parent:
                return False
        except Exception:
            return False
    return True

# --- Photoshop helpers ---
def _open_doc(ps, path: Path):
    return _with_retries(lambda: ps.Open(str(path)))

def _close_doc_no_save(doc):
    _with_retries(lambda: doc.Close(2))  # 2 = psDoNotSave

def _export_single_page_pdf(ps, in_path: Path, out_pdf_path: Path):
    """
    Open a file, duplicate/flatten, save as single-page PDF, close.
    Export is tuned for smaller size (JPEGQuality + downsample).
    """
    doc = _open_doc(ps, in_path)
    dup = _with_retries(lambda: doc.Duplicate(doc.Name + "_TMP_DUP", True))
    try:
        _with_retries(lambda: dup.Flatten())
    except Exception:
        pass

    pdf_opts = win32com.client.Dispatch("Photoshop.PDFSaveOptions")

    def _safe_set(obj, attr, value):
        try:
            setattr(obj, attr, value)
        except Exception:
            pass

    # Size-oriented knobs (keep these modest; Ghostscript will do the heavy lift)
    _safe_set(pdf_opts, "PreserveEditing", False)
    _safe_set(pdf_opts, "OptimizeForWeb", True)
    _safe_set(pdf_opts, "View", False)
    _safe_set(pdf_opts, "Layers", False)
    _safe_set(pdf_opts, "EmbedColorProfile", True)

    # Lower JPEG quality to shrink pages. Range 0–12; 6–8 is a sweet spot.
    _safe_set(pdf_opts, "JPEGQuality", 7)

    # Ask PS to downsample large images inside the page (some builds support this)
    _safe_set(pdf_opts, "DownSample", True)
    _safe_set(pdf_opts, "DownSampleSize", 200.0)  # target ~200 DPI

    _with_retries(lambda: dup.SaveAs(str(out_pdf_path), pdf_opts, True))
    _close_doc_no_save(dup)
    _close_doc_no_save(doc)


def _build_pdf_presentation(parent: Path, paths: list[Path]) -> Path:
    """
    Make a multi-page PDF from `paths`, saved/versioned in `parent`.
    Returns the final (possibly Ghostscript-compressed) output PDF path.
    """
    from PyPDF2 import PdfMerger
    _ensure_folder(parent)
    base = parent.name + "Designs"
    out_stem = _next_version_name(parent, base)

    out_pdf_raw = parent / f"{out_stem}_raw.pdf"  # temporary merged (uncompressed)
    out_pdf = parent / f"{out_stem}.pdf"          # final name

    temp_dir = _ensure_folder(Path(tempfile.gettempdir()) / f"ps_pdf_build_{int(time.time())}")
    temp_pdfs = []

    ps = _ensure_photoshop()
    try:
        _with_retries(lambda: setattr(ps, "DisplayDialogs", 3))
    except Exception:
        pass

    try:
        # Export each page
        # Continuous progress helper
        def _pct_export(i, total):
            # Export is 85% of the bar
            if total <= 0: return 0.0
            return min(85.0, (i / total) * 85.0)

        total_pages = len(paths)

        for idx, p in enumerate(paths, start=1):
            tmp = temp_dir / f"page_{idx:03d}.pdf"
            print(f"[PDF-Pres] {idx}/{total_pages} {p.name} -> {tmp.name}")
            # Update progress BEFORE/AFTER long work so the user sees movement
            run_on_tk_thread(lambda root, i=idx-1: _progress_modal_update(root, _pct_export(i, total_pages), f"Exporting… {i}/{total_pages}"))
            _export_single_page_pdf(ps, p, tmp)
            temp_pdfs.append(tmp)
            run_on_tk_thread(lambda root, i=idx: _progress_modal_update(root, _pct_export(i, total_pages), f"Exporting… {i}/{total_pages}"))

        # --- MERGE ---
        print(f"[PDF-Pres] Merging {len(temp_pdfs)} pages…")
        run_on_tk_thread(lambda root: _progress_modal_update(root, 90.0, "Merging pages…"))
        merger = PdfMerger()
        for pdf in temp_pdfs:
            merger.append(str(pdf))
        with open(out_pdf_raw, "wb") as f_out:
            merger.write(f_out)
        merger.close()

        # --- COMPRESS (Ghostscript) ---
        # If you keep Ghostscript:
        run_on_tk_thread(lambda root: _progress_modal_update(root, 95.0, "Compressing…"))
        ok, err = _ghostscript_compress(out_pdf_raw, out_pdf, quality="ebook")
        if not ok:
            print("[PDF-Pres] Ghostscript failed:", err)
            # Fall back to RAW
            if out_pdf.exists():
                try: out_pdf.unlink()
                except: pass
            out_pdf_raw.rename(out_pdf)
        else:
            pass

        # Done!
        run_on_tk_thread(lambda root: _progress_modal_update(root, 100.0, "Done"))

        if ok:
            try: out_pdf_raw.unlink(missing_ok=True)
            except: pass
            return out_pdf
        else:
            # Fallback: keep the raw merged PDF
            print(f"[PDF-Pres] Ghostscript not used ({err}). Keeping RAW PDF.")
            if out_pdf.exists():
                try: out_pdf.unlink()
                except: pass
            out_pdf_raw.rename(out_pdf)
            return out_pdf
    finally:
        for t in temp_pdfs:
            try: t.unlink(missing_ok=True)
            except: pass
        try: temp_dir.rmdir()
        except: pass


def _find_ghostscript() -> str | None:
    # Try PATH first
    for exe in ("gswin64c", "gswin32c"):
        p = shutil.which(exe)
        if p:
            return p
    # Probe common install locations
    candidates = []
    for env in ("ProgramFiles", "ProgramFiles(x86)"):
        base = os.environ.get(env)
        if not base:
            continue
        gsroot = Path(base) / "gs"
        if gsroot.exists():
            for v in sorted(gsroot.glob("gs*"), reverse=True):
                candidates.append(v / "bin" / "gswin64c.exe")
                candidates.append(v / "bin" / "gswin32c.exe")
    for c in candidates:
        if c.exists():
            return str(c)
    return None

def _ghostscript_compress(in_pdf: Path, out_pdf: Path, quality: str = "ebook") -> tuple[bool, str | None]:
    """
    Recompress PDF using Ghostscript. quality in: 'screen', 'ebook', 'printer', 'prepress', 'default'
    Returns (ok, error_message).
    """
    exe = _find_ghostscript()
    if not exe:
        return False, "Ghostscript not found"

    qmap = {
        "screen": "/screen",    # ~72 dpi
        "ebook": "/ebook",      # ~150 dpi  ← good default
        "printer": "/printer",  # ~300 dpi
        "prepress": "/prepress",
        "default": "/default",
    }
    q = qmap.get(quality.lower(), "/ebook")

    cmd = [
        exe,
        "-dBATCH", "-dNOPAUSE", "-dSAFER",
        "-sDEVICE=pdfwrite",
        "-dCompatibilityLevel=1.5",
        f"-dPDFSETTINGS={q}",
        "-dDetectDuplicateImages=true",
        # Tighten downsample targets (override preset if needed)
        "-dDownsampleColorImages=true",
        "-dDownsampleGrayImages=true",
        "-dDownsampleMonoImages=true",
        "-dColorImageResolution=150",
        "-dGrayImageResolution=150",
        "-dMonoImageResolution=300",
        f"-sOutputFile={str(out_pdf)}",
        str(in_pdf),
    ]
    try:
        subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return True, None
    except subprocess.CalledProcessError as e:
        return False, f"Ghostscript failed: {e}"

def _is_photoshop_foreground() -> bool:
    """Return True if the current foreground window belongs to Photoshop."""
    try:
        hwnd = win32gui.GetForegroundWindow()
        if not hwnd:
            return False

        # Primary: process-name detection (most reliable)
        try:
            _tid, pid = win32process.GetWindowThreadProcessId(hwnd)
            if pid and psutil is not None:
                name = psutil.Process(pid).name() or ""
                # e.g., "Photoshop.exe" or "Adobe Photoshop 2025.exe"
                if "photoshop" in name.lower():
                    return True
        except Exception:
            pass

        # Fallback A: class name
        try:
            cls = win32gui.GetClassName(hwnd) or ""
            # Photoshop main windows often use these class names:
            #  - "Photoshop", "PhotoshopFrameClass"
            if cls.lower().startswith("photoshop"):
                return True
        except Exception:
            pass

        # Fallback B: window title contains "Photoshop"
        try:
            title = win32gui.GetWindowText(hwnd) or ""
            if "photoshop" in title.lower():
                return True
        except Exception:
            pass

        return False
    except Exception:
        return False


def _get_active_ps_parent() -> Path | None:
    """
    If Photoshop is running and an active doc is saved, return its parent folder.
    Returns None if no saved active doc is available.
    """
    try:
        ps = _ensure_photoshop()
        doc = _with_retries(lambda: ps.ActiveDocument)
        fullpath_str = _with_retries(lambda: str(doc.FullName))
        p = Path(fullpath_str)
        return p.parent if p.exists() else None
    except Exception:
        return None



# --- Hotkey handler ---
def _run_pdf_presentation_picker():
    # Redirect old entrypoint to safe Tk-thread version
    run_on_tk_thread(_run_pdf_presentation_picker_on_root)
    """
    Ctrl+Alt+D:
      - If Photoshop is foreground and active doc has a saved path,
        use that doc's parent folder (skip parent-folder prompt).
      - Otherwise, ask for parent folder.
      - Then ask for files within that parent.
      - Build multi-page PDF named {Parent}Designs{N}.pdf in that parent.
    """
    # Gate scanner/clipboard during the PDF flow so other hotkeys keep working cleanly
    global SCANNER_ENABLED, PDF_BUILD_ACTIVE
    # Provide safe defaults if not defined elsewhere
    if 'SCANNER_ENABLED' not in globals():
        SCANNER_ENABLED = True
    if 'PDF_BUILD_ACTIVE' not in globals():
        PDF_BUILD_ACTIVE = False

    SCANNER_ENABLED = False
    PDF_BUILD_ACTIVE = True
    try:
        # Decide parent folder
        parent = None
        if _is_photoshop_foreground():
            parent = _get_active_ps_parent()
        if parent is None:
            parent = _pick_parent_folder()
            if not parent:
                return

        # Pick files from the chosen/derived parent (shows reorder popup in your updated flow)
        files = _pick_files_in_folder(parent)
        if not files:
            try:
                notification.notify(title="JR & Co", message="No files selected.", timeout=3)
            except:
                pass
            return

        # Ensure all files live inside parent (or its subfolders)
        if not _all_within_parent(files, parent):
            try:
                notification.notify(title="JR & Co", message="All files must be inside the chosen parent folder.", timeout=5)
            except:
                pass
            return

        # Build
        try:
            out = _build_pdf_presentation(parent, files)
            try:
                notification.notify(title="JR & Co", message=f"PDF Presentation created:\n{out}", timeout=6)
            except:
                print(f"[PDF-Pres] Created: {out}")
        except Exception as e:
            try:
                notification.notify(title="JR & Co", message=f"PDF Presentation failed: {e}", timeout=6)
            except:
                print("[PDF-Pres] Error:", e)
    finally:
        # Re-enable scanner/clipboard when we’re done
        PDF_BUILD_ACTIVE = False
        SCANNER_ENABLED = True


def listen_for_ctrl_alt_d():
    """Thread-target: register Ctrl+Alt+D for PDF Presentation builder and wait."""
    # Run the picker in a background thread so this hotkey doesn't block others
    keyboard.add_hotkey('ctrl+alt+d', lambda: _run_async(_run_pdf_presentation_picker))
    keyboard.wait()

# ================================
# SECTION 9: Scanner / QR Logic
# ================================

import re
import webbrowser
import urllib.parse
import threading
import time
import keyboard
import pyperclip

# Tuning for scanner keystroke accumulation (no-Enter scanners)
TIME_THRESHOLD = 0.12   # seconds: gap larger than this = start a new buffer
_accumulated   = ""
_last_time     = 0.0

# Scan + open dedupe
DEDUP_WINDOW_SEC = 3.0        # treat identical scans/URLs within this window as duplicates
_recent_scan_at  = {}         # url -> last handled timestamp
_last_opened_at  = {}         # url -> last opened timestamp
_scan_token      = 0          # increments when Enter is handled to cancel delayed handlers

_URL_RE = re.compile(r'(https?://[^\s]+)', re.IGNORECASE)


# De-duplicate identical URLs opened within a short window.
# This prevents multiple tabs when both keyboard + clipboard listeners fire,
# or when a handler races with a retry.
DEDUP_WINDOW_SEC = 3.0
_last_opened_at = {}

def open_url_in_browser(url: str):
    """Open a URL in a new browser tab (deduped within DEDUP_WINDOW_SEC)."""
    try:
        now  = time.time()
        last = _last_opened_at.get(url, 0.0)
        if now - last < DEDUP_WINDOW_SEC:
            print(f"[open_url_in_browser] Deduped within {DEDUP_WINDOW_SEC}s: {url}")
            return
        _last_opened_at[url] = now
        webbrowser.open_new_tab(url)
    except Exception as e:
        print("[open_url_in_browser] Failed:", e)


def _open_success_page(order: str):
    """Open local green page with order image + details."""
    url = f"http://127.0.0.1:5001/success?order={urllib.parse.quote(str(order))}"
    open_url_in_browser(url)

def _parse_qr_url(url: str):
    """
    Return a tuple: (action, order, tab)
      action: "DEPT" for list tabs, "SHIP" for shipping, "UNKNOWN" otherwise
      order: order number string or ""
      tab:   one of ["Fur List", "Cut List", "Print List", "Embroidery List"] or ""
    """
    try:
        from urllib.parse import urlparse, parse_qs
        p = urlparse(url)
        q = parse_qs(p.query)
        event = (q.get("event", [""])[0] or "").strip().lower()
        order = (q.get("order", [""])[0] or "").strip()

        if event in ("fur_list", "cut_list", "print_list", "embroidery_list"):
            mapping = {
                "fur_list":         "Fur List",
                "cut_list":         "Cut List",
                "print_list":       "Print List",
                "embroidery_list":  "Embroidery List",
            }
            return ("DEPT", order, mapping[event])

        if event == "ship" or "machineschedule" in (p.netloc or "").lower():
            return ("SHIP", order, "")

        return ("UNKNOWN", order, "")
    except Exception:
        return ("UNKNOWN", "", "")

def update_tab_quantity_made(order: str, tab_name: str) -> bool:
    """Set per-tab 'Quantity Made' for the given order.
       Special case: **Cut List** uses material-specific companion columns.
    """
    try:
        ws = sheet_spreadsheet.worksheet(tab_name)
        headers = [h.strip() for h in ws.row_values(1)]

        def find_col(names):
            names = {n.lower() for n in names}
            for i, h in enumerate(headers, start=1):
                if h.strip().lower() in names:
                    return i
            return None

        # Locate the order row
        order_col = find_col({"order", "order #", "order number", "#", "order no", "order id"})
        if not order_col:
            print(f"[QtyMade] '{tab_name}': could not find an Order column")
            return False

        # Scan down the order column to find the target row
        target_row = None
        order_key = clean_value(order).strip().lower()
        col_vals = ws.col_values(order_col)
        for i, v in enumerate(col_vals[1:], start=2):  # skip header row
            if clean_value(v).strip().lower() == order_key:
                target_row = i
                break

        if not target_row:
            print(f"[QtyMade] Order {order} not found in '{tab_name}'")
            return False

        # Get the job's order quantity
        try:
            qty = int(get_order_quantity(order, sheet_spreadsheet))
        except Exception:
            qty = 0

        # --- Special handling for CUT LIST: write qty next to each filled Material cell ---
        if tab_name.strip().lower() == "cut list":
            # Identify all material columns by header
            material_indices = []
            for idx, h in enumerate(headers, start=1):
                h_norm = h.strip().lower()
                h_compact = h_norm.replace(" ", "")
                if (
                    h_compact in {"material1","material2","material3","material4","material5"}
                    or h_norm == "back material"
                    or h_compact == "backmaterial"
                ):
                    material_indices.append(idx)

            if not material_indices:
                print(f"[CutQtyMade] No material columns found in '{tab_name}' headers")
                return True  # not a hard failure; still show green page

            writes = 0
            for mi in material_indices:
                mat_val = str(ws.cell(target_row, mi).value or "").strip()
                if not mat_val:
                    continue  # skip blank material cells

                qty_col = mi + 1  # companion column immediately to the right
                current = str(ws.cell(target_row, qty_col).value or "").strip()
                if current == str(qty):
                    print(f"[CutQtyMade] Skipped (already {qty}) @ row {target_row}, col {qty_col} next to '{headers[mi-1]}'")
                    continue

                ws.update_cell(target_row, qty_col, qty)
                writes += 1
                print(f"[CutQtyMade] Wrote {qty} @ row {target_row}, col {qty_col} next to '{headers[mi-1]}'")

            if writes == 0:
                print(f"[CutQtyMade] No writes performed for order {order} (materials present but companions already set or qty==blank)")
            return True  # treat as success either way

        # --- Default handling for other tabs: single 'Quantity Made' column ---
        qmade_col = find_col({"quantity made", "qty made", "made"})
        if not qmade_col:
            print(f"[QtyMade] '{tab_name}': could not find a 'Quantity Made' column")
            return False

        # Idempotent write: only update if value actually changes
        current = str(ws.cell(target_row, qmade_col).value or "").strip()
        if current == str(qty):
            print(f"[QtyMade] Skipped (already {qty}) for {order} in '{tab_name}' row {target_row}")
            return True

        ws.update_cell(target_row, qmade_col, qty)
        print(f"[QtyMade] Wrote {qty} for {order} in '{tab_name}' row {target_row}")
        return True

    except Exception as e:
        print("[QtyMade] Error:", e)
        return False

def _handle_scanned_url(url: str):
    """
    Open behavior by type:
      - SHIP: open the scanned URL (tab) + show green success
      - DEPT (fur/cut/print/embroidery): NO popup of the scanned URL; do a headless GET and continue
      - UNKNOWN: do not open the scanned URL; show red fail page
    """
    if not url:
        return

    # Decide what kind of QR this is
    action, order, tab = _parse_qr_url(url)

    if action == "SHIP":
        # Shipping flows keep the original tab open behavior
        try:
            open_url_in_browser(url)
        except Exception as e:
            print("[SCAN] Failed to open URL:", e)

    elif action == "DEPT":
        # Department flows: do NOT open the scanned Apps Script URL in a tab.
        # Touch it headlessly in case it does side work, but keep UX clean.
        try:
            requests.get(url, timeout=3)
            print("[SCAN] Headless GET OK:", url)
        except Exception as e:
            print("[SCAN] Headless GET failed:", e)

    # invalid QR → red page
    if action == "UNKNOWN":
        open_url_in_browser("http://127.0.0.1:5001/fail?msg=Invalid%20QR")
        return


    # missing order → red page
    if not order:
        open_url_in_browser("http://127.0.0.1:5001/fail?msg=Missing%20order%20%23")
        return

    # SHIPPING: open URL already done; show green page for 7s
    if action == "SHIP":
        _open_success_page(order)
        return

    # DEPARTMENT: write Quantity Made, then green/red page
    if action == "DEPT":
        ok = update_tab_quantity_made(order, tab)
        if ok:
            _open_success_page(order)
        else:
            open_url_in_browser(
                "http://127.0.0.1:5001/fail?msg=" + urllib.parse.quote(f"{tab} update failed")
            )
        return

# --- Keyboard-wedge listener (works if scanner types characters) ---
def _scanner_keyboard_hook(e):
    # Accumulate fast keystrokes; on Enter, process the URL
    global _accumulated, _last_time, _scan_token
    if e.event_type != "down":
        return
    t = time.time()
    if t - _last_time > TIME_THRESHOLD:
        _accumulated = ""
    _last_time = t

    if e.name == "enter":
        # Cancel any in-flight delayed handler by bumping the token
        _scan_token += 1
        payload = _accumulated
        _accumulated = ""
        m = _URL_RE.search(payload)
        if m:
            _handle_scanned_url(m.group(1))
        return

    mapping = {
        "space": " ", "slash": "/", "dot": ".", "question": "?",
        "ampersand": "&", "equal": "=", "minus": "-", "colon": ":",
        "underscore": "_", "backslash": "\\"
    }
    if len(e.name) == 1:
        _accumulated += e.name
    elif e.name in mapping:
        _accumulated += mapping[e.name]

    # If scanner doesn't send Enter: detect a full URL and a "pause"
    if "http" in _accumulated.lower():
        token_snapshot = _scan_token
        start_time     = _last_time

        # If enough time passes without keys, treat as end-of-scan
        def _delayed_process(snapshot, token_snapshot, start_time):
            time.sleep(TIME_THRESHOLD * 3)
            # Only proceed if no newer keys AND Enter wasn't pressed since
            if _last_time == start_time and token_snapshot == _scan_token:
                m = _URL_RE.search(snapshot)
                if m:
                    _handle_scanned_url(m.group(1))
                # reset buffer after handling
                global _accumulated
                _accumulated = ""

        threading.Thread(
            target=_delayed_process,
            args=(_accumulated, token_snapshot, start_time),
            daemon=True
        ).start()


# --- Clipboard listener (works if scanner copies URL to clipboard) ---
def clipboard_listener():
    import re, time, webbrowser
    # Only handle URLs that match our JRCO QR format:
    qr_pat = re.compile(
        r"""^https?://script\.google\.com/.+?/exec\?event=[a-z_]+&order=\$?\d+\b""",
        re.IGNORECASE,
    )
    while True:
        try:
            time.sleep(0.05)
            if not SCANNER_ENABLED:
                continue

            txt = pyperclip.paste() or ""
            txt = txt.strip()
            if not txt:
                continue

            # Ignore normal copy operations (Ctrl+C) unless it matches QR pattern
            if not qr_pat.match(txt):
                # silently ignore; DO NOT open /fail
                continue

            # At this point it's a valid JRCO QR link → handle it
            handle_scanned_qr_url(txt)

            # Optional: clear clipboard so we don't re-trigger
            try:
                pyperclip.copy("")
            except Exception:
                pass

        except Exception:
            # Never bring up a browser /fail page here.
            # Keep the listener resilient and quiet.
            time.sleep(0.2)

# ================================
# SECTION 10: Main
# ================================
import logging
import time
import threading

# --- Job number helpers (replace your existing get_active_job_number block) ---
import re
import win32gui

def _get_wilcom_title() -> str:
    try:
        hwnd = win32gui.GetForegroundWindow()
        return win32gui.GetWindowText(hwnd) or ""
    except Exception:
        return ""

def get_active_job_number() -> str:
    """
    Parse the Wilcom title bar to extract the job/order token.
    Prefers the bracket group that looks like a filename/token (has '$' or '.emb').
    Example title:
      EmbroideryStudio 2025 – Designing - [$VP Tajima TBF] - [eckardjustin@gmail.com]
    Returns 'VP' from '$VP ...'
    """
    title = _get_wilcom_title()
    groups = re.findall(r"\[(.*?)\]", title)

    if not groups:
        m = re.search(r"-\s*\[(.*?)\]\s*$", title)
        if m:
            groups = [m.group(1)]
        else:
            return ""

    # Prefer a group that contains a '$' or looks like a filename
    preferred = None
    for g in groups:
        if "$" in g or ".emb" in g.lower():
            preferred = g
            break
    if not preferred:
        # Fall back to the FIRST group (email is usually last)
        preferred = groups[0]

    token = (preferred or "").strip().split()[0]  # take first token e.g. "$VP"
    if not token:
        return ""

    token = re.sub(r"\.emb$", "", token, flags=re.IGNORECASE).strip()
    token = token.lstrip("$")  # $VP -> VP
    # If we somehow still picked the email, bail out
    if "@" in token:
        return ""
    return token


def on_control_alt_t():
    """Handle the Control+Alt+T key combination."""
    print("[Control+Alt+T] Triggered.")

    # Display the Cut Type selection window
    def set_cut_type(selection):
        nonlocal cut_type
        cut_type = selection
        root.destroy()

    root = tk.Tk()
    root.title("Cut Type Selection")
    root.geometry("600x380")
    root.minsize(600, 380)
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

    if cut_type:
        print(f"[Control+Alt+T] Selected Cut Type: {cut_type}")

        # Extract the job number from the active window title
        job_number = get_active_job_number()
        if not job_number:
            print("[Control+Alt+T] No active job found.")
            return

        print(f"[Control+Alt+T] Active Job Number: {job_number}")

        # Write the selected Cut Type to the Google Sheet
        write_cut_type_to_sheet(job_number, cut_type)
    else:
        print("[Control+Alt+T] No Cut Type selected.")

def on_control_alt_p():
    """Handle the Control+Alt+P key combination."""
    print("[Control+Alt+P] Triggered.")
    on_control_alt_t()  # Call Control+Alt+T functionality first
    print("[Control+Alt+P] Proceeding with original functionality...")
    # Add the original Control+Alt+P logic here

if __name__ == "__main__":
    # 1) Connect to Google Sheets & ensure output folder exists
    sheet_thread, sheet_spreadsheet = connect_google_sheet(SHEET_NAME, TAB_NAME)
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    # 2) Silence Flask/Werkzeug dev-server warning
    log = logging.getLogger('werkzeug')
    log.setLevel(logging.ERROR)

    # 3) Start embedded Flask queue-service on port 5001 (only once)
    flask_thread = threading.Thread(
        target=lambda: app.run(
            host="127.0.0.1",   # use loopback
            port=5001,
            debug=False,
            use_reloader=False
        ),
        daemon=True
    )
    flask_thread.start()
    print("[Startup] Embedded Flask server started on http://127.0.0.1:5001")

    GOOGLE_DRIVE_FOLDER_ID = "1OhEUfjWI-njg9lRFobeQ8q5U_ti9fhI-"

    # 4) Begin watching the folder for new PDFs
    print(f"Monitoring Google Drive folder: Production Sheets ({GOOGLE_DRIVE_FOLDER_ID})")
    folder_thread = threading.Thread(
        target=monitor_drive_folder,
        args=(GOOGLE_DRIVE_FOLDER_ID, sheet_thread, sheet_spreadsheet),
        daemon=True
    )
    folder_thread.start()
    print("[Startup] Drive monitor thread started.")

    # 5) Start listeners for scanner input
    #    a) Clipboard listener (for scanners that copy URL to clipboard)
    try:
        threading.Thread(target=clipboard_listener, daemon=True).start()
        print("[Startup] Clipboard listener started.")
    except Exception as e:
        print("[Startup] Clipboard listener failed:", e)

    #    b) Keyboard hook (for wedge scanners that type characters)
    try:
        keyboard.hook(_scanner_keyboard_hook)
        print("[Startup] Scanner hook registered (no-Enter mode).")
    except Exception as e:
        print("[Startup] Keyboard hook failed:", e)

    # 6) Hotkey listeners
    #    Ctrl+Alt+O
    threading.Thread(target=listen_for_ctrl_alt_o, daemon=True).start()
    #    Ctrl+Alt+N
    threading.Thread(target=listen_for_ctrl_alt_n, daemon=True).start()
    #    Ctrl+Alt+P (external listener from controlaltp2)
    hotkey_thread = threading.Thread(target=listen_for_hotkey, daemon=True)
    hotkey_thread.start()

    # Ensure hotkey callbacks are registered
    print("[Startup] Registering hotkeys.")
    keyboard.add_hotkey("ctrl+alt+t", on_control_alt_t)
    print("[Startup] Hotkey registered: Ctrl+Alt+T → on_control_alt_t()")
    keyboard.add_hotkey("ctrl+alt+p", on_control_alt_p)
    print("[Startup] Hotkey registered: Ctrl+Alt+P → on_control_alt_p()")

    # Use the Tk-thread safe handler for D:
    keyboard.add_hotkey("ctrl+alt+d", lambda: run_on_tk_thread(_run_pdf_presentation_picker_on_root))
    print("[Startup] Hotkey registered: Ctrl+Alt+D → Tk-dispatched PDF picker")


    # 7) Block forever so threads stay alive
    print("Service running. Press Ctrl+C to exit.")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("Shutting down…")

