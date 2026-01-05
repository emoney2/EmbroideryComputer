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
from PIL import Image as PILImage, ImageTk, ImageDraw

# Prefer psd-tools for real PSD/PSB previews
_PSD_OK = False
try:
    from psd_tools import PSDImage  # pip install psd-tools
    _PSD_OK = True
except Exception as _e:
    _PSD_OK = False
    # Optional: print once so you know why PSDs are gray
    print("[Thumbs] psd-tools not available -> PSD/PSB previews will use placeholders.", _e)

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

# Monitor placement defaults - will update on startup
# Monitor placement defaults - will update on startup
PRIMARY_X = 0
PRIMARY_Y = 0
PRIMARY_W = 0
PRIMARY_H = 0

SECONDARY_X = 0
SECONDARY_Y = 0
SECONDARY_W = 0
SECONDARY_H = 0




def _scanner_keyboard_hook_normal(e):
    """
    Normal mode: allow typing.
    When a scan burst starts, switch into lock mode immediately.
    """
    global _last_time
    if e.event_type != "down":
        return

    now = time.time()
    delta = now - _last_time
    _last_time = now

    # If this looks like a scanner burst, immediately lock keyboard
    if delta < TIME_THRESHOLD and not POPUP_LOCK_ACTIVE:
        # Turn on lock mode; keep normal mode installed too (fine),
        # lock hook will suppress typing from here.
        install_lock_hook()


def _scanner_keyboard_hook_lock(e):
    """
    Lock mode: suppress typing but still capture scan and trigger URL handler.
    Works even if scanner does NOT send Enter (uses pause-detect).
    Stays locked until popup closes (open_url_in_browser watcher turns lock off).
    """
    global _accumulated, _last_time, _scan_token

    if e.event_type != "down":
        return

    now = time.time()
    delta = now - _last_time

    # reset buffer if we had a "human pause"
    if delta > TIME_THRESHOLD:
        _accumulated = ""

    _last_time = now
    key = e.name

    mapping = {
        "space": " ", "slash": "/", "dot": ".", "question": "?",
        "ampersand": "&", "equal": "=", "minus": "-", "colon": ":",
        "underscore": "_", "backslash": "\\"
    }

    # Enter-based scanners (still supported)
    if key == "enter":
        payload = _accumulated
        _accumulated = ""
        m = _URL_RE.search(payload)
        if m:
            print("[SCAN] QR detected (lock mode / enter)", flush=True)
            _handle_scanned_url(m.group(1))
        return  # suppressed by hook

    # accumulate characters
    if len(key) == 1:
        _accumulated += key
    elif key in mapping:
        _accumulated += mapping[key]

    # ---- NO-ENTER scanners: detect URL + pause ----
    if "http" in _accumulated.lower():
        token_snapshot = _scan_token
        start_time = _last_time
        snapshot = _accumulated

        def _delayed_process():
            global _accumulated, _scan_token
            time.sleep(TIME_THRESHOLD * 3)

            # Only proceed if no newer keys since snapshot
            if _last_time == start_time and token_snapshot == _scan_token:
                m = _URL_RE.search(snapshot)
                if m:
                    print("[SCAN] QR detected (lock mode / pause)", flush=True)
                    _handle_scanned_url(m.group(1))
                _accumulated = ""

        threading.Thread(target=_delayed_process, daemon=True).start()

    return  # suppressed by hook




# Scanner/clipboard gating
SCANNER_ENABLED = True
PDF_BUILD_ACTIVE = False
# Dedicated Chrome profile for success windows (forces monitor position reliably)
CHROME_PROFILE_DIR = os.path.join(os.getenv("LOCALAPPDATA"), "JRCO_SuccessChrome")
os.makedirs(CHROME_PROFILE_DIR, exist_ok=True)


# Sentinel to suppress Drive-poller duplicate prompts right after Ctrl+Alt+P
_CTRLALTP_RECENT = {"pdf": None, "ts": 0}  # set for ~60s after stamping

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

def enable_scanner_logic():
    global SCANNER_ENABLED
    SCANNER_ENABLED = True
    print("[Scanner] Logic re-enabled")



import time
import pygetwindow as gw
from screeninfo import get_monitors

def move_success_window_to_monitor2():
    print("[MoveSuccess] starting helper…", flush=True)
    try:
        time.sleep(2.0)

        # --- get all monitors ---
        monitors = get_monitors()
        print("[MoveSuccess] detected monitors:")
        for m in monitors:
            print(f"   {m}", flush=True)

        # --- choose secondary monitor (anything not primary) ---
        primary = next((m for m in monitors if m.is_primary), None)
        secondary = next((m for m in monitors if not m.is_primary), None)

        if not secondary:
            print("[MoveSuccess] Only one monitor detected, aborting.")
            return False
        
        print(f"[MoveSuccess] secondary monitor = {secondary}", flush=True)

        # --- find all window titles ---
        titles = gw.getAllTitles()
        candidates = [t for t in titles if "success" in (t or "").lower()]
        if not candidates:
            print("[MoveSuccess] no success window found")
            return False

        for title in candidates:
            try:
                win = gw.getWindowsWithTitle(title)[0]

                # calculate centered coords for secondary display
                target_x = secondary.x + (secondary.width - 1200)//2
                target_y = secondary.y + (secondary.height - 800)//2

                print(f"[MoveSuccess] moving '{title}' to monitor2 center", flush=True)
                win.moveTo(target_x, target_y)
                win.resizeTo(1200, 800)
                print("[MoveSuccess] ok", flush=True)
                return True
            except Exception as e:
                print(f"[MoveSuccess] error: {e}", flush=True)

        return False

    except Exception as e:
        print(f"[MoveSuccess] FAILED: {e}", flush=True)
        return False




# === PDF row detection & preview rendering ===
def _detect_sequence_rows_with_pdfplumber(pdf_path, max_seq):
    """
    Returns a dict {seq_number:int -> (x0, y_center)} for page 1, where (x0, y) are
    PDF coordinates in points (72 dpi). We look for words that match '1', '2', ..., str(max_seq)
    at the start of each usage row.
    """
    import pdfplumber
    positions = {}
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if not pdf.pages:
                return positions
            page = pdf.pages[0]
            words = page.extract_words(use_text_flow=True, keep_blank_chars=False)
            # Gather candidates keyed by integer text
            buckets = {}
            for w in words:
                txt = (w.get("text") or "").strip()
                if txt.isdigit():
                    n = int(txt)
                    if 1 <= n <= max_seq:
                        buckets.setdefault(n, []).append(w)
            # Pick the leftmost word for each sequence (assumes seq number is first col)
            for n, lst in buckets.items():
                # choose smallest x0
                w = min(lst, key=lambda ww: ww.get("x0", 1e9))
                x0 = float(w.get("x0", 0))
                top = float(w.get("top", 0))
                bottom = float(w.get("bottom", top + 10))
                y_center = (top + bottom) / 2.0
                positions[n] = (x0, y_center)
    except Exception as e:
        print(f"[StopsUI] pdfplumber failed: {e}", flush=True)
    return positions

def _locate_stop_mark_positions(pdf_path: str, max_seq: int) -> dict[int, tuple[float, float]]:
    """
    Return {n -> (x_pts, y_pts)} where:
      - x_pts is just LEFT of the 'St.' column
      - y_pts is BETWEEN row n and row n+1  (so 'After n' hits the gap)
    All coordinates are in PDF points (72 dpi) for PAGE 1.

    Strategy:
      1) Find the left edge of the 'St'/'St.' header on page 1.
      2) For each sequence number row, get its (top,bottom).
      3) y_between(n) = (bottom(n) + top(n+1)) / 2
    """
    import statistics
    positions: dict[int, tuple[float, float]] = {}
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if not pdf.pages:
                return positions
            page = pdf.pages[0]
            words = page.extract_words(use_text_flow=True, keep_blank_chars=False)

            # 1) Find the 'Stop Sequence' header line, then pick the 'St.' on THAT line only.
            #    This avoids grabbing stray 'st' strings elsewhere on the page.
            #    Header looks like: "# Color St. Code Name"
            def _norm(s: str) -> str:
                return (s or "").strip().lower()

            # Group words into text lines by overlapping vertical bands
            lines: list[list[dict]] = []
            for w in words:
                placed = False
                for line in lines:
                    # if this word vertically overlaps the existing line band, add to it
                    top = float(w.get("top", 0.0)); bottom = float(w.get("bottom", 0.0))
                    ltop = min(float(x.get("top", 0.0)) for x in line)
                    lbot = max(float(x.get("bottom", 0.0)) for x in line)
                    # overlap if at least 1pt intersects
                    if not (bottom < ltop + 1.0 or top > lbot - 1.0):
                        line.append(w); placed = True; break
                if not placed:
                    lines.append([w])

            header_line = None
            for line in lines:
                texts = [_norm(w.get("text")) for w in line]
                # Require all header tokens on the same line
                if (any(t in ("#", "•", "no.", "no") for t in texts) or "#" in " ".join(texts)) and \
                   any(t == "color" for t in texts) and \
                   any(t in ("st", "st.") for t in texts) and \
                   any(t == "code" for t in texts) and \
                   any(t == "name" for t in texts):
                    header_line = line
                    break

            if header_line:
                # 'St.' x = x0 of that token on the header line
                st_tokens = [w for w in header_line if _norm(w.get("text")) in ("st", "st.")]
                if st_tokens:
                    st_x_left = float(min(st_tokens, key=lambda ww: float(ww.get("x0", 1e9)))["x0"])
                else:
                    # Shouldn't happen if header_line matched, but keep a conservative fallback
                    st_x_left = float(min(header_line, key=lambda ww: float(ww.get("x0", 1e9)))["x0"]) + 100.0
            else:
                # Last-resort fallback if header wasn't found: DO NOT roam the whole page for 'st'
                # Use a conservative center-ish default that keeps triangles in the table band once rows are found
                st_x_left = 240.0

            # Optional debug — see which line we recognized as the header:
            try:
                texts = " | ".join([w.get("text") for w in (header_line or [])])
                yband = (min(float(w.get("top", 0.0)) for w in (header_line or [])),
                         max(float(w.get("bottom", 0.0)) for w in (header_line or [])))
                print(f"[_locate_stop_sequence_rows] header='{texts}' yband={yband} st_left={st_x_left:.2f}", flush=True)
            except Exception:
                pass


            # 2) Collect row bands for 1..max_seq (top/bottom per row)
            row_bands = {}  # n -> (top, bottom)
            buckets: dict[int, list[dict]] = {}
            for w in words:
                txt = (w.get("text") or "").strip()
                if txt.isdigit():
                    n = int(txt)
                    if 1 <= n <= max_seq:
                        buckets.setdefault(n, []).append(w)
            for n, lst in buckets.items():
                # choose leftmost occurrence (usual first column)
                ww = min(lst, key=lambda ww: ww.get("x0", 1e9))
                top = float(ww.get("top", 0.0))
                bottom = float(ww.get("bottom", top + 12.0))
                row_bands[n] = (top, bottom)

            if not row_bands:
                return positions

            # typical row spacing for last-row fallback
            tops_sorted = [row_bands[k][0] for k in sorted(row_bands)]
            diffs = [b - a for a, b in zip(tops_sorted, tops_sorted[1:])]
            typical_gap = statistics.median(diffs) if diffs else 22.0

            # Small inset so the tip sits just LEFT of the St. column
            x_tip = st_x_left - 8.0

            # 3) Compute BETWEEN positions
            # For n in [1..max_seq): midpoint between bottom(n) and top(n+1)
            # For the last n: place halfway below n using typical gap
            for n in range(1, max_seq + 1):
                if (n in row_bands) and ((n + 1) in row_bands):
                    y = (row_bands[n][1] + row_bands[n + 1][0]) / 2.0
                elif n in row_bands:
                    y = row_bands[n][1] + typical_gap * 0.5
                else:
                    # if the row is missing entirely, synthesize a reasonable line
                    base_top = tops_sorted[0] if tops_sorted else 120.0
                    y = base_top + typical_gap * (n - 0.5)
                positions[n] = (x_tip, y)
    except Exception as e:
        print(f"[_locate_stop_mark_positions] failed: {e}", flush=True)
    return positions


def _detect_max_sequence(pdf_path: str) -> int:
    """
    Detect the largest Stop Sequence row number by anchoring to the Stop Sequence header line
    ("# Color St. Code Name") and then collecting ONLY the left-column integer tokens that
    appear on lines below that header and to the LEFT of the 'St.' column.
    """
    try:
        import pdfplumber, statistics
        with pdfplumber.open(pdf_path) as pdf:
            if not pdf.pages:
                return 0
            page = pdf.pages[0]
            words = page.extract_words(use_text_flow=True, keep_blank_chars=False)

            def _norm(s: str) -> str:
                return (s or "").strip().lower()

            # Group into rough lines by vertical overlap
            lines: list[list[dict]] = []
            for w in words:
                placed = False
                for line in lines:
                    top  = float(w.get("top", 0.0))
                    bot  = float(w.get("bottom", top + 10.0))
                    ltop = min(float(x.get("top", 0.0)) for x in line)
                    lbot = max(float(x.get("bottom", 0.0)) for x in line)
                    if not (bot < ltop + 1.0 or top > lbot - 1.0):
                        line.append(w); placed = True; break
                if not placed:
                    lines.append([w])

            # Find the header line that contains the Stop Sequence labels
            header_line = None
            for line in lines:
                texts = [_norm(w.get("text")) for w in line]
                if (any(t in ("#", "•", "no.", "no") for t in texts) or "#" in " ".join(texts)) and \
                   any(t == "color" for t in texts) and \
                   any(t in ("st", "st.") for t in texts) and \
                   any(t == "code" for t in texts) and \
                   any(t == "name" for t in texts):
                    header_line = line
                    break
            if not header_line:
                return 0

            # Get x of 'St.' on the header line so we only consider numbers left of it
            st_tokens = [w for w in header_line if _norm(w.get("text")) in ("st", "st.")]
            st_x_left = float(min(st_tokens, key=lambda ww: float(ww.get("x0", 1e9)))["x0"]) if st_tokens else 9999.0

            header_y_max = max(float(w.get("bottom", 0.0)) for w in header_line)

            # Collect digit tokens BELOW header and LEFT of St. column; pick the leftmost band
            num_words = []
            for w in words:
                t = (w.get("text") or "").strip()
                if not t.isdigit():
                    continue
                x0 = float(w.get("x0", 1e9))
                top = float(w.get("top", 0.0))
                if top <= header_y_max:
                    continue  # above header
                if x0 >= st_x_left:
                    continue  # to right of 'St.' col
                num_words.append(w)
            if not num_words:
                return 0

            # Lock to the *leftmost* numeric band (first column)
            left_band_x = min(float(w.get("x0", 1e9)) for w in num_words)
            band = [w for w in num_words if abs(float(w.get("x0", 1e9)) - left_band_x) <= 8.0]

            ints = []
            for w in band:
                try:
                    n = int((w.get("text") or "").strip())
                    if 1 <= n <= 200:
                        ints.append(n)
                except Exception:
                    pass
            return max(ints) if ints else 0
    except Exception:
        return 0



def prompt_stops_interactive(pdf_path: str, seq_count: int):
    """
    Modal visual picker that runs on the Tk service thread.
    Shows page 1 of the PDF with checkboxes; returns sorted list of selected rows.
    """
    import queue
    ensure_tk_service()  # make sure Tk service thread is running
    out_q: "queue.Queue[list[int]]" = queue.Queue(maxsize=1)

    def _show(root):
        import tkinter as tk
        from tkinter import ttk
        from PIL import ImageTk as PILImageTk

        # Render background image
        import fitz  # PyMuPDF
        import PIL.Image as PILImage
        doc = fitz.open(pdf_path)
        page = doc[0] if len(doc) else None
        if page is None:
            raise RuntimeError("Empty PDF")

        rect = page.rect  # in points
        fit_w = 600.0
        base_zoom = min(1.4, max(0.9, fit_w / float(rect.width)))

        mat = fitz.Matrix(base_zoom, base_zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = PILImage.frombytes("RGB", [pix.width, pix.height], pix.samples)

        pw, ph = float(rect.width), float(rect.height)
        sx = pix.width / pw
        sy = pix.height / ph
        doc.close()

        # UI-only sequence count (don’t shadow outer seq_count)
        seq_count_ui = int(seq_count) if seq_count else 0
        try:
            seq_count_ui = max(seq_count_ui, _detect_max_sequence(pdf_path))
        except Exception:
            pass


        # Toplevel (NOT another Tk)
        win = tk.Toplevel(root)
        win.title("Select Stops (click checkboxes)")
        try:
            win.attributes("-topmost", True)
        except Exception:
            pass
        win.grab_set()             # modal grab
        win.focus_force()
        win.bell()

        # Left controls; Right canvas
        left = ttk.Frame(win)
        left.pack(side="left", fill="y", padx=10, pady=10)
        right = ttk.Frame(win)
        right.pack(side="right", fill="both", expand=True, padx=10, pady=10)

        canvas = tk.Canvas(right, width=img.width, height=img.height, highlightthickness=0)
        canvas.pack(fill="both", expand=True)

        # Create the image INSIDE the same interpreter as the canvas
        photo = PILImageTk.PhotoImage(image=img, master=canvas)

        # Draw and keep strong refs so Tk doesn't GC it
        bg_img = canvas.create_image(0, 0, anchor="nw", image=photo)
        canvas._bg_photo = photo
        canvas._bg_img_id = bg_img

        default_x_px = int(36 * sx)
        triangle_items = {}          # seq -> canvas item id
        check_vars = {}              # seq -> BooleanVar

        # --- Stop-mark positions LEFT of 'St.' and BETWEEN rows (Stop Sequence block) ---
        # 1) Compute a generous upper bound so we don't cap detection at thread rows.
        try:
            guess_from_page = _detect_max_sequence(pdf_path)  # anchored to Stop Sequence header
        except Exception:
            guess_from_page = 0
        max_guess = max(int(guess_from_page or 0), int(seq_count_ui or 0), 1)   # no artificial 12-row floor

        # 2) Try BOTH locators with the larger cap and pick the one that finds more rows.
        pos_a = {}
        st_left_a = None
        try:
            pos_a, st_left_a = _locate_stop_sequence_rows(pdf_path, max_rows=max_guess)
        except Exception:
            pos_a, st_left_a = {}, None

        pos_b = _locate_stop_mark_positions(pdf_path, max_seq=max_guess)
        st_left_b = None  # _locate_stop_mark_positions already bakes in the left-tip X

        # 3) Choose best: prefer A if it found at least a few rows; else fall back to B
        if len(pos_a) >= 3:
            positions, st_left = pos_a, st_left_a
        elif len(pos_b) >= 1:
            positions, st_left = pos_b, st_left_b
        else:
            positions, st_left = {}, None


        # 4) If nothing found, fall back to the anchored detector’s number for checkbox count
        if not positions:
            seq_count_ui = max(int(guess_from_page or 0), 1)
        else:
            # Normalize checkbox count from actual detected rows
            try:
                seq_count_ui = max(int(k) for k in positions.keys())
            except Exception:
                seq_count_ui = max(len(positions), 1)


        ui_inset_pts = 3.0   # TIP sits just inside the Stop Sequence box, slightly LEFT of 'St.'
        ui_tri_size  = 14    # must match draw_triangle size for correct centering
        seq_positions_px = {}

        if positions and (st_left is not None):
            # We have the 'St.' column; compute left-tip from st_left - inset, and use row Y
            for n in range(1, seq_count_ui + 1):
                y_pts = positions.get(n)
                if y_pts is None:
                    continue
                x_tip_pts = float(st_left) - ui_inset_pts
                x_px = int(x_tip_pts * sx) + (ui_tri_size // 2)  # center the triangle around the tip
                y_px = int(float(y_pts) * sy)
                seq_positions_px[n] = (x_px, y_px)
        elif positions:
            # positions already contain coordinates; convert to pixels directly
            for k, v in positions.items():
                try:
                    n = int(k)
                except Exception:
                    continue
                x_pts, y_pts = v if isinstance(v, (list, tuple)) and len(v) == 2 else (None, None)
                if x_pts is None:
                    continue
                x_px = int(float(x_pts) * sx) + (ui_tri_size // 2)
                y_px = int(float(y_pts) * sy)
                seq_positions_px[n] = (x_px, y_px)
        else:
            # Evenly spaced fallback only if nothing detected
            top = int(120 * sy)
            step = int(24 * sy)
            for n in range(1, max(seq_count_ui, 1) + 1):
                seq_positions_px[n] = (default_x_px, top + (n - 1) * step)

        # Optional: log what the UI will show
        print(f"[StopsUI] rows_a={len(pos_a)} rows_b={len(pos_b)} chosen={len(positions)} seq_count_ui={seq_count_ui}", flush=True)



        def draw_triangle(x, y, size=14):
            # LEFT-pointing triangle (90° CCW), centered at (x, y)
            half = size // 2
            return canvas.create_polygon(
                x - half, y,            # tip (left)
                x + half, y - half,     # upper base corner
                x + half, y + half,     # lower base corner
                fill="#C00000", outline=""
            )

        def toggle_seq(n, var):
            # Remove old
            try:
                if triangle_items.get(n):
                    canvas.delete(triangle_items[n])
            except Exception:
                pass
            triangle_items[n] = None
            # Draw new if checked
            if var.get() and (n in seq_positions_px):
                x, y = seq_positions_px[n]
                triangle_items[n] = draw_triangle(x, y, size=14)

        ttk.Label(left, text=f"Detected sequences: {max(seq_count_ui, 1)}", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(0,6))


        list_frame = ttk.Frame(left)
        list_frame.pack(anchor="w", fill="y", expand=True)
        cols = 2
        for i in range(max(seq_count_ui, 1)):
            n = i + 1
            # Default ON for row 1
            v = tk.BooleanVar(master=win, value=(n == 1))
            check_vars[n] = v
            cb = ttk.Checkbutton(
                list_frame,
                text=f"After {n}",
                variable=v,
                command=lambda n=n, v=v: toggle_seq(n, v)
            )
            r, c = divmod(i, cols)
            cb.grid(row=r, column=c, sticky="w", padx=6, pady=4)

            if n == 1 and (n in seq_positions_px):
                toggle_seq(n, v)


        quick = ttk.Frame(left); quick.pack(anchor="w", pady=(6,6))
        def _all():  [check_vars[n].set(True)  or toggle_seq(n, check_vars[n]) for n in range(1, max(seq_count_ui,1)+1)]
        def _none(): [check_vars[n].set(False) or toggle_seq(n, check_vars[n]) for n in range(1, max(seq_count_ui,1)+1)]

        ttk.Button(quick, text="Select All", command=_all).pack(side="left", padx=(0,6))
        ttk.Button(quick, text="Clear",      command=_none).pack(side="left")

        # OK / Cancel
        def _finish(vals):
            # 1) Write the sidecar FIRST (sync to disk), so the stamper can read it immediately
            try:
                import os, json
                ui_tri_size = 14  # keep in sync with the preview triangle size
                marks_pdf = {}
                for n, (cx_px, cy_px) in (seq_positions_px or {}).items():
                    x_tip_pts = (float(cx_px) - (ui_tri_size // 2)) / float(sx)  # left-tip in PDF pts
                    y_pts     = float(cy_px) / float(sy)
                    marks_pdf[int(n)] = [x_tip_pts, y_pts]
                sidecar = os.path.splitext(pdf_path)[0] + ".marks.json"
                with open(sidecar, "w", encoding="utf-8") as f:
                    json.dump({"pdf_path": pdf_path, "marks": marks_pdf},
                      f, ensure_ascii=False, indent=2)
                    f.flush()
                    os.fsync(f.fileno())
                print(f"[UI->Sidecar] wrote {sidecar} with {len(marks_pdf)} marks", flush=True)
            except Exception as e:
                print(f"[UI->Sidecar] failed to write: {e}", flush=True)

            # 2) ONLY NOW unblock the caller
            try:
                out_q.put_nowait(sorted(vals))
            except Exception:
                pass

            # 3) Close the window
            try:
                win.grab_release()
            except Exception:
                pass
            win.destroy()

        def _ok():
            vals = [n for n in range(1, max(seq_count_ui,1)+1) if check_vars.get(n) and check_vars[n].get()]
            if not vals and max(seq_count_ui,1) >= 1:
                vals = [1]
            _finish(vals)


        def _cancel():
            _finish([])

        btns = ttk.Frame(left); btns.pack(anchor="e", pady=(8,0), fill="x")
        ttk.Button(btns, text="OK", command=_ok).pack(side="right")
        ttk.Button(btns, text="Cancel", command=_cancel).pack(side="right", padx=(0,8))

        win.protocol("WM_DELETE_WINDOW", _cancel)
        win.lift()
        win.update_idletasks()

        # Center the window and keep it pinned on top until closed
        try:
            sw = win.winfo_screenwidth()
            sh = win.winfo_screenheight()
            ww = min(img.width + 360, sw - 40)
            wh = min(img.height + 80, sh - 80)
            x  = max(20, (sw - ww)//2)
            y  = max(40, (sh - wh)//3)
            win.geometry(f"{ww}x{wh}+{x}+{y}")

            # Keep forcing front focus/topmost periodically so it doesn't hide behind Wilcom
            def _force_front():
                try:
                    win.attributes("-topmost", True)
                    win.lift()
                    win.focus_force()
                except Exception:
                    pass
                # repeat every 1s while the window exists
                try:
                    win.after(1000, _force_front)
                except Exception:
                    pass

            _force_front()
        except Exception:
            pass


    # Run UI on Tk thread and WAIT for result
    run_on_tk_thread(_show)
    try:
        # 30 minutes max; adjust if you like
        return out_q.get(timeout=1800)
    except Exception:
        return []
# --- Debug UI helper ---
def _debug_msgbox(title: str, text: str):
    """Show a Windows MessageBox; log errors if it fails."""
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(None, text, title, 0)
        print(f"[DEBUG] MessageBox shown: {title}", flush=True)
    except Exception as e:
        print(f"[DEBUG] MessageBox FAILED: {e}", flush=True)


# ==== TK SERVICE ==============================================================
import threading, queue

_tk_ready = threading.Event()
_tk_cmds  = queue.Queue()
_tk_root  = None
_tk_thread = None

def _tk_thread_main():
    import tkinter as tk
    global _tk_root
    _tk_root = tk.Tk()
    _tk_root.withdraw()  # keep the root hidden
    _tk_ready.set()

    def _drain():
        # Run any scheduled UI functions
        try:
            while True:
                fn = _tk_cmds.get_nowait()
                try:
                    fn(_tk_root)
                except Exception as e:
                    print(f"[Tk] handler error: {e}", flush=True)
        except queue.Empty:
            pass
        _tk_root.after(50, _drain)

    _tk_root.after(50, _drain)
    _tk_root.mainloop()

def ensure_tk_service():
    global _tk_thread
    if _tk_thread and _tk_thread.is_alive():
        return
    _tk_thread = threading.Thread(target=_tk_thread_main, name="TkService", daemon=True)
    _tk_thread.start()
    _tk_ready.wait(timeout=5)

def run_on_tk_thread(fn):
    """Schedule fn(root) to run on the Tk thread."""
    ensure_tk_service()
    _tk_cmds.put(fn)
# =============================================================================
def pick_stops_via_ui(pdf_path: str, seq_count: int, timeout: int = 600) -> list[int]:
    """
    Synchronously run prompt_stops_interactive on the Tk thread and return the selected stops.
    """
    import queue, threading, time
    ensure_tk_service()
    out_q: "queue.Queue[list[int]]" = queue.Queue(maxsize=1)

    def _task(_root):
        try:
            vals = prompt_stops_interactive(pdf_path, seq_count)
        except Exception:
            vals = []
        try:
            out_q.put_nowait(vals)
        except Exception:
            pass

    run_on_tk_thread(_task)

    # Wait for the UI to finish
    try:
        return out_q.get(timeout=timeout)
    except Exception:
        return []


def prompt_stops_for_sequences(seq_count: int, pdf_path: str | None = None):
    """
    Always show the visual PDF stop selector (checkboxes aligned to rows).
    If called by the Drive poller (pdf_path is None), we suppress immediately
    if Ctrl+Alt+P just ran (within ~60s) to avoid double prompting.
    """
    # Suppress Drive-poller duplicate prompt right after Ctrl+Alt+P
    try:
        import time
        if pdf_path is None and (_CTRLALTP_RECENT.get("ts") and (time.time() - _CTRLALTP_RECENT["ts"] < 60)):
            print("[Stops] Skipping (recent Ctrl+Alt+P handled stops).", flush=True)
            return []
    except Exception:
        pass

    if pdf_path:
        print(f"[Stops] Showing visual stop selector for {os.path.basename(pdf_path)}...", flush=True)
        try:
            return prompt_stops_interactive(pdf_path, seq_count)
        except Exception as e:
            print(f"[Stops] Visual selector failed ({e}); returning no stops.", flush=True)
            return []
    else:
        # Fallback: show the UI against a blank temp page (even spacing)
        print("[Stops] No pdf_path; visual list with even spacing.", flush=True)
        try:
            from fpdf import FPDF
            tmp = os.path.join(os.getenv("TEMP") or ".", "jrco_temp_blank.pdf")
            pdf = FPDF(unit='pt', format=(612, 792)); pdf.add_page(); pdf.output(tmp)
            vals = prompt_stops_interactive(tmp, seq_count)
            try: os.remove(tmp)
            except Exception: pass
            return vals
        except Exception:
            return []


# Start the Tk thread once at startup
threading.Thread(target=_tk_thread_main, daemon=True).start()

# ---- Percent Progress Modal (determinate) ----
_progress_modal = {"win": None, "bar": None, "txt": None}

def _progress_modal_open(root, title="Working...", text="Starting..."):
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

# ---- Tiny "Loading previews..." modal (for long UI work) ----
def _open_loading_modal(root, total_count:int|None=None, text="Loading previews..."):
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
CREDENTIALS_PATH = r"C:\Users\Embroidery1\Desktop\OrderEntry,Inventory,QR,PrintPDF - Copy\credentials.json"
MONITOR_FOLDER = r"C:\Users\Embroidery1\Desktop\Embroidery Sheets"
OUTPUT_FOLDER  = r"C:\Users\Embroidery1\Desktop\Embroidery Sheets\PrintPDF"
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



# --- Embedded Queue Service Configuration & Route ---
import subprocess
import os
import time
from flask import Flask, request
from pywinauto import Application

# Folder where your .emb files live:
EMB_FOLDER = r"C:\Users\Embroidery1\Desktop\EMB"
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

        # NEW: close just the design window, leave the app running
        win.close()

        return "Queued OK and closed design window", 200
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

    # Find the oldest image in the order's folder
    job_folder = get_job_folder(order)
    img_path, _ = find_oldest_image(job_folder)

    # Base64 encode the image (optional: if none, we'll just hide the <img>)
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
    dept = (request.args.get("dept") or "").strip()

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
      display: flex; flex-direction: column; text-align: center; align-items:center;
    }}
    .dept {{ font-size: 6rem; font-weight: 900; color: #064; margin-bottom: 0.1em; text-align: center; }}
    .order {{ font-size: 3.0rem; font-weight: 800; line-height: 1.05; margin: 0.15em 0; text-align:center; }}
    .company {{ font-size: 2.0rem; font-weight: 700; opacity: 0.95; margin: 0.1em 0; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 36vw; }}
    .mid {{ font-size: 1.6rem; font-weight: 700; opacity: 0.95; margin: 0.25em 0; }}
    .qty {{ display:inline-flex; align-items:center; justify-content:center; width:72px; height:72px; border-radius:999px; background:#000; color:#fff; font-weight:900; font-size:1.75rem; margin-top:0.5rem; }}

    /* Stack on very small screens */
    @media (max-width: 800px) {{
      .wrap {{ flex-direction: column; text-align: center; }}
      .info-col {{ text-align: center; }}
      .dept {{ font-size: 4rem; }}
      .order {{ font-size: 2.5rem; }}
      .company {{ font-size: 1.2rem; max-width: 80vw; white-space: normal; }}
      .mid {{ font-size: 1.2rem; }}
      .img-col img {{ max-width: 80vw; max-height: 50vh; }}
    }}
  </style>
</head>
<body>
  <div class="wrap">
    {("<div class='img-col'><img src='data:image/png;base64," + img_b64 + "' alt='Order image' /></div>") if img_b64 else ""}
    <div class="info-col">
      {f"<div class=\"dept\">{dept}</div>" if dept else ""}
      <div class="order">ORDER #{order}</div>
      <div class="company">{company}</div>
      <div class="mid">{design}</div>
      {f"<div class=\"qty\">{qty}</div>" if qty else ""}
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

# --- START: Ctrl+Alt+O (open .emb) routines ---
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
    """Thread-target: wait for Ctrl+Alt+O and call open_order_file."""
    while True:
        keyboard.wait(HOTKEY_O)
        open_order_file()


# --- START: Ctrl+Alt+N (Save As PDF) routines ---
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
    EXACT KEY CHORD: Alt+F -> wait -> G -> wait -> Enter (via low-level VKs),
    then: wait for file dialog -> navigate to folder -> type the selected image filename -> Enter.
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
    then File -> Import Graphic with the oldest PNG in that folder, then insert Product template if available.
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
                "Try again, or open File -> Import manually."
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
# --- END: Ctrl+Alt+N (Save As PDF) routines ---

# --- Helper: get active order # from Wilcom window (robust) ---
def _active_order_from_wilcom() -> str | None:
    """
    Tries to read the active EmbroideryStudio 2025 window title and pull the order id.
    Accepts formats like: '... - [$VP Tajima TBF] - [email]' or '... - [172] - [...]'
    Returns clean order string (keeps '$' if present) or None.
    """
    try:
        from pywinauto import Desktop
        desk = Desktop(backend="uia")
        win = desk.window(title_re=EMB_WINDOW_TITLE)
        title = (win.window_text() or "").strip()
        m = re.search(r"\[([^\]]+)\]", title)
        if m:
            raw = m.group(1).strip()
            raw = re.split(r"\s{2,}| - ", raw)[0].strip()
            raw = re.sub(r"\.emb$", "", raw, flags=re.I)
            return raw
    except Exception:
        pass
    return None


# --- Helper: save current Wilcom doc to PDF into the order folder; returns pdf path ---
def _export_current_pdf_for_order(order_number: str) -> str | None:
    """
    Uses the same keystroke flow you use elsewhere:
      - Ctrl+S (ensure doc saved)
      - Alt+N (Save As / PDF path box), type path G:\My Drive\Orders\<order>\<order>
      - press Enter
    Returns the full path '<folder>\<order>.pdf' if it appears and stabilizes, else None.
    """
    base_no_ext = os.path.join(BASE_PATH, order_number, order_number)
    target_pdf  = base_no_ext + ".pdf"

    try:
        os.makedirs(os.path.dirname(base_no_ext), exist_ok=True)
        keyboard.send('ctrl+s')
        time.sleep(2)
        keyboard.send('alt+n')
        time.sleep(0.2)
        keyboard.write(base_no_ext)
        time.sleep(0.1)
        keyboard.send('enter')
    except Exception:
        return None

    ok = wait_for_stable_file(target_pdf, max_wait=20)
    return target_pdf if ok else None


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
    # NOTE: Explicitly load the JSON key so there's no ambiguity
    creds_path = r"C:\Users\Embroidery1\Desktop\OrderEntry,Inventory,QR,PrintPDF - Copy\Keys\poetic-logic-454717-h2-3dd1bedb673d.json"
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

# --- robust write: Production Orders / Cut Type at row matching Order # ---
def write_cut_type_to_sheet(order_number: str, cut_type: str) -> bool:
    """
    Writes cut_type into 'Production Orders' tab, 'Cut Type' column,
    at the row where the 'Order #' (or 'Order Number' / 'Order') cell matches order_number.
    Matching is normalized (172 == '172' == 172.0 == '$172').
    """
    try:
        ss = sheet_spreadsheet
        if ss is None:
            print("[CutType] sheet_spreadsheet is None (not connected yet?)")
            return False

        # 1) Find the worksheet by title (case-insensitive)
        target_ws = None
        try:
            for w in ss.worksheets():
                if w.title.strip().lower() == "production orders":
                    target_ws = w
                    break
        except Exception as e:
            print("[CutType] Could not enumerate worksheets:", e)
            return False

        if target_ws is None:
            try:
                present = [w.title for w in ss.worksheets()]
            except Exception:
                present = []
            print("[CutType] Tab 'Production Orders' not found. Tabs present:", present)
            return False

        # 2) Read headers (row 1) and find columns
        header = target_ws.row_values(1)
        header_lc = [h.strip().lower() for h in header]

        # Order column candidates
        order_candidates = ["order #", "order number", "order"]
        order_col_idx = None
        for name in order_candidates:
            if name in header_lc:
                order_col_idx = header_lc.index(name) + 1
                break
        if not order_col_idx:
            print("[CutType] Could not find an 'Order #' column. Header row:", header)
            return False

        # Cut Type column (accept slight variants)
        cut_col_idx = None
        for i, h in enumerate(header_lc, start=1):
            if h in ("cut type", "cut-type", "cuttype"):
                cut_col_idx = i
                break
        if not cut_col_idx:
            print("[CutType] Could not find a 'Cut Type' column. Header row:", header)
            return False

        # 3) Normalize order strings for comparison
        def _norm_order_str(s: str) -> str:
            s = (s or "").strip()
            if s.startswith("$"):
                s = s[1:].strip()
            # Remove spaces/commas commonly present in numeric text
            s = s.replace(",", "").strip()
            # If it looks like a float "172.0", coerce to int-like string
            try:
                if s and all(ch in "0123456789.-" for ch in s):
                    f = float(s)
                    if f.is_integer():
                        s = str(int(f))
                    else:
                        s = str(f)
            except Exception:
                pass
            return s

        ord_norm = _norm_order_str(order_number)

        # 4) Scan the Order column for a match
        col_vals = target_ws.col_values(order_col_idx)
        target_row = None
        for r in range(2, len(col_vals) + 1):
            cell_raw = col_vals[r - 1]
            if _norm_order_str(cell_raw) == ord_norm:
                target_row = r
                break

        if not target_row:
            print(f"[CutType] Order not found (looking for '{ord_norm}') in column {order_col_idx}. "
                  f"First 12 values: {col_vals[1:13]}")
            return False

        # 5) Compute A1 and write with USER_ENTERED
        def _col_to_a1(cidx: int) -> str:
            s = ""
            n = cidx
            while n > 0:
                n, rem = divmod(n - 1, 26)
                s = chr(65 + rem) + s
            return s

        a1 = f"{_col_to_a1(cut_col_idx)}{target_row}"
        rng = f"{target_ws.title}!{a1}"

        # Use the Values API (gspread) to update a single cell
        target_ws.spreadsheet.values_update(
            rng,
            params={"valueInputOption": "USER_ENTERED"},
            body={"values": [[cut_type]]}
        )
        print(f"[CutType] Wrote '{cut_type}' for order {order_number} at {rng}")
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

def shrink_page_and_stamp_horizontal_qrs(original_pdf_path, order_number, spreadsheet, margin_pts=72, stops=None):

    # PIL alias to avoid any local shadowing of the name "Image"
    import PIL.Image as PILImage

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
        img = qr.make_image().convert('1', dither=PILImage.NONE)
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
        img = qr.make_image().convert('1', dither=PILImage.NONE)
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

    # ---------- Draw red triangles next to each selected sequence row ----------
    if stops:
        try:
            stops = [int(s) for s in stops if str(s).isdigit()]
            stops = sorted({s for s in stops if 1 <= s})
        except Exception:
            stops = []

    if stops:
        # Use the exact PDF-space points saved by the UI (and fallback to the same parser the UI uses).
        # These are already in ORIGINAL PDF coordinates (top-down).
        marks = get_stop_mark_points_pdfspace(original_pdf_path, max(stops))

        # Build a LEFT-pointing triangle once.
        # IMPORTANT: We draw it so the TIP is the image's LEFT edge midpoint.
        TRI_W, TRI_H = 12, 12
        tri_path = os.path.join(os.getenv("TEMP") or ".", "jrco_triangle_red.png")
        try:
            from PIL import Image, ImageDraw
            img_tri = Image.new("RGBA", (TRI_W, TRI_H), (0, 0, 0, 0))
            draw = ImageDraw.Draw(img_tri)
            # Left-pointing triangle: tip at (0, TRI_H/2)
            draw.polygon([(0, TRI_H // 2), (TRI_W - 1, 0), (TRI_W - 1, TRI_H - 1)], fill=(200, 0, 0, 255))
            img_tri.save(tri_path)
        except Exception:
            tri_path = None  # still stamp a caret if image build fails

        # compute overlay offset once (top-left coords)
        top_offset = h - (scale * h) - ty
        print(f"[StampDBG] overlay map: scale={scale:.4f} tx={tx:.1f} ty={ty:.1f} h={h:.1f} top_offset={top_offset:.1f}", flush=True)


        for n in stops:
            try:
                if n in marks:
                    x_tip_pts, y_mid_pts = marks[n]  # top-down PDF points used by the UI
                    # Apply the same page transform, but Y must use top_offset (not ty)
                    # x' = x*scale + tx;  y' = y*scale + top_offset   (FPDF uses a top-left origin with y down)
                    x_pdf = (x_tip_pts * scale) + tx
                    y_pdf = (y_mid_pts * scale) + top_offset
                else:
                    # Conservative fallback if a row wasn't found by the locator
                    x_pdf = pad + 40
                    y_pdf = h - bottom_margin - max_h + 24 * n


                if tri_path and os.path.exists(tri_path):
                    # Place the image so its TIP lands exactly at (x_pdf, y_pdf):
                    # FPDF.image uses TOP-LEFT as the anchor; our tip is at the LEFT edge midpoint.
                    pdf.image(tri_path, x=x_pdf, y=y_pdf - (TRI_H / 2), w=TRI_W, h=TRI_H)
                else:
                    # Minimal caret fallback (still tip-aligned)
                    pdf.set_draw_color(200, 0, 0)
                    pdf.set_line_width(1)
                    pdf.line(x_pdf, y_pdf, x_pdf + 6, y_pdf)
                    pdf.line(x_pdf + 3, y_pdf - 6, x_pdf, y_pdf)
                    pdf.line(x_pdf + 3, y_pdf - 6, x_pdf + 6, y_pdf)

            except Exception as e:
                print(f"[Stamp] Triangle draw failed at row {n}: {e}", flush=True)

        # Cleanup
        try:
            if tri_path and os.path.exists(tri_path):
                os.remove(tri_path)
        except Exception:
            pass


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

    # Do NOT delete/move the source here.
    # The file watcher will move the just-processed PDF into a 'processed/' subfolder.
    logging.debug("Leaving source PDF in place for watcher to move: %s", original_pdf_path)

def _locate_stop_sequence_rows(pdf_path: str, max_rows: int = 200):
    """
    Returns:
      positions: {row_number: y_center_pts}   # vertical centers between rows (biased slightly high)
      st_left:   float                        # X (points) of the left edge of 'St' column from header

    Strategy:
      - Find the '# Color St. Code Name' header line.
      - Read the x0 of the 'St' header token => st_left.
      - For each row (1., 2., ...), compute a y_center between rows (biased higher so it doesn't look low).
      - We will draw triangles at x = st_left - margin (between Color and St), y = y_center - nudge.
    """
    import pdfplumber, re

    def _cluster_lines(words, tol=1.8):
        buckets = {}
        for w in words:
            ky = round(w["top"] / tol) * tol
            buckets.setdefault(ky, []).append(w)
        lines = []
        for ky, ws in buckets.items():
            ws = sorted(ws, key=lambda w: w["x0"])
            text = " ".join(w["text"] for w in ws)
            top = min(w["top"] for w in ws)
            bottom = max(w["bottom"] for w in ws)
            x0 = min(w["x0"] for w in ws)
            x1 = max(w["x1"] for w in ws)
            lines.append({"top": top, "bottom": bottom, "x0": x0, "x1": x1, "words": ws, "text": text})
        lines.sort(key=lambda L: L["top"])
        return lines

    def _norm(s: str) -> str:
        return re.sub(r"\s+", " ", (s or "").lower()).strip()

    with pdfplumber.open(pdf_path) as pdf:
        if not pdf.pages:
            return {}, None
        page = pdf.pages[0]
        words = page.extract_words(keep_blank_chars=False, use_text_flow=True, x_tolerance=1.0, y_tolerance=1.0)
        if not words:
            return {}, None

        lines = _cluster_lines(words, tol=1.8)

        # 1) "Stop Sequence:" line (optional step for alignment)
        header_idx = None
        for i, L in enumerate(lines):
            if _norm(L["text"]).startswith("stop sequence:"):
                header_idx = i
                break

        # 2) Column header line: "# Color St. Code Name"
        col_header_idx = header_idx + 1 if header_idx is not None else None
        if col_header_idx is None or col_header_idx >= len(lines):
            for i, L in enumerate(lines):
                t = _norm(L["text"])
                if t.startswith("# color st. code name") or t.startswith("# color st code name"):
                    col_header_idx = i
                    break
        if col_header_idx is None or col_header_idx >= len(lines):
            return {}, None

        col_hdr = lines[col_header_idx]
        hdr_ws = col_hdr["words"]
        hdr_texts = [w["text"].lower().strip(".:") for w in hdr_ws]

        # 2a) Derive st_left from the 'St' header token
        st_left = None
        for idx, t in enumerate(hdr_texts):
            if t in ("st", "st.", "stitch"):
                st_left = float(hdr_ws[idx]["x0"])
                break
        if st_left is None:
            # Fallback: use right edge of header line - better than nothing
            st_left = float(col_hdr["x1"])

        # 3) Collect row lines: lines that start with "1." / "2." ...
        start_i = col_header_idx + 1
        row_lines = []
        for L in lines[start_i:]:
            t = L["text"].strip()
            if re.match(r"^\s*\d+\s*[\.\)]?\b", t):
                row_lines.append(L)
                if len(row_lines) >= max_rows:
                    break
            elif row_lines:
                break
        if not row_lines:
            return {}, st_left

        # 4) For each row, compute a y_center biased higher so it doesn't look low
        positions = {}
        for idx, L in enumerate(row_lines):
            if idx < len(row_lines) - 1:
                gap = row_lines[idx + 1]["top"] - L["bottom"]
                y_center = L["bottom"] + 0.50 * gap  # exact midpoint between rows
            else:
                prev_gap = (row_lines[idx]["top"] - row_lines[idx - 1]["bottom"]) if idx >= 1 else 16.0
                y_center = L["bottom"] + 0.50 * prev_gap

            # Extract row number
            m = re.match(r"^\s*(\d+)", L["text"])
            if not m:
                continue
            n = int(m.group(1))

            positions[n] = float(y_center)

        return positions, float(st_left)

def _persist_ui_marks_to_sidecar(pdf_path: str, seq_positions_px: dict, sx: float, sy: float, ui_tri_size: int = 14) -> None:
    """
    Save the EXACT tip positions the UI used into <pdf>.marks.json.
    - seq_positions_px maps n -> (cx_px, cy_px) where cx_px is CENTER X in pixels (UI),
      so we subtract half the triangle width and convert pixels -> PDF points.
    - sx, sy are your UI scale factors (px per PDF point).
    """
    try:
        import os, json
        marks_pdf: dict[int, list[float]] = {}
        for n, (cx_px, cy_px) in seq_positions_px.items():
            x_tip_pts = (float(cx_px) - (ui_tri_size // 2)) / float(sx)  # back to PDF points (left tip)
            y_pts     = float(cy_px) / float(sy)                         # PDF points (top-down)
            marks_pdf[int(n)] = [x_tip_pts, y_pts]
        sidecar = os.path.splitext(pdf_path)[0] + ".marks.json"
        with open(sidecar, "w", encoding="utf-8") as f:
            json.dump({"pdf_path": pdf_path, "marks": marks_pdf}, f, ensure_ascii=False, indent=2)
            f.flush()
            os.fsync(f.fileno())
        print(f"[UI->Sidecar] wrote {sidecar} with {len(marks_pdf)} marks", flush=True)

    except Exception as e:
        print(f"[UI->Sidecar] could not write sidecar: {e}", flush=True)


def get_stop_mark_points_pdfspace(pdf_path: str, max_seq: int, tip_inset_pts: float = 3.0) -> dict[int, tuple[float, float]]:
    """
    Always return {n: (x_tip_pts, y_pts)} in PDF points (page 1).
    Normalizes across sidecar, st_left+Y-only, and full (x,y) providers.
    """
    import os, json

    # 1) Sidecar from UI (already [x_tip, y]) — safest & preferred
    out: dict[int, tuple[float, float]] = {}
    try:
        sidecar = os.path.splitext(pdf_path)[0] + ".marks.json"
        with open(sidecar, "r", encoding="utf-8") as f:
            j = json.load(f)
            marks = j.get("marks") or {}
            for k, v in marks.items():
                try:
                    n = int(k)
                except Exception:
                    continue
                if isinstance(v, (list, tuple)) and len(v) == 2:
                    out[n] = (float(v[0]), float(v[1]))
    except Exception:
        pass

    if out:
        # Already normalized to (x_tip, y)
        return out

    # 2) No sidecar — detect from page
    try:
        # A: may return (positions_y_only, st_left)
        pos_a, st_left = _locate_stop_sequence_rows(pdf_path, max_rows=max(12, int(max_seq or 0)))
    except Exception:
        pos_a, st_left = {}, None

    # B: returns full tuples {n: (x_tip, y)}
    pos_b = _locate_stop_mark_positions(pdf_path, max_seq=max(12, int(max_seq or 0)))

    # Choose best: more rows wins; tie → prefer A (when st_left present)
    positions = pos_b if len(pos_b) > len(pos_a) else pos_a

    # 3) Normalize
    norm: dict[int, tuple[float, float]] = {}
    if positions and st_left is not None:
        # positions is Y-only; synthesize x from st_left - inset
        x_tip = float(st_left) - float(tip_inset_pts)
        # If max_seq looks small, still iterate over keys we actually have
        keys = sorted(set(positions.keys()) | set(range(1, int(max_seq or 0) + 1)))
        for n in keys:
            y = positions.get(n)
            if y is None:
                continue
            try:
                n_int = int(n)
            except Exception:
                continue
            norm[n_int] = (x_tip, float(y))
        return norm

    # Otherwise positions should already be {n: (x, y)} — but be defensive:
    for k, v in (positions or {}).items():
        try:
            n = int(k)
        except Exception:
            continue
        if isinstance(v, (list, tuple)) and len(v) == 2:
            norm[n] = (float(v[0]), float(v[1]))
        else:
            # Rare case: Y-only arrived with no st_left — skip rather than breaking
            continue

    return norm


def _atomic_replace(src_tmp: str, dst_final: str, retries: int = 25, backoff: float = 0.2) -> None:
    """
    Atomically move/replace src_tmp -> dst_final with retries to survive Windows locks
    (Explorer preview, antivirus, sync daemons). Backoff increases slightly each try.
    """
    import errno
    import ctypes, time

    # Make sure parent dir exists
    os.makedirs(os.path.dirname(dst_final), exist_ok=True)

    last_err = None
    for i in range(max(1, int(retries))):
        try:
            # On Windows, os.replace() does an atomic rename/overwrite if possible
            os.replace(src_tmp, dst_final)
            return
        except PermissionError as e:
            last_err = e
        except OSError as e:
            last_err = e
            # Some AV/sync tools report generic access denied or sharing violation
            if getattr(e, "errno", None) not in (errno.EACCES, errno.EPERM, errno.ETXTBSY):
                # If it’s not a typical lock, don’t spin forever
                pass

        # Nudge Explorer to release handles by pinging the shell (best-effort, ignore errors)
        try:
            ctypes.windll.shell32.SHChangeNotify(0x08000000, 0x0000, None, None)  # SHCNE_ASSOCCHANGED
        except Exception:
            pass

        # Backoff and retry
        time.sleep(backoff + i * 0.05)

    raise RuntimeError(f"_atomic_replace failed for {dst_final}: {last_err}")


def stamp_triangles_only(original_pdf_path: str, order_number: str, stops: list[int],
                         tri_size_pts: int = 14, inset_tip_pts: float = 3.0) -> str:
    """
    Draw ONLY red triangles on page 1:
      • one per selected stop number
      • LEFT TIP sits just LEFT of the 'St.' column (inside Stop Sequence)
      • Y is BETWEEN row n and row n+1 (“After n”)
      • points LEFT; size = tri_size_pts
    Output → <OUTPUT_FOLDER>/<order>_Stamped.pdf (temp + atomic replace)
    """
    import os, time
    out_path = os.path.join(OUTPUT_FOLDER, f"{order_number}_Stamped.pdf")
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    tmp_suffix = f".tri.{int(time.time() * 1000)}.tmp.pdf"
    out_tmp = out_path + tmp_suffix

    try:
        import fitz  # PyMuPDF

        # If no stops, copy-through
        if not stops:
            with open(original_pdf_path, "rb") as src, open(out_tmp, "wb") as dst:
                dst.write(src.read())
            _atomic_replace(out_tmp, out_path, retries=25, backoff=0.2)
            return out_path

        # Try to use the exact UI positions from a sidecar JSON (if present)
        import os, json
        marks_from_ui: dict[int, tuple[float, float]] = {}

        def _try_load_sidecar(path: str) -> dict[int, tuple[float, float]]:
            try:
                if path and os.path.exists(path):
                    with open(path, "r", encoding="utf-8") as f:
                        data = json.load(f)
                    return {
                        int(k): (float(v[0]), float(v[1]))
                        for k, v in (data.get("marks") or {}).items()
                        if isinstance(v, (list, tuple)) and len(v) >= 2
                    }
            except Exception as e:
                print(f"[TrianglesOnly] Sidecar read failed ({path}): {e}", flush=True)
            return {}

        base_no_ext = os.path.splitext(original_pdf_path)[0]
        order_str   = str(order_number).strip()

        # Primary: same folder
        candidates = [base_no_ext + ".marks.json"]

        # Also check common locations by filename / order number
        try:
            stops_queue_dir     = r"G:\My Drive\Stops Queue"
            production_sheets   = r"G:\My Drive\Production Sheets"
            stops_cache_dir     = r"G:\My Drive\Stops Cache"
            base_name           = os.path.basename(base_no_ext)  # e.g., "116"

            candidates += [
                os.path.join(stops_queue_dir,     f"{base_name}.marks.json"),
                os.path.join(stops_queue_dir,     f"{order_str}.marks.json"),
                os.path.join(production_sheets,   f"{base_name}.marks.json"),
                os.path.join(production_sheets,   f"{order_str}.marks.json"),
                os.path.join(stops_cache_dir,     f"{base_name}.marks.json"),
                os.path.join(stops_cache_dir,     f"{order_str}.marks.json"),
            ]
        except Exception:
            pass

        marks_from_ui = {}
        for attempt in range(10):  # ~1.0s total max
            for c in candidates:
                marks_from_ui = _try_load_sidecar(c)
                if marks_from_ui:
                    print(f"[TrianglesOnly] Using UI sidecar marks from: {c} ({len(marks_from_ui)} entries)", flush=True)
                    break
            if marks_from_ui:
                break
            time.sleep(0.1)

        if marks_from_ui:
            # Build positions (Y in PDF points, top-down) from the sidecar
            positions = {n: float(y) for n, (_x, y) in marks_from_ui.items()}
            st_left = None  # not used when sidecar gives us X
        else:
            positions, st_left = _locate_stop_sequence_rows(original_pdf_path, max_rows=max(stops or [1]))
            print(f"[TrianglesOnly] Locator: rows={len(positions)}, st_left={st_left}", flush=True)

        if not positions or (st_left is None and not marks_from_ui):


                with open(original_pdf_path, "rb") as src, open(out_tmp, "wb") as dst:
                    dst.write(src.read())
                _atomic_replace(out_tmp, out_path)
                return out_path


        doc = fitz.open(original_pdf_path)
        if not len(doc):
            raise RuntimeError("Empty PDF")
        page = doc[0]

        red  = (200/255.0, 0.0, 0.0)
        half = float(tri_size_pts) / 2.0
        uniq = sorted({int(s) for s in stops if str(s).isdigit()})

        page_height = float(page.rect.height)

        # Convert UI (top-down) Y into PyMuPDF (bottom-up) Y safely.
        page_height = float(page.rect.height)

        # Build a safe Y window for the Stop Sequence block:
        # positions contains BETWEEN-row y's in top-down coords. Convert bounds to bottom-up, too.
        if positions:
            y_td_vals = [float(v) for v in positions.values()]
            y_td_min, y_td_max = min(y_td_vals), max(y_td_vals)
            # We'll compute the bottom-up band *after* we pick the final page size below.
        else:
            y_td_min = 0.0
            y_td_max = 0.0

        MARGIN = 6.0  # pts — small cushion so we don't reject valid edges

        # Align pdfplumber (top-down) coordinate space with PyMuPDF
        try:
            import pdfplumber
            with pdfplumber.open(original_pdf_path) as _p:
                plumber_w = float(_p.pages[0].width)
                plumber_h = float(_p.pages[0].height)
        except Exception:
            # Fall back to PyMuPDF sizes if we can't read plumber
            plumber_w = float(page.rect.width)
            plumber_h = float(page.rect.height)

        # Pick the PyMuPDF size (rect vs bound) that best matches pdfplumber’s page
        rect_w,  rect_h  = float(page.rect.width),  float(page.rect.height)
        bound_w, bound_h = float(page.bound().width), float(page.bound().height)

        # Choose whichever is closer to pdfplumber's size (handles rotation / cropbox)
        if plumber_w and plumber_h:
            if abs(bound_w - plumber_w) + abs(bound_h - plumber_h) < abs(rect_w - plumber_w) + abs(rect_h - plumber_h):
                page_width_use, page_height_use = bound_w, bound_h
            else:
                page_width_use, page_height_use = rect_w, rect_h
        else:
            page_width_use, page_height_use = rect_w, rect_h

        # Compute scale from pdfplumber space -> PyMuPDF space
        scale_x = (page_width_use  / plumber_w) if plumber_w else 1.0
        scale_y = (page_height_use / plumber_h) if plumber_h else 1.0

        # Now that we’ve chosen the page size, compute the expected vertical band (bottom-up) for sanity checks
        # Compute the expected bottom-up band AFTER picking the best page height
        # We align pdfplumber (top-down) Y scale with PyMuPDF page height to avoid drift.
        # 1) Get pdfplumber page height (top-down space)
        try:
            import pdfplumber
            with pdfplumber.open(original_pdf_path) as _p:
                plumber_h = float(_p.pages[0].height)
        except Exception:
            plumber_h = float(page.rect.height)

        # 2) Choose the PyMuPDF height that best matches pdfplumber’s notion of the page
        rect_h  = float(page.rect.height)
        bound_h = float(page.bound().height)
        if plumber_h and abs(bound_h - plumber_h) < abs(rect_h - plumber_h):
            page_height_use = bound_h
        else:
            page_height_use = rect_h

        # 3) Scale factor from pdfplumber (top-down) -> PyMuPDF height
        scale_y = (page_height_use / plumber_h) if plumber_h else 1.0

        # 4) Build a safe vertical band (in bottom-up coords) using the chosen height + scale
        if positions:
            y_td_vals = [float(v) for v in positions.values()]
            y_td_min, y_td_max = min(y_td_vals), max(y_td_vals)
            y_fit_min = page_height_use - (y_td_max * scale_y)
            y_fit_max = page_height_use - (y_td_min * scale_y)
        else:
            y_fit_min = 0.0
            y_fit_max = page_height_use

        MARGIN = 6.0  # pts — small cushion so we don't reject valid edges

        # --- ensure we have chosen page size + scales (if you already set these above, this keeps them) ---
        try:
            _ = page_width_use  # check if defined
            _ = page_height_use
        except NameError:
            # Pick the PyMuPDF size that best matches pdfplumber’s page
            rect_w,  rect_h  = float(page.rect.width),  float(page.rect.height)
            bound_w, bound_h = float(page.bound().width), float(page.bound().height)
            try:
                import pdfplumber
                with pdfplumber.open(original_pdf_path) as _p:
                    plumber_w = float(_p.pages[0].width)
                    plumber_h = float(_p.pages[0].height)
            except Exception:
                plumber_w = rect_w
                plumber_h = rect_h

            if abs(bound_w - plumber_w) + abs(bound_h - plumber_h) < abs(rect_w - plumber_w) + abs(rect_h - plumber_h):
                page_width_use, page_height_use = bound_w, bound_h
            else:
                page_width_use, page_height_use = rect_w, rect_h

            # If you already had scale_y above, we won't overwrite it; otherwise set both:
            if 'scale_y' not in locals():
                scale_y = (page_height_use / plumber_h) if plumber_h else 1.0
            if 'scale_x' not in locals():
                scale_x = (page_width_use / plumber_w)  if plumber_w  else 1.0

        # If scale_x wasn't defined earlier, define it now based on chosen sizes
        if 'scale_x' not in locals():
            try:
                import pdfplumber
                with pdfplumber.open(original_pdf_path) as _p:
                    plumber_w = float(_p.pages[0].width)
            except Exception:
                plumber_w = float(page.rect.width)
            scale_x = (page_width_use / plumber_w) if plumber_w else 1.0

        # --- DEBUG: report mapping once ---
        try:
            rect_w,  rect_h  = float(page.rect.width),  float(page.rect.height)
            bound_w, bound_h = float(page.bound().width), float(page.bound().height)
            print(
                "[StampDBG] sizes: "
                f"chosen=({page_width_use:.2f}x{page_height_use:.2f}) "
                f"rect=({rect_w:.2f}x{rect_h:.2f}) "
                f"bound=({bound_w:.2f}x{bound_h:.2f}) "
                f"scale_x={scale_x:.6f} scale_y={scale_y:.6f} "
                f"st_left={float(st_left):.2f} inset={float(inset_tip_pts):.2f}",
                flush=True
            )
        except Exception:
            pass

        for n in uniq:
            y_mid_td = positions.get(n)
            if y_mid_td is None:
                continue

            if marks_from_ui and (n in marks_from_ui):
                # ✅ UI sidecar already stored points in the SAME space PyMuPDF draws in:
                #    top-left origin, Y increases downward. So DO NOT scale or flip.
                x_tip_fit = float(marks_from_ui[n][0])   # left-tip X in PDF points
                y_mid_fit = float(marks_from_ui[n][1])   # mid-line Y in PDF points
            else:
                # Fallback for when there is no UI sidecar: convert pdfplumber (top-down)
                # into PyMuPDF drawing coords using the chosen page size and scale.
                y_mid_fit = page_height_use - (scale_y * float(y_mid_td))
                if not (y_fit_min - MARGIN <= y_mid_fit <= y_fit_max + MARGIN):
                    y_mid_fit = page_height_use - float(y_mid_td)

                x_tip_td  = float(st_left) - float(inset_tip_pts)
                x_tip_fit = (scale_x * x_tip_td)

            # Draw a LEFT-pointing triangle whose TIP is exactly at (x_tip_fit, y_mid_fit)
            x_center = x_tip_fit + half
            pts = [
                (x_center - half, y_mid_fit),        # TIP (left)
                (x_center + half, y_mid_fit - half), # upper base
                (x_center + half, y_mid_fit + half), # lower base
            ]

            try:
                page.draw_polygon(pts, color=None, fill=red, close=True)
            except AttributeError:
                shp = page.new_shape()
                shp.draw_polyline(pts)
                shp.finish(fill=red, color=None)
                shp.commit()




        doc.save(out_tmp)
        doc.close()
        _atomic_replace(out_tmp, out_path, retries=25, backoff=0.2)
        return out_path

    except Exception as e:
        print(f"[TrianglesOnly] Failed: {e}")
        try:
            with open(original_pdf_path, "rb") as src, open(out_tmp, "wb") as dst:
                dst.write(src.read())
            _atomic_replace(out_tmp, out_path, retries=25, backoff=0.2)
        except Exception:
            pass
        return out_path


# ================================
# SECTION 8: File Monitoring
# ================================
processed = set()
RECENTLY_HANDLED = {}     # {filepath: last_ts}
DEBOUNCE_SECONDS = 1.0

class StopsQueueHandler(FileSystemEventHandler):
    """Watches G:\\My Drive\\Stops Queue -> show visual picker -> stamp triangles -> copy to Production."""

    def _handle_new_pdf(self, p: str):
        try:
            if not p or not p.lower().endswith(".pdf"):
                return
            # Ensure it’s a file and fully written
            if not os.path.isfile(p):
                return
            if not wait_for_stable_file(p):
                return

            base  = os.path.basename(p)
            order = clean_value(os.path.splitext(base)[0])
            print(f"[StopsQueue] New PDF in Stops Queue: {base} (order {order})")

            # 🔕 Debounce duplicate FS events for the same file
            import time as pytime  # alias avoids any local 'time' name shadowing
            now_ts = pytime.time()
            last_ts = RECENTLY_HANDLED.get(p, 0)
            if (now_ts - last_ts) < DEBOUNCE_SECONDS:
                print(f"[StopsQueue] Debounced duplicate FS event for {base}")
                return
            RECENTLY_HANDLED[p] = now_ts

            # Determine sequence count
            # Determine sequence count from the Stop Sequence (never cap by thread usage)
            try:
                guessed = _detect_max_sequence(p)  # anchored to "# Color St. Code Name"
            except Exception:
                guessed = 0

            # Try both detectors with a generous cap
            try:
                pos_a, _ = _locate_stop_sequence_rows(p, max_rows=max(guessed, 24))
            except Exception:
                pos_a, _ = {}, None

            # ✅ Prefer the header-anchored detector (A) when it finds at least a few rows
            pos_b = _locate_stop_mark_positions(p, max_seq=max(guessed, 24))
            positions = pos_a if len(pos_a) >= 3 else (pos_b if len(pos_b) >= 1 else {})

            # ✅ Honor what we actually found (no artificial 12-row floor)
            seq_max = max(len(pos_a), len(pos_b), int(guessed or 0), 1)
            seq_max = min(seq_max, 150)

            print(f"[StopsQueue] Opening stop selection UI for {order}...", flush=True)
            stops = prompt_stops_interactive(p, seq_max)
            # normalize to a clean, sorted int list within [1..seq_max]
            stops = sorted({int(s) for s in (stops or []) if str(s).isdigit() and 1 <= int(s) <= seq_max})
            print(f"[StopsQueue] Stops chosen: {stops}")


            if (not stops) and seq_max >= 1:
                stops = [1]
                print("[StopsQueue] No stops selected; defaulting to [1].")

            # Stamp triangles on the original, then copy to Production Sheets
            stamped_path = stamp_triangles_only(
                original_pdf_path=p,
                order_number=order,
                stops=stops,
                tri_size_pts=14,
            )

            # NEW: open the stamped PDF for visual confirmation (non-blocking)
            try:
                os.startfile(stamped_path)
            except Exception as e:
                print(f"[StopsQueue] Note: could not auto-open stamped PDF: {e}")

            # Copy/overwrite into Production Sheets; tolerate brief sync locks
            prod_pdf = fr"G:\My Drive\Production Sheets\{order}.pdf"
            os.makedirs(os.path.dirname(prod_pdf), exist_ok=True)

            import time
            for i in range(6):  # ~3s total
                try:
                    shutil.copyfile(stamped_path, prod_pdf)
                    break
                except PermissionError as e:
                    time.sleep(0.5)
                except Exception as e:
                    print(f"[StopsQueue] Copy failed (attempt {i+1}): {e}")
                    time.sleep(0.5)
            print(f"[StopsQueue] Copied TRIANGLED PDF -> {prod_pdf} (Drive will shrink this as a whole)")

            # NEW: also copy the UI sidecar next to the Production Sheets PDF (if it exists)
            try:
                sidecar_src = os.path.splitext(p)[0] + ".marks.json"   # p == source in Stops Queue
                sidecar_dst = os.path.splitext(prod_pdf)[0] + ".marks.json"
                if os.path.exists(sidecar_src):
                    shutil.copyfile(sidecar_src, sidecar_dst)
                    print(f"[StopsQueue] Copied sidecar -> {sidecar_dst}")
            except Exception as e:
                print(f"[StopsQueue] Note: could not copy sidecar: {e}")


            # Optional: move source into Stops Queue\processed; tolerate brief locks
            try:
                processed_dir = os.path.join(os.path.dirname(p), "processed")
                os.makedirs(processed_dir, exist_ok=True)
                for i in range(6):
                    try:
                        shutil.move(p, os.path.join(processed_dir, base))
                        break
                    except PermissionError:
                        time.sleep(0.5)
                else:
                    print("[StopsQueue] Note: could not move to processed/ after retries; leaving in place.")
            except Exception as e:
                print(f"[StopsQueue] Note: could not move source to processed/: {e}")

            print("[StopsQueue] ✅ Done handling stops & copy-back.")


        except Exception as e:
            print(f"[StopsQueue] Error handling {p}: {e}")

    def on_created(self, event):
        if event.is_directory:
            return
        self._handle_new_pdf(event.src_path)

    def on_moved(self, event):
        # Handles moves/renames into Stops Queue (common on Windows/Drive)
        if event.is_directory:
            return
        dest = getattr(event, "dest_path", None) or getattr(event, "dest_path", "")
        if dest:
            self._handle_new_pdf(dest)

    def on_modified(self, event):
        # Some setups emit modified (not created) for new files; harmless to handle
        if event.is_directory:
            return
        self._handle_new_pdf(event.src_path)

class ProdToStopsInterceptor(FileSystemEventHandler):
    """On creation in Production Sheets, immediately move to Stops Queue for stop selection."""
    def on_created(self, event):
        if event.is_directory or not event.src_path.lower().endswith(".pdf"):
            return
        p = event.src_path
        if not wait_for_stable_file(p):
            return

        # Skip if this file was just copied back from Stops Queue
        try:
            import time, os
            ap = os.path.abspath(event.src_path)
            ts = _RECENT_RETURNED_TO_PROD.get(ap)
            if ts and (time.time() - ts) < 120:
                print(f"[Intercept] Skipping (just returned from Stops Queue): {os.path.basename(ap)}")
                # Clean up the sentinel so future real creations will be seen
                _RECENT_RETURNED_TO_PROD.pop(ap, None)
                return
        except Exception:
            pass

        try:
            base = os.path.basename(p)
            order = os.path.splitext(base)[0]
            src_dir = os.path.dirname(p)
            if os.path.basename(src_dir).lower() != "production sheets":
                return  # only intercept the Production Sheets directory
            stops_dir = r"G:\My Drive\Stops Queue"
            os.makedirs(stops_dir, exist_ok=True)
            dst = os.path.join(stops_dir, base)
            print(f"[Intercept] Moving {base} -> Stops Queue")
            shutil.move(p, dst)
        except Exception as e:
            print(f"[Intercept] Failed to move to Stops Queue: {e}")



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
        print(f"[Stops] Detected PDF for order {order} at {p}", flush=True)
        usage = extract_thread_usage(p)
        if not usage:
            print(f"[Stops] No usage parsed from {p}; skipping stops dialog.", flush=True)
            return

        # DEBUG - prove we are here and can show UI
        _debug_msgbox("JR & Co - Stops checkpoint", f"Detected usage rows: {len(usage)} for order {order}")

        # Prompt for stops (1..N)
        try:
            print("[Stops] Opening interactive PDF stops window...", flush=True)
            stops = prompt_stops_interactive(p, len(usage))
            print(f"[Stops] Picked: {stops}", flush=True)
        except Exception as e:
            print(f"[Stops] Prompt failed ({e}); proceeding with no stops.", flush=True)
            stops = []

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

        # Stamp PDF and add QR codes (with stops)
        print(f"[Stops] Stamping with stops: {stops}", flush=True)
        shrink_page_and_stamp_horizontal_qrs(p, order, sheet_spreadsheet, stops=stops)

        # MOVE the just-processed PDF into "processed/" so Watchdog won't see it again
        import shutil, os
        processed_dir = os.path.join(os.path.dirname(event.src_path), 'processed')
        os.makedirs(processed_dir, exist_ok=True)
        shutil.move(event.src_path, os.path.join(processed_dir, os.path.basename(event.src_path)))


from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

def monitor_drive_folder(folder_id, sh, ss):
    global sheet_thread, sheet_spreadsheet
    sheet_thread, sheet_spreadsheet = sh, ss

    creds_path = r"C:\Users\Embroidery1\Desktop\OrderEntry,Inventory,QR,PrintPDF - Copy\Keys\poetic-logic-454717-h2-3dd1bedb673d.json"
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
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    while True:
        try:
            results = service.files().list(
                q=f"'{folder_id}' in parents and mimeType='application/pdf'",
                spaces='drive',
                fields='files(id, name, modifiedTime)',
                orderBy='modifiedTime desc',
                supportsAllDrives=True
            ).execute()
            items = results.get('files', [])

            for file in items:
                file_id = file['id']
                name = file['name']

                if file_id in processed_ids:
                    continue
                if name.endswith('_Stamped.pdf') or name.startswith('~$') or name.startswith('.') or name.startswith('_'):
                    continue  # Skip already processed or temp files

                print(f"New PDF found on Drive: {name}")
                request = service.files().get_media(fileId=file_id)
                local_path = os.path.join("temp_downloads", name)

                with open(local_path, 'wb') as f:
                    downloader = MediaIoBaseDownload(f, request)
                    done = False
                    while not done:
                        _, done = downloader.next_chunk()

                order = clean_value(os.path.splitext(name)[0])

                # Thread usage - sheet update (unchanged)
                usage = extract_thread_usage(local_path)
                if not usage:
                    print(f"No thread usage found in {name}, skipping.")
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

                # Single pass: shrink + other overlay (triangles were already applied pre-shrink)
                print("[Stops] Stamping overlay on already-triangled PDF...", flush=True)
                shrink_page_and_stamp_horizontal_qrs(local_path, order, sheet_spreadsheet, stops=None)

                stamped_path = os.path.join(
                    OUTPUT_FOLDER, os.path.splitext(name)[0] + '_Stamped.pdf'
                )

                # Ensure stamped exists
                if not os.path.exists(stamped_path):
                    time.sleep(0.5)
                    if not os.path.exists(stamped_path):
                        print(f"Stamped file not found yet for {name}. Skipping this cycle.")
                        continue

                # Upload to Stamped PS
                file_metadata = {'name': os.path.basename(stamped_path), 'parents': [stamped_folder_id]}
                media = MediaFileUpload(stamped_path, mimetype='application/pdf')
                uploaded = service.files().create(
                    body=file_metadata,
                    media_body=media,
                    fields='id',
                    supportsAllDrives=True
                ).execute()
                print(f"Uploaded stamped version to Stamped PS: {uploaded['id']}")

                # Simulate printing
                print("Simulating print... (placeholder)")
                time.sleep(2)
                print("Dummy print complete.")

                # Move to Printed PS
                service.files().update(
                    fileId=uploaded['id'],
                    addParents=printed_folder_id,
                    removeParents=stamped_folder_id,
                    supportsAllDrives=True,
                    fields='id, parents'
                ).execute()
                print("Moved file to Printed PS folder.")

                # File into Order subfolder
                order_folder_name = os.path.splitext(name)[0]
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
                    print(f"Filed into Order subfolder: {order_folder_name}")
                else:
                    print(f"No subfolder found for order '{order_folder_name}' in Orders folder.")

                # Move original to Archive
                service.files().update(
                    fileId=file_id,
                    addParents=archive_folder_id,
                    removeParents=folder_id,
                    supportsAllDrives=True,
                    fields='id, parents'
                ).execute()
                print(f"Moved original to Archived PS: {file_id}")

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
    loading_win, loading_pb, set_progress = _open_loading_modal(root, total_count=len(files), text="Loading previews...")
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

    loading_win, loading_pb, set_progress = _open_loading_modal(root, total_count=len(files), text="Loading previews...")

    try:
        # Let the reorder dialog update the progress as it builds rows
        files = _reorder_files_dialog_with_root(root, files, on_progress=set_progress, loading_win=loading_win)
    finally:
        _close_loading_modal(loading_win, loading_pb)

    if not files:
        # (User canceled in reorder dialog - keep earlier guard behavior if you like)
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

        # Files we just copied back to Production (skip intercept for ~2 minutes)
        _RECENT_RETURNED_TO_PROD: dict[str, float] = {}

        # Open progress modal on the Tk thread
        run_on_tk_thread(lambda root: _progress_modal_open(root, title="Building PDF...", text="Preparing... 0%"))

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

    # Lower JPEG quality to shrink pages. Range 0-12; 6-8 is a sweet spot.
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
            run_on_tk_thread(lambda root, i=idx-1: _progress_modal_update(root, _pct_export(i, total_pages), f"Exporting... {i}/{total_pages}"))
            _export_single_page_pdf(ps, p, tmp)
            temp_pdfs.append(tmp)
            run_on_tk_thread(lambda root, i=idx: _progress_modal_update(root, _pct_export(i, total_pages), f"Exporting... {i}/{total_pages}"))

        # --- MERGE ---
        print(f"[PDF-Pres] Merging {len(temp_pdfs)} pages...")
        run_on_tk_thread(lambda root: _progress_modal_update(root, 90.0, "Merging pages..."))
        merger = PdfMerger()
        for pdf in temp_pdfs:
            merger.append(str(pdf))
        with open(out_pdf_raw, "wb") as f_out:
            merger.write(f_out)
        merger.close()

        # --- COMPRESS (Ghostscript) ---
        # If you keep Ghostscript:
        run_on_tk_thread(lambda root: _progress_modal_update(root, 95.0, "Compressing..."))
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
        "ebook": "/ebook",      # ~150 dpi <- good default
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
        # Re-enable scanner/clipboard when we're done
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

def open_url_in_browser(url: str, force_activate: bool = False):
    """Open URL in a NEW Chrome window on secondary monitor WITHOUT interfering with scanner."""

    import subprocess, time, threading
    import win32gui, win32con, win32process

    try:
        # --- Dedup ---
        now = time.time()
        last = _last_opened_at.get(url, 0.0)
        if now - last < DEDUP_WINDOW_SEC:
            return
        _last_opened_at[url] = now

        # --- Remember currently focused window ---
        prev_hwnd = win32gui.GetForegroundWindow()

        chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

        chrome_args = [
            chrome_path,
            f"--user-data-dir={CHROME_PROFILE_DIR}",
            "--new-window",
            "--no-first-run",
            "--disable-popup-blocking",
            "--disable-background-mode",
            "--disable-features=TranslateUI",
            url,
        ]

        print(f"[open_url] Launching Chrome (force_activate={force_activate})", flush=True)

        proc = subprocess.Popen(
            chrome_args,
            creationflags=subprocess.CREATE_NEW_PROCESS_GROUP
        )

        chrome_pid = proc.pid
        print(f"[open_url] Chrome PID={chrome_pid}", flush=True)

        # --- Move + RESIZE ONLY the new Chrome window ---
        def _move_only_new_chrome():
            time.sleep(1.2)  # allow window to be created

            def enum_cb(hwnd, _):
                try:
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    if pid != chrome_pid:
                        return

                    # Restore in case it opens minimized/maximized
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

                    # HARD clamp to secondary monitor bounds
                    win32gui.SetWindowPos(
                        hwnd,
                        None,
                        SECONDARY_X,
                        SECONDARY_Y,
                        SECONDARY_W,
                        SECONDARY_H,
                        win32con.SWP_NOZORDER | win32con.SWP_NOACTIVATE
                    )

                except Exception:
                    pass

            win32gui.EnumWindows(enum_cb, None)

        threading.Thread(target=_move_only_new_chrome, daemon=True).start()

        # --- Restore focus to original window ---
        time.sleep(0.3)
        try:
            win32gui.SetForegroundWindow(prev_hwnd)
        except Exception:
            pass

    except Exception as e:
        print(f"[open_url] FAILED: {e}", flush=True)




def restore_scanner():
    global SCANNER_ENABLED
    SCANNER_ENABLED = True
    print("[open_url_in_browser] Scanner re-enabled", flush=True)






def _open_success_page(order: str, dept: str | None = None):
    """Open local green page with order image + details."""
    url = f"http://127.0.0.1:5001/success?order={urllib.parse.quote(str(order))}"
    if dept:
        url += f"&dept={urllib.parse.quote(str(dept))}"
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

    # invalid QR - red page
    if action == "UNKNOWN":
        open_url_in_browser("http://127.0.0.1:5001/fail?msg=Invalid%20QR")
        return


    # missing order - red page
    if not order:
        open_url_in_browser("http://127.0.0.1:5001/fail?msg=Missing%20order%20%23")
        return

    # SHIPPING: open URL already done; show green page for 7s
    if action == "SHIP":
        _open_success_page(order, dept="Shipping")
        return

    # DEPARTMENT: write Quantity Made, then green/red page
    if action == "DEPT":
        ok = update_tab_quantity_made(order, tab)
        if ok:
            _open_success_page(order, dept=tab)
        else:
            open_url_in_browser(
                "http://127.0.0.1:5001/fail?msg=" + urllib.parse.quote(f"{tab} update failed")
            )
        return

# --- Keyboard-wedge listener (works if scanner types characters) ---
# --- Keyboard-wedge listener (works if scanner types characters) ---
def _scanner_keyboard_hook(e):
    """
    Scanner-only keystroke capture.
    - Allows keyboard events to pass through so the user's keyboard is never locked.
    - Works with scanners that do NOT send Enter
    """
    global _accumulated, _last_time, _scan_token

    if e.event_type != "down":
        return None

    now = time.time()
    delta = now - _last_time

    # reset buffer if human pause
    if delta > TIME_THRESHOLD:
        _accumulated = ""

    _last_time = now
    key = e.name

    mapping = {
        "space": " ", "slash": "/", "dot": ".", "question": "?",
        "ampersand": "&", "equal": "=", "minus": "-", "colon": ":",
        "underscore": "_", "backslash": "\\"
    }

    if key == "enter":
        payload = _accumulated
        _accumulated = ""
        m = _URL_RE.search(payload)
        if m:
            print("[SCAN] QR detected", flush=True)
            _handle_scanned_url(m.group(1))
        # Allow Enter to pass through so the physical keyboard remains usable
        return None

    if len(key) == 1:
        _accumulated += key
    elif key in mapping:
        _accumulated += mapping[key]

    # no-enter scanners: detect URL + pause
    if "http" in _accumulated.lower():
        token_snapshot = _scan_token
        start_time = _last_time
        snapshot = _accumulated

        def _delayed_process():
            global _accumulated
            time.sleep(TIME_THRESHOLD * 3)
            if _last_time == start_time and token_snapshot == _scan_token:
                m = _URL_RE.search(snapshot)
                if m:
                    print("[SCAN] QR detected (pause)", flush=True)
                    _handle_scanned_url(m.group(1))
                _accumulated = ""

        threading.Thread(target=_delayed_process, daemon=True).start()

    return None  # allow typing (do not block the keyboard)

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

            # At this point it's a valid JRCO QR link -> handle it
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
    Read the active EmbroideryStudio window by regex (not the foreground Tk window),
    then extract the order token from the first [ ... ] group that looks like a filename/number.
    """
    try:
        from pywinauto import Desktop
        desk = Desktop(backend="uia")
        win = desk.window(title_re=EMB_WINDOW_TITLE)  # e.g., EmbroideryStudio 2025 ...
        title = (win.window_text() or "").strip()
    except Exception:
        return ""

    # Examples:
    #  EmbroideryStudio 2025 - Designing - [172      Tajima TBF] - [eckardjustin@gmail.com]
    blocks = re.findall(r"\[([^\]]+)\]", title) or []
    for b in blocks:
        s = b.strip()
        tok = re.split(r"\s{2,}| - ", s)[0].strip()      # "172      Tajima TBF" -> "172"
        tok = re.sub(r"\.emb$", "", tok, flags=re.I)
        if tok and (tok[0] == '$' or re.search(r"\d", tok)):
            return tok.lstrip("$")
    return ""



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
    """
    Ctrl+Alt+P:
      1) Show Cut Type picker and write to sheet (Ctrl+Alt+T flow)
      2) Do NOT wait for any PDF here. Let the local watchers (below) handle:
         Prep -> Production (intercept) -> Stops Queue (visual UI) -> stamp -> back to Production.
    """
    print("[Control+Alt+P] Triggered.")
    on_control_alt_t()

    # Best-effort log of the order we'll be working on (optional)
    job = get_active_job_number()
    if job:
        print(f"[Ctrl+Alt+P] Using order: {job}")
    else:
        print("[Ctrl+Alt+P] No job detected (will proceed when PDF appears via watchers).")





if __name__ == "__main__":

    # detect primary + secondary monitor coordinates once at startup
    # detect primary + secondary monitor coordinates once at startup
    try:
        from screeninfo import get_monitors
        mons = get_monitors()

        primary = next((m for m in mons if m.is_primary), None)
        secondary = next((m for m in mons if not m.is_primary), None)

        if primary:
            PRIMARY_X = primary.x
            PRIMARY_Y = primary.y
            PRIMARY_W = primary.width
            PRIMARY_H = primary.height

        if secondary:
            SECONDARY_X = secondary.x
            SECONDARY_Y = secondary.y
            SECONDARY_W = secondary.width
            SECONDARY_H = secondary.height

        print(
            f"[Monitors] PRIMARY=({PRIMARY_X},{PRIMARY_Y},{PRIMARY_W}x{PRIMARY_H}), "
            f"SECONDARY=({SECONDARY_X},{SECONDARY_Y},{SECONDARY_W}x{SECONDARY_H})"
        )
    except Exception as e:
        print("[Monitor detection failed]", e)

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
    # DISABLED: PDF monitoring and stamping moved to another computer
    # print(f"Monitoring Google Drive folder: Production Sheets ({GOOGLE_DRIVE_FOLDER_ID})")
    # folder_thread = threading.Thread(
    #     target=monitor_drive_folder,
    #     args=(GOOGLE_DRIVE_FOLDER_ID, sheet_thread, sheet_spreadsheet),
    #     daemon=True
    # )
    # folder_thread.start()
    # print("[Startup] Drive monitor thread started.")
    print("[Startup] Drive monitor thread DISABLED (running on another computer).")

    # 4.5) Local filesystem watchers
    try:
        from watchdog.observers import Observer

        ENABLE_PROD_TO_STOPS = False  # <- keep False now that controlaltp2.py sends to Stops Queue directly

        prod_dir = r"G:\My Drive\Production Sheets"
        os.makedirs(prod_dir, exist_ok=True)
        if ENABLE_PROD_TO_STOPS:
            prod_observer = Observer()
            prod_observer.schedule(ProdToStopsInterceptor(), path=prod_dir, recursive=False)
            prod_observer.start()
            print(f"[Startup] Interceptor watching: {prod_dir}")
        else:
            print(f"[Startup] Interceptor disabled for: {prod_dir}")

        # Start Stops Queue watcher: opens checkbox UI -> stamp triangles -> copy back to Production
        # DISABLED: PDF stamping moved to another computer
        # stops_dir = r"G:\My Drive\Stops Queue"
        # os.makedirs(stops_dir, exist_ok=True)
        # sq_observer = Observer()
        # sq_observer.schedule(StopsQueueHandler(), path=stops_dir, recursive=False)
        # sq_observer.start()
        # print(f"[Startup] Stops Queue watcher started: {stops_dir}")
        print("[Startup] Stops Queue watcher DISABLED (running on another computer).")

    except Exception as e:
        print("[Startup] Local watchers failed:", e)



    # 5) Start listeners for scanner input
    #    a) Clipboard listener (for scanners that copy URL to clipboard)
    try:
        threading.Thread(target=clipboard_listener, daemon=True).start()
        print("[Startup] Clipboard listener started.")
    except Exception as e:
        print("[Startup] Clipboard listener failed:", e)

    #    b) Scanner keystroke listener (NORMAL typing mode)
    try:
        SCANNER_HOOK = keyboard.on_press(_scanner_keyboard_hook, suppress=False)
        print("[Startup] Scanner listener registered (non-suppressing).")
    except Exception as e:
        print("[Startup] Scanner listener failed:", e)






    # 6) Hotkey listeners
    #    Ctrl+Alt+O
    threading.Thread(target=listen_for_ctrl_alt_o, daemon=True).start()
    #    Ctrl+Alt+N
    threading.Thread(target=listen_for_ctrl_alt_n, daemon=True).start()
    #    Ctrl+Alt+P (external listener from controlaltp2)
    # hotkey_thread = threading.Thread(target=listen_for_hotkey, daemon=True)
    # hotkey_thread.start()

    # Ensure hotkey callbacks are registered
    print("[Startup] Registering hotkeys.")
    keyboard.add_hotkey("ctrl+alt+t", on_control_alt_t)
    print("[Startup] Hotkey registered: Ctrl+Alt+T -> on_control_alt_t()")
    keyboard.add_hotkey("ctrl+alt+p", on_control_alt_p)
    print("[Startup] Hotkey registered: Ctrl+Alt+P -> on_control_alt_p()")

    keyboard.add_hotkey("ctrl+alt+d", lambda: run_on_tk_thread(_run_pdf_presentation_picker_on_root))
    print("[Startup] Hotkey registered: Ctrl+Alt+D -> Tk-dispatched PDF picker")

    # Ensure Tk service thread is running before any watcher tries to show UI
    try:
        ensure_tk_service()
        print("[Startup] Tk service thread started.")
    except Exception as e:
        print("[Startup] Tk service failed:", e)

    # 7) Block forever so threads stay alive
    print("Service running. Press Ctrl+C to exit.")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("Shutting down...")

keyboard.hook(_scanner_keyboard_hook, suppress=False)
print("[Scanner] Scanner hook installed (non-suppressing)")

