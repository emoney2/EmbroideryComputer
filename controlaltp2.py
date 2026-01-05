import os
import time
import traceback
import re
import shutil
import threading
from datetime import datetime

import pdfplumber
import pyautogui
import pygetwindow as gw
from plyer import notification
from pywinauto import Application, Desktop
from plyer import notification

# === CONFIGURATION ===
# === CONFIGURATION ===
PREP_FOLDER = r"G:\My Drive\Prep Sheets"
PRODUCTION_FOLDER = r"G:\My Drive\Production Sheets"  # still used elsewhere
STOPS_QUEUE_FOLDER = r"G:\My Drive\Stops Queue"       # ← NEW: validated PDFs go here
LOG_FILE = "auto_pdf_log.txt"


import keyboard
import time

has_triggered_once = False

def guarded_run():
    global has_triggered_once
    if has_triggered_once:
        run_full_sequence()
    else:
        has_triggered_once = True
        log("⚠️ First Ctrl+Alt+P ignored to avoid auto-trigger.")

def on_ctrl_alt_p():
    # 1️⃣ show “starting” toast
    notification.notify(
        title="JR & Co Automation",
        message="🔑 Ctrl+Alt+P detected — process starting!",
        timeout=4
    )
    log("🔔 Desktop notification sent; beginning run_full_sequence()")

    try:
        run_full_sequence()
        log("✅ run_full_sequence() completed successfully")
    except Exception as err:
        # 2️⃣ log full traceback to console
        log("❌ Error in run_full_sequence():\n" + traceback.format_exc())
        # 3️⃣ notify user of failure
        notification.notify(
            title="JR & Co Automation — ERROR",
            message="❌ Process failed! Check logs or restart the app.",
            timeout=6
        )

def listen_for_hotkey():
    log("🕒 Waiting 3 seconds before registering hotkey to avoid auto-trigger...")
    time.sleep(3)

    log("🔑 Hotkey registered: Ctrl+Alt+P → on_ctrl_alt_p()")
    keyboard.add_hotkey("ctrl+alt+p", on_ctrl_alt_p)
    
    log("👀 Listening in background. Press Ctrl+C to exit.")
    keyboard.wait()

def log(msg):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] {msg}\n")
    print(f"[{timestamp}] {msg}")

def click_save_button():
    from pywinauto import Desktop
    import time

    print("💾 Looking for 'Save' button in main EmbroideryStudio window...")

    try:
        win = Desktop(backend="uia").window(title_re=".*EmbroideryStudio.*", found_index=0)
        all_controls = win.descendants()
        for ctrl in all_controls:
            if ctrl.element_info.control_type == "Button" and ctrl.element_info.name.strip() == "Save":
                ctrl.set_focus()
                time.sleep(0.5)
                ctrl.invoke()
                print("🖱️ Clicked 'Save' button")
                return True

        print("❌ 'Save' button not found")
        return False

    except Exception as e:
        print(f"❌ Error while clicking 'Save' button: {e}")
        return False

def get_emb_name_from_title():
    try:
        win = gw.getWindowsWithTitle("EmbroideryStudio 2025")
        if not win:
            log("❌ Wilcom window not found.")
            return None
        title = win[0].title
        match = re.search(r"- \[(.+?)\] - \[", title)
        if not match:
            log("❌ Could not extract file name from title.")
            return None
        filename = match.group(1).strip()
        if not filename:
            log("❌ Extracted filename is empty.")
            return None
        return filename.split()[0]
    except Exception as e:
        log(f"❌ Error getting file name: {e}")
        return None

def click_print_preview_button():
    from pywinauto import Desktop
    import time

    print("🔍 Looking for 'Preview' button in main EmbroideryStudio window...")

    try:
        win = Desktop(backend="uia").window(title_re=".*EmbroideryStudio.*", found_index=0)
        all_controls = win.descendants()
        for ctrl in all_controls:
            if ctrl.element_info.control_type == "Button" and ctrl.element_info.name.strip() == "Preview":
                ctrl.set_focus()
                time.sleep(0.5)
                ctrl.invoke()
                print("🖱️ Clicked 'Preview' button")
                return True

        print("❌ 'Preview' button not found")
        return False

    except Exception as e:
        print(f"❌ Error while clicking 'Preview' button: {e}")
        return False

def press_save_as_pdf(timeout: int = 30) -> bool:
    """
    Within the main EmbroideryStudio window, find the Save-as-PDF button
    and invoke it (no mouse movement).
    """
    try:
        log("[Ctrl+Alt+P] 🧠 Attaching to desktop (uia backend)...")
        desk = Desktop(backend="uia")

        # 1) Find the main EmbroideryStudio 2025 window
        log("[Ctrl+Alt+P] 🔍 Looking for EmbroideryStudio 2025 main window...")
        es_window = None
        end_time = time.time() + timeout

        while time.time() < end_time and es_window is None:
            for w in desk.windows():
                try:
                    title = (w.window_text() or "").lower()
                except Exception:
                    continue

                # Match the ES main window by title
                if "embroiderystudio 2025" in title:
                    es_window = w
                    break

            if es_window is None:
                time.sleep(0.5)

        if es_window is None:
            log("[Ctrl+Alt+P] ❌ Could not find EmbroideryStudio 2025 window within timeout.")
            return False

        log(f"[Ctrl+Alt+P] ✅ Found EmbroideryStudio window: '{es_window.window_text()}'")
        try:
            es_window.set_focus()
        except Exception as e:
            log(f"[Ctrl+Alt+P] ⚠️ Could not set focus to ES window: {e}")

        # Give Wilcom a moment after the Preview click to finish drawing UI
        time.sleep(1.0)

        # 2) Search for a button whose name contains 'pdf' inside ES window
        log("[Ctrl+Alt+P] 🔍 Searching for a button containing 'PDF' inside EmbroideryStudio...")
        try:
            buttons = es_window.descendants(control_type="Button")
        except Exception as e:
            log(f"[Ctrl+Alt+P] ❌ Failed to enumerate buttons: {e}")
            return False

        pdf_button = None
        for btn in buttons:
            try:
                name = btn.window_text() or ""
            except Exception:
                continue

            if "pdf" in name.lower():
                pdf_button = btn
                log(f"[Ctrl+Alt+P] ✅ Found candidate PDF button: '{name}'")
                break

        if pdf_button is None:
            log("[Ctrl+Alt+P] ❌ No button containing 'pdf' was found in EmbroideryStudio. Available buttons:")
            for btn in buttons:
                try:
                    name = btn.window_text()
                except Exception:
                    name = "<unreadable>"
                log(f"    - Button: '{name}'")
            return False

        # 3) Invoke the button directly (no mouse movement)
        try:
            time.sleep(0.3)  # tiny delay to let UI settle
            pdf_button.invoke()
        except Exception as e:
            log(f"[Ctrl+Alt+P] ❌ Failed to invoke PDF button: {e}\n{traceback.format_exc()}")
            return False

        log("[Ctrl+Alt+P] ✅ Successfully invoked the Save-as-PDF button.")
        return True

    except Exception as e:
        log(f"[Ctrl+Alt+P] ❌ Exception in press_save_as_pdf(): {e}\n{traceback.format_exc()}")
        return False


def type_save_path(filepath):
    from pywinauto.keyboard import send_keys
    import time

    log("⏳ Waiting 3 seconds for Save As dialog to appear...")
    time.sleep(3)  # fixed delay before typing

    try:
        send_keys(filepath, with_spaces=True)
        time.sleep(0.2)
        send_keys("{ENTER}")
        log(f"💾 Typed save path: {filepath}")
    except Exception as e:
        log(f"❌ Failed to type save path: {e}")


def wait_for_file(path, timeout=10):
    for _ in range(timeout * 2):
        if os.path.exists(path):
            log(f"📄 File saved: {path}")
            return True
        time.sleep(0.5)
    log(f"❌ PDF file not found after {timeout}s: {path}")
    return False

def validate_thread_codes(pdf_path):
    log("📄 Validating thread codes in preview PDF...")

    try:
        # Wait for file to be fully written and accessible
        for _ in range(20):  # 10 seconds max
            if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
                try:
                    with open(pdf_path, "rb") as f:
                        f.read(1)
                    break  # File is accessible
                except Exception:
                    pass
            time.sleep(0.5)
        else:
            log("❌ File not ready after waiting")
            return False

        # Open and extract text
        with pdfplumber.open(pdf_path) as pdf:
            full_text = "\n".join(
                page.extract_text() or "" for page in pdf.pages
            )

        # Extract lines containing thread info
        thread_lines = [
            line.strip()
            for line in full_text.splitlines()
            if re.search(r"^\s*\d+\.\s+\d+\s+\d{1,4}\w*", line)
        ]

        # Extract the code portion (3rd column after number and color ID)
        invalid_codes = []
        for line in thread_lines:
            parts = line.strip().split()
            if len(parts) >= 3:
                code = parts[2]
                if not re.fullmatch(r"\d{4}", code):
                    invalid_codes.append(code)

        if invalid_codes:
            for code in invalid_codes:
                log(f"❌ Invalid code found: {code}")
            os.remove(pdf_path)
            log("🗑️ Invalid codes found – deleted PDF")
            return False

        log("✅ All thread codes valid")
        return True

    except Exception as e:
        log(f"❌ PDF validation error: {e}")
        try:
            os.remove(pdf_path)
        except Exception:
            pass
        return False


def show_notification(title, message):
    notification.notify(title=title, message=message, timeout=6)


def run_full_sequence():
    print("🚀 Triggered PDF generation and validation sequence")

    # 💾 Save the open embroidery file
    if not click_save_button():
        show_notification("Embroidery Script", "❌ Could not press Save")
        return

    # 🧠 Get filename after saving
    emb_name = get_emb_name_from_title()
    if not emb_name:
        show_notification("Embroidery Script", "No active .emb file found")
        return

    prep_pdf_path  = os.path.join(PREP_FOLDER, emb_name + ".pdf")
    stops_pdf_path = os.path.join(STOPS_QUEUE_FOLDER, emb_name + ".pdf") 

    # 🖨️ Open Print Preview
    if not click_print_preview_button():
        print("❌ Could not open Print Preview window")
        return

    # 💾 Click Save As PDF
    if not press_save_as_pdf():
        show_notification("Embroidery Script", "Could not press Save as PDF")
        return

    type_save_path(prep_pdf_path)

    if not wait_for_file(prep_pdf_path):
        show_notification("Embroidery Script", "PDF did not save in time")
        return

    if validate_thread_codes(prep_pdf_path):
        shutil.move(prep_pdf_path, stops_pdf_path)
        log(f"📂 Moved to Stops Queue: {stops_pdf_path}")
        pyautogui.hotkey("shift", "alt", "q")
        log("🎯 Queued embroidery file (Shift+Alt+Q)")
        show_notification("Embroidery Script", "✅ PDF saved & queued!")
    else:
        try:
            os.remove(prep_pdf_path)
        except Exception:
            pass
        log("🗑️ Invalid codes found – deleted PDF")
        show_notification("Embroidery Script", "❌ Invalid thread codes in PDF. Check the file.")