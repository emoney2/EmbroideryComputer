import os
import time
import re
import shutil
import threading
from datetime import datetime

import pdfplumber
import pyautogui
import pygetwindow as gw
from plyer import notification
from pywinauto import Application, Desktop

# === CONFIGURATION ===
PREP_FOLDER = r"G:\My Drive\Prep Sheets"
PRODUCTION_FOLDER = r"G:\My Drive\Production Sheets"
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

def listen_for_hotkey():
    log("🕒 Waiting 3 seconds before registering hotkey to avoid auto-trigger...")
    time.sleep(3)

    log("🔑 Hotkey registered: Ctrl+Alt+P → run_full_sequence()")
    keyboard.add_hotkey("ctrl+alt+p", run_full_sequence)
    
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

def press_save_as_pdf():
    log("🧠 Attaching to active EmbroideryStudio window (uia)...")
    try:
        # Attach to any active window with 'Save As PDF' button
        win = Desktop(backend="uia").window(title_re=".*EmbroideryStudio.*", found_index=0)
        log("🔍 Searching for button labeled 'Save As PDF'...")

        save_button = win.child_window(title="Save As PDF", control_type="Button")
        save_button.wait('visible', timeout=10)
        save_button.invoke()
        log("🖱️ Clicked 'Save As PDF' button")
        return True
    except Exception as e:
        raise RuntimeError(f"Could not click 'Save As PDF': {e}")

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

    prep_pdf_path = os.path.join(PREP_FOLDER, emb_name + ".pdf")
    prod_pdf_path = os.path.join(PRODUCTION_FOLDER, emb_name + ".pdf")

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
        shutil.move(prep_pdf_path, prod_pdf_path)
        log(f"📂 Moved to Production Sheets: {prod_pdf_path}")
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

if __name__ == "__main__":
    listen_for_hotkey()