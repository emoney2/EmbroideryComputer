# controlalto.py

import os
import time
import keyboard
import tkinter as tk
from tkinter import simpledialog, messagebox
import logging

# ─── Configuration ─────────────────────────────────────────────────────────────
BASE_PATH = r"G:\My Drive\Orders"
HOTKEY    = "ctrl+alt+o"
# ────────────────────────────────────────────────────────────────────────────────

# Lower log level
logging.basicConfig(
    level=logging.WARNING,
    format='[%(asctime)s] %(levelname)s: %(message)s'
)

# Guard to prevent duplicate prompts
_busy = False


def prompt_order_number():
    """Prompt for the order number in a frontmost dialog."""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    order = simpledialog.askstring(
        "Open Order File", "Enter Order Number:", parent=root
    )
    root.destroy()
    return order.strip() if order else None


def open_order_file():
    """Open .emb if exists, else show a single frontmost error popup."""
    global _busy
    if _busy:
        return
    _busy = True

    try:
        order = prompt_order_number()
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
            err.destroy()
    finally:
        # Wait until keys are fully released
        time.sleep(0.5)
        _busy = False


def listen_for_hotkey():
    """Blockingly wait for HOTKEY, then handle once per press."""
    while True:
        # Wait for the full combination
        keyboard.wait(HOTKEY)
        open_order_file()

if __name__ == "__main__":
    listen_for_hotkey()
