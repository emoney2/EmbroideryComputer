import keyboard
import time
import tkinter as tk
from tkinter import simpledialog, messagebox
import logging
import os

# Configure logging: show INFO and above
logging.basicConfig(level=logging.INFO, format='[%(asctime)s] %(levelname)s: %(message)s')


def prompt_order_number():
    # Use a single simple dialog for reliability on repeated calls
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    order = simpledialog.askstring(
        "Enter Order Number",
        "Order Number:",
        parent=root
    )
    root.destroy()
    return order.strip() if order else None


def save_as_order_pdf():
    # Prompt once
    order_number = prompt_order_number()
    if not order_number:
        logging.info("No order number provided, aborting.")
        return

    # Trigger Save As via Ctrl+S
    keyboard.send('ctrl+s')

    # Wait for dialog
    time.sleep(2)

    # Focus filename via Alt+N
    keyboard.send('alt+n')
    time.sleep(0.1)

    # Type path and filename (no extension)
    target = f"G:\\My Drive\\Orders\\{order_number}\\{order_number}"
    keyboard.write(target)

    # Check for existing file (assumes .emb default)
    full_path = target + ".emb"
    if os.path.exists(full_path):
        messagebox.showwarning(
            "File already exists",
            f"A file named {os.path.basename(full_path)} already exists."
        )
        return

    # Confirm and show saved popup
    keyboard.send('enter')
    messagebox.showinfo(
        "Saved!",
        f"File saved to Orders/{order_number}/{order_number}"
    )


def listen_for_save_hotkey():
    logging.info("Registering Ctrl+Alt+N hotkey")
    keyboard.add_hotkey('ctrl+alt+n', save_as_order_pdf)
    keyboard.wait()

if __name__ == '__main__':
    listen_for_save_hotkey()
