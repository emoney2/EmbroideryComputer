# queue_service.py

from flask import Flask, request
import os
import time
from pywinauto import Application

app = Flask(__name__)

# Folder where your .emb files live:
EMB_FOLDER = r"C:\Users\eckar\Desktop\EMB"

# Matches any EmbroideryStudio 2025 window (regardless of file name):
EMB_WINDOW_TITLE = r"EmbroideryStudio\s*2025.*"

@app.route("/queue")
def queue_design():
    # 1) Get the base filename from the URL, e.g. ?file=Order123
    file_base = request.args.get("file")
    if not file_base:
        return "Error: no file specified", 400

    # 2) Build the full .emb path
    emb_path = os.path.join(EMB_FOLDER, f"{file_base}.emb")
    if not os.path.exists(emb_path):
        return f"Error: {emb_path} not found", 404

    # 3) Open the .emb in EmbroideryStudio (default app)
    os.startfile(emb_path)

    # 4) Poll up to 5 minutes for the window to appear
    max_wait = 300   # seconds
    interval = 2     # seconds between checks
    start_time = time.time()
    while True:
        try:
            emb = Application(backend="uia").connect(title_re=EMB_WINDOW_TITLE)
            win = emb.window(title_re=EMB_WINDOW_TITLE)
            break
        except Exception:
            if time.time() - start_time > max_wait:
                return (f"Error: EmbroideryStudio did not appear "
                        f"in {max_wait} seconds"), 500
            time.sleep(interval)

    # 5) Once it’s up, send Shift+Alt+Q silently
    try:
        win.type_keys("+!q", set_foreground=False)
        return "Queued ✅", 200
    except Exception as e:
        print("Queue error sending keys:", e)
        return f"Error sending keys: {e}", 500

if __name__ == "__main__":
    print("🚀 queue_service.py starting up…")
    app.run(host="0.0.0.0", port=5001, debug=True)
