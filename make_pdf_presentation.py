# make_pdf_presentation.py
# Build a multi-page PDF from Photoshop/bitmap files.
# Modes:
#   - Default: use file paths from currently OPEN documents (non-destructive).
#   - --pick : ask for a PARENT FOLDER, then let you pick files from there; saves/versions in that folder.
#
# Requirements: pip install pywin32 PyPDF2

import sys
import time
import re
import tempfile
from pathlib import Path
from collections import Counter

import pythoncom
import win32com.client
from win32com.client import gencache
from PyPDF2 import PdfMerger

# ---------------------------
# Small utilities
# ---------------------------

def with_retries(fn, *, retries=40, delay=0.25, busy_hresults=(-2147417846,)):
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

def ensure_photoshop():
    pythoncom.CoInitialize()
    try:
        return gencache.EnsureDispatch("Photoshop.Application")
    except Exception as e:
        print("Could not connect to Photoshop. Open Photoshop and try again.")
        raise e

def ensure_folder(p: Path):
    p.mkdir(parents=True, exist_ok=True)
    return p

def next_version_name(folder: Path, base: str) -> str:
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

# ---------------------------
# Picker helpers (Tk)
# ---------------------------

def pick_parent_folder():
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    folder = filedialog.askdirectory(title="Select the PARENT folder containing your designs")
    root.update()
    root.destroy()
    return Path(folder) if folder else None

def pick_files_in_folder(initialdir: Path):
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    files = filedialog.askopenfilenames(
        title="Select files for PDF (order ~= selection order)",
        initialdir=str(initialdir),
        filetypes=[
            ("Images & PSD", "*.psd;*.psb;*.png;*.jpg;*.jpeg;*.tif;*.tiff"),
            ("Photoshop", "*.psd;*.psb"),
            ("Images", "*.png;*.jpg;*.jpeg;*.tif;*.tiff"),
            ("All files", "*.*"),
        ],
    )
    root.update()
    root.destroy()
    return [Path(f) for f in files]

def all_within_parent(paths, parent: Path) -> bool:
    parent = parent.resolve()
    for p in paths:
        try:
            if parent not in p.resolve().parents and p.resolve().parent != parent:
                return False
        except Exception:
            return False
    return True

# ---------------------------
# Photoshop helpers
# ---------------------------

def get_paths_from_open_docs(ps):
    """Collect file paths from currently open docs. Skip unsaved documents."""
    paths = []
    count = with_retries(lambda: ps.Documents.Count)
    for i in range(1, count + 1):
        d = with_retries(lambda i=i: ps.Documents.Item(i))
        try:
            fullpath_str = with_retries(lambda: str(d.FullName))
            p = Path(fullpath_str)
            if p.exists():
                paths.append(p)
        except Exception:
            pass
    # Keep order but dedupe by path
    seen = set()
    ordered = []
    for p in paths:
        if p not in seen:
            seen.add(p)
            ordered.append(p)
    return ordered

def open_doc(ps, path: Path):
    return with_retries(lambda: ps.Open(str(path)))

def close_doc_no_save(doc):
    with_retries(lambda: doc.Close(2))  # 2 = psDoNotSave

def export_single_page_pdf_from_path(ps, in_path: Path, out_pdf_path: Path):
    """
    Open a file, duplicate/flatten, save as single-page PDF, close.
    """
    doc = open_doc(ps, in_path)
    dup = with_retries(lambda: doc.Duplicate(doc.Name + "_TMP_DUP", True))
    try:
        with_retries(lambda: dup.Flatten())
    except Exception:
        pass

    pdf_opts = win32com.client.Dispatch("Photoshop.PDFSaveOptions")

    def safe_set(obj, attr, value):
        try:
            setattr(obj, attr, value)
        except Exception:
            pass

    safe_set(pdf_opts, "PreserveEditing", False)
    safe_set(pdf_opts, "OptimizeForWeb", False)
    safe_set(pdf_opts, "View", False)
    safe_set(pdf_opts, "Layers", False)
    safe_set(pdf_opts, "EmbedColorProfile", True)
    safe_set(pdf_opts, "JPEGQuality", 10)

    with_retries(lambda: dup.SaveAs(str(out_pdf_path), pdf_opts, True))
    close_doc_no_save(dup)
    close_doc_no_save(doc)

# ---------------------------
# Main flow
# ---------------------------

def main():
    use_picker = any(arg.lower() == "--pick" for arg in sys.argv[1:])
    ps = ensure_photoshop()
    try:
        with_retries(lambda: setattr(ps, "DisplayDialogs", 3))  # 3 = psDisplayNoDialogs
    except Exception:
        pass

    if use_picker:
        parent = pick_parent_folder()
        if not parent:
            print("Canceled (no parent folder selected).")
            return
        paths = pick_files_in_folder(parent)
        if not paths:
            print("No files selected.")
            return
        # Ensure everything is inside the chosen parent (or subfolders)
        if not all_within_parent(paths, parent):
            print("Error: All selected files must be inside the chosen parent folder.")
            return
        target_folder = parent
    else:
        paths = get_paths_from_open_docs(ps)
        if not paths:
            print("No usable open documents found (they might be unsaved). Try --pick to select files.")
            return
        # If using open docs, derive the most common parent
        parents = [p.parent for p in paths]
        if parents:
            target_folder = Counter(parents).most_common(1)[0][0]
        else:
            print("Could not determine a parent folder from open documents.")
            return

    base = target_folder.name + "Designs"
    out_stem = next_version_name(target_folder, base)
    out_pdf = target_folder / f"{out_stem}.pdf"
    ensure_folder(target_folder)

    temp_dir = ensure_folder(Path(tempfile.gettempdir()) / f"ps_pdf_build_{int(time.time())}")
    temp_pdfs = []

    try:
        for idx, p in enumerate(paths, start=1):
            temp_pdf = temp_dir / f"page_{idx:03d}.pdf"
            print(f"Saving {p.name} -> {temp_pdf.name}")
            export_single_page_pdf_from_path(ps, p, temp_pdf)
            temp_pdfs.append(temp_pdf)

        merger = PdfMerger()
        for t in temp_pdfs:
            merger.append(str(t))
        merger.write(str(out_pdf))
        merger.close()

        print(f"\n✅ Done: {out_pdf}")
    finally:
        for t in temp_pdfs:
            try:
                t.unlink(missing_ok=True)
            except Exception:
                pass
        try:
            temp_dir.rmdir()
        except Exception:
            pass

if __name__ == "__main__":
    main()
