import win32com.client
import time

def run_wilcom_export(emb_file, svg_file):
    # The ProgID "Wilcom.EmbroideryStudio" is illustrative—check your documentation.
    app = win32com.client.Dispatch("Wilcom.EmbroideryStudio")
    app.OpenFile(emb_file)
    # Wait for the file to load.
    time.sleep(2)
    # Set export options if available; this is an example.
    app.ExportOptions.IncludeVectorFills = True
    app.ExportAsSVG(svg_file)
    app.CloseFile()

if __name__ == "__main__":
    emb_path = r"C:\Users\eckar\Desktop\EMB\10.emb"
    svg_path = r"C:\Users\eckar\Desktop\EMB\10_exported.svg"
    run_wilcom_export(emb_path, svg_path)
