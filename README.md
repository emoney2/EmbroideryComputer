# Embroidery Computer

Order entry, inventory management, QR code scanning, and PDF printing system for embroidery operations.

## Features

- **Order Entry & Management**: Create and manage embroidery orders
- **Inventory Tracking**: Track thread inventory and materials
- **QR Code Scanning**: Scan QR codes to update order status and quantities
- **PDF Generation & Printing**: Generate stamped PDFs and print labels
- **Label Printing**: Automatic folder watcher for printing UPS labels to PL80E printer
- **Google Sheets Integration**: Sync data with Google Sheets
- **Wilcom Integration**: Queue embroidery files in Wilcom software

## Setup

1. Install Python dependencies:
```bash
pip install -r requirements.txt
```

2. Configure Google Sheets credentials:
   - Place your service account JSON file in the `Keys/` folder
   - Update the credentials path in the code if needed

3. Configure printer settings:
   - Ensure PL80E printer is installed and set to 4×6 inch paper size
   - Update `PRINTER_NAME` in the code if your printer name differs

## Usage

Run the main application:
```bash
python CreateStampedPDFandSubmitStitchCount11.py
```

The application will:
- Start watching the "Label Printer" folder for new PDFs
- Automatically print labels to the PL80E printer
- Process QR code scans
- Manage order workflows

## Folder Structure

- `CreateStampedPDFandSubmitStitchCount11.py` - Main application file
- `JRCO_PrintServer.py` - Print server for label printing
- `machine-scheduler-backend/` - Backend server for machine scheduling
- `Keys/` - Google service account credentials (not committed)
- `Old Versions/` - Previous versions of scripts

## License

Private project - All rights reserved
