from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from io import BytesIO
import os
import subprocess
import tempfile
from datetime import datetime
import re
import logging
from apify import Actor

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Config
today_date = datetime.now().strftime("%m/%d/%Y")  # Dynamically get current date
SIGNATURE_MARKER = "Signature:"
DATE_MARKER = "Date:"
DATE_PATTERN = r'\b(\d{1,2}/\d{1,2}/\d{4})\b'
default_signature_coords = (250, 293)
default_date_coords = (260, 259)
default_signature_p2 = (230, 162)
font_size = 9
signature_width = 120
signature_height = 30
blank_image_width = 80
blank_image_height = 10
signature_width2 = 160
signature_height2 = 20

async def main():
    async with Actor:
        # Get input from Apify
        input_data = await Actor.get_input() or {}
        xlsx_key = input_data.get('xlsx_key')  # Key for XLSX file in Apify key-value store
        signature_key = input_data.get('signature_key', 'signature.png')
        blank_image_key = input_data.get('blank_image_key', 'line.png')

        if not xlsx_key:
            raise ValueError("No XLSX file key provided in input")

        # Retrieve files from Apify key-value store
        kv_store = await Actor.open_key_value_store()
        xlsx_data = await kv_store.get_value(xlsx_key)
        signature_data = await kv_store.get_value(signature_key)
        blank_image_data = await kv_store.get_value(blank_image_key)

        if not xlsx_data or not signature_data or not blank_image_data:
            raise FileNotFoundError("One or more input files not found in key-value store")

        # Save files to temporary directory
        with tempfile.TemporaryDirectory() as temp_dir:
            xlsx_path = os.path.join(temp_dir, 'input.xlsx')
            signature_path = os.path.join(temp_dir, 'signature.png')
            blank_image_path = os.path.join(temp_dir, 'line.png')
            output_pdf_path = os.path.join(temp_dir, 'output.pdf')

            with open(xlsx_path, 'wb') as f:
                f.write(xlsx_data)
            with open(signature_path, 'wb') as f:
                f.write(signature_data)
            with open(blank_image_path, 'wb') as f:
                f.write(blank_image_data)

            # Convert XLSX to PDF
            pdf_path = convert_xlsx_to_pdf(xlsx_path, temp_dir)

            # Add signature and date
            add_signature_and_date(pdf_path, signature_path, blank_image_path, output_pdf_path, today_date)

            # Save output PDF to key-value store
            with open(output_pdf_path, 'rb') as f:
                output_key = f"signed_{xlsx_key}.pdf"
                await kv_store.set_value(output_key, f.read(), content_type='application/pdf')

            # Log result and push to dataset
            result = {"status": "success", "output_file": output_key}
            await Actor.push_data(result)
            logger.info(f"Processed file: {result}")

def convert_xlsx_to_pdf(xlsx_path, out_dir):
    """Convert XLSX to PDF using LibreOffice."""
    base = os.path.splitext(os.path.basename(xlsx_path))[0]
    pdf_out = os.path.join(out_dir, f"{base}.pdf")
    soffice = "/usr/lib/libreoffice/program/soffice"  # Path in Apify container
    cmd = [
        soffice, "--headless",
        "--convert-to", "pdf",
        "--outdir", out_dir,
        xlsx_path
    ]
    logger.info(f"Running conversion: {' '.join(cmd)}")
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode == 0 and os.path.exists(pdf_out):
        logger.info(f"Converted to PDF: {pdf_out}")
        return pdf_out
    raise FileNotFoundError(f"Conversion failed: {result.stderr}")

def find_text_coordinates(page, marker):
    """Extract text and find coordinates of the marker."""
    try:
        text_data = page.extract_text()
        if not text_data:
            return None, False
        lines = text_data.split('\n')
        for i, line in enumerate(lines):
            if marker.lower() in line.lower():
                y = letter[1] - (i + 1) * 20
                x = 100
                has_date = False
                if marker == DATE_MARKER:
                    current_line = line.strip()
                    next_line = lines[i + 1].strip() if i + 1 < len(lines) else ""
                    date_match = re.search(DATE_PATTERN, current_line) or re.search(DATE_PATTERN, next_line)
                    has_date = bool(date_match)
                return (x, y), has_date
        return None, False
    except Exception:
        return None, False

def add_signature_and_date(input_pdf, signature_img, blank_img, output_pdf, date_text):
    """Overlay signature and date on PDF."""
    reader = PdfReader(input_pdf)
    writer = PdfWriter()
    total = len(reader.pages)
    end_ix = min(19, total)
    logger.info(f"Processing PDF with {total} pages, limiting to {end_ix}")

    # Process page 1
    page0 = reader.pages[0]
    packet0 = BytesIO()
    c0 = canvas.Canvas(packet0, pagesize=letter)
    sig_coords0, _ = find_text_coordinates(page0, SIGNATURE_MARKER)
    date_coords, has_date = find_text_coordinates(page0, DATE_MARKER)
    sig_coords0 = sig_coords0 or default_signature_coords
    sig_coords0 = (sig_coords0[0] + 80, sig_coords0[1] - 10)
    date_coords = date_coords or default_date_coords
    date_coords = (date_coords[0] + 80, date_coords[1] - 10)

    c0.drawImage(signature_img, sig_coords0[0], sig_coords0[1], width=signature_width, height=signature_height)
    if not has_date:
        c0.drawImage(blank_img, date_coords[0], date_coords[1], width=blank_image_width, height=blank_image_height)
        c0.setFont("Helvetica", font_size)
        date_y = date_coords[1] + blank_image_height / 2
        c0.drawString(date_coords[0], date_y, date_text)
    c0.save()
    packet0.seek(0)
    overlay0 = PdfReader(packet0).pages[0]
    page0.merge_page(overlay0)
    writer.add_page(page0)

    # Process page 2 if it exists
    if total > 1:
        page1 = reader.pages[1]
        packet1 = BytesIO()
        c1 = canvas.Canvas(packet1, pagesize=letter)
        sig_coords1, _ = find_text_coordinates(page1, SIGNATURE_MARKER)
        sig_coords1 = sig_coords1 or default_signature_p2
        sig_coords1 = (sig_coords1[0] + 80, sig_coords1[1] - 10)
        c1.drawImage(signature_img, sig_coords1[0], sig_coords1[1], width=signature_width2, height=signature_height2)
        c1.save()
        packet1.seek(0)
        overlay1 = PdfReader(packet1).pages[0]
        page1.merge_page(overlay1)
        writer.add_page(page1)
        for p in reader.pages[2:end_ix]:
            writer.add_page(p)
    else:
        for p in reader.pages[1:end_ix]:
            writer.add_page(p)

    with open(output_pdf, "wb") as f:
        writer.write(f)
    logger.info(f"Saved {end_ix}-page PDF as {output_pdf}")

if __name__ == "__main__":
    Actor.main(main)