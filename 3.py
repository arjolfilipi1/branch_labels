import os
import tempfile
import subprocess
import time
from pylibdmtx.pylibdmtx import encode
from reportlab.lib.pagesizes import mm
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from io import BytesIO
from PIL import Image

def print_labels_directly(data_list, sumatra_path=None):
    """Print labels directly using SumatraPDF"""
    
    # Try to find SumatraPDF if path not specified
    if sumatra_path is None:
        sumatra_path = find_sumatra()
        if sumatra_path is None:
            raise FileNotFoundError("SumatraPDF not found. Please specify path or install it.")

    # Create temporary PDF file (kept open to prevent deletion)
    temp_pdf = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False)
    temp_pdf_path = temp_pdf.name
    temp_pdf.close()  # Close handle so Sumatra can access it

    try:
        # Generate PDF
        c = canvas.Canvas(temp_pdf_path, pagesize=(49*mm, 9*mm))
        
        for item in data_list:
            # [Your existing PDF generation code here]
            # Text
            font_size = 6
            c.setFont("Helvetica", font_size)
            while c.stringWidth(item, "Helvetica", font_size) > 39*mm - 2*mm and font_size > 3:
                font_size -= 0.5
            
            text_width = c.stringWidth(item, "Helvetica", font_size)
            text_x = (39*mm - text_width) / 2
            text_y = (9*mm - font_size) / 2
            
            c.setFont("Helvetica", font_size)
            c.drawString(text_x, text_y, item)
            
            # Barcode
            encoded = encode(item.encode('utf-8'))
            img = Image.frombytes('RGB', (encoded.width, encoded.height), encoded.pixels)
            img_buffer = BytesIO()
            img.save(img_buffer, format='PNG')
            img_buffer.seek(0)
            
            c.drawImage(ImageReader(img_buffer), 
                       49*mm - 10*mm - 1*mm, 
                       (9*mm - 8*mm)/2, 
                       width=8*mm, 
                       height=8*mm,
                       preserveAspectRatio=True)
            
            c.showPage()
        
        c.save()

        # Print with SumatraPDF
        print_command = [
            sumatra_path,
            '-print-to-default',
            '-silent',
            temp_pdf_path
        ]
        
        # Wait briefly to ensure file is fully written
        time.sleep(0.5)
        
        result = subprocess.run(print_command, capture_output=True, text=True)
        
        if result.returncode != 0:
            raise RuntimeError(f"Printing failed: {result.stderr}")
        
        print("Labels sent to printer successfully")
        
    finally:
        # Clean up - try to delete the temp file
        try:
            os.unlink(temp_pdf_path)
        except:
            pass

def find_sumatra():
    """Try to locate SumatraPDF automatically"""
    common_paths = [
        r'C:\Program Files\SumatraPDF\SumatraPDF.exe',
        r'C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe',
        os.path.expanduser(r'~\AppData\Local\SumatraPDF\SumatraPDF.exe'),
        'SumatraPDF.exe'  # In PATH
    ]
    
    for path in common_paths:
        if os.path.exists(path):
            return path
    return None

# Example usage:
data_list = ["71A05X1", "X1824"]

# Option 1: Auto-detect SumatraPDF
print_labels_directly(data_list)

# Option 2: Specify path explicitly
# print_labels_directly(data_list, r'C:\path\to\SumatraPDF.exe')