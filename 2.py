import os
from pylibdmtx.pylibdmtx import encode
from reportlab.lib.pagesizes import mm
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from io import BytesIO
from PIL import Image

def create_labels_with_barcode_images(data_list, output_pdf="labels.pdf"):
    
    # Label dimensions
    label_width = 49 * mm
    label_height = 9 * mm
    
    # Text area dimensions
    text_width = 39 * mm
    text_height = 9 * mm  # Available height for text
    
    # Data matrix dimensions
    dm_width = 10 * mm
    dm_height = 10 * mm
    
    # Create PDF
    c = canvas.Canvas(output_pdf, pagesize=(label_width, label_height))
    
    for index, item in enumerate(data_list):
        # Start new page for each label
        c.setPageSize((label_width, label_height))
        
        # TEXT (Left side)
        font_size = 6
        c.setFont("Helvetica", font_size)
        
        # Adjust font size if needed
        while c.stringWidth(item, "Helvetica", font_size) > text_width - 2*mm and font_size > 3:
            font_size -= 0.5
        
        # Calculate text metrics for vertical centering
        text_width_actual = c.stringWidth(item, "Helvetica", font_size)
        text_height_actual = font_size * 1.2  # Approximate text height
        
        # Calculate horizontal centering for text
        text_x = (text_width - text_width_actual) / 2
        
        # Calculate vertical centering - accounts for actual text height
        vertical_offset = (label_height - text_height_actual) / 2
        
        # Draw centered text (both horizontally and vertically)
        c.setFont("Helvetica", font_size)
        c.drawString(text_x, vertical_offset, item)
        
        # Generate DataMatrix barcode
        encoded = encode(
            item.encode('utf-8'),
            size='SquareAuto'
        )
        
        # Create PIL Image
        img = Image.frombytes('RGB', (encoded.width, encoded.height), encoded.pixels)
        
        # Prepare image for PDF
        img_buffer = BytesIO()
        img.save(img_buffer, format='PNG', dpi=(300, 300))
        img_buffer.seek(0)
        
        # BARCODE (Right side) - Centered vertically with text
        dm_x = label_width - dm_width - 1 * mm
        dm_y = (label_height - dm_height) / 2  # Center vertically
        
        # Add to PDF
        c.drawImage(
            ImageReader(img_buffer),
            dm_x, dm_y,
            width=dm_width,
            height=dm_height,
            preserveAspectRatio=True,
            mask='auto'
        )
        
        c.showPage()
    
    c.save()
    print(f"PDF with perfectly centered labels created: {output_pdf}")

# Example usage
data_list = ["71A05/X1", "X1824"]
create_labels_with_barcode_images(data_list)