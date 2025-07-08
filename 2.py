from pylibdmtx.pylibdmtx import encode
from PIL import Image
import os

def create_datamatrix(data, filename, size=200):
    """
    Create a Data Matrix barcode image
    
    Parameters:
    - data: The data to encode in the barcode
    - filename: Output file name (e.g., 'barcode.png')
    - size: The size (width and height) of the output image in pixels
    """
    try:
        # Encode the data to Data Matrix format
        encoded = encode(data.encode('utf-8'))
        
        # Create an image from the encoded data
        img = Image.frombytes('RGB', (encoded.width, encoded.height), encoded.pixels)
        
        # Resize the image (optional)
        img = img.resize((size, size), Image.NEAREST)
        
        # Save the image
        img.save(filename)
        print(f"Successfully created Data Matrix barcode: {os.path.abspath(filename)}")
    except Exception as e:
        print(f"Error creating barcode: {e}")

# Example usage
if __name__ == "__main__":
    # Customize these values
    data_to_encode = "https://www.example.com/product123"
    output_filename = "datamatrix_barcode.png"
    image_size = 300  # pixels
    
    create_datamatrix(data_to_encode, output_filename, image_size)