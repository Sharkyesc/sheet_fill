import os
import time
from typing import List, Dict, Any
from docx2pdf import convert
from pdf2image import convert_from_path
from PIL import Image
import base64
from io import BytesIO
from config import Config
import fitz  # PyMuPDF

class PDFProcessor:
    def __init__(self):
        self.temp_dir = Config.TEMP_DIR
        os.makedirs(self.temp_dir, exist_ok=True)

    def docx_to_pdf(self, docx_path: str) -> str:
        """Convert docx file to PDF"""
        try:
            base_name = os.path.splitext(os.path.basename(docx_path))[0]
            pdf_path = os.path.join(self.temp_dir, f"{base_name}_{int(time.time())}.pdf")
            
            print(f"Converting {docx_path} to PDF...")
            convert(docx_path, pdf_path)
            print(f"PDF conversion completed: {pdf_path}")
            return pdf_path
        except Exception as e:
            print(f"PDF conversion failed: {e}")
            raise

    def pdf_to_images(self, pdf_path: str, dpi: int = 200) -> List[Dict[str, Any]]:
        """Use PyMuPDF to convert each page of PDF to an image, save to TMP_DIR, and return image path list"""
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        page_images = []
        doc = fitz.open(pdf_path)
        for page_num in range(len(doc)):
            page = doc[page_num]
            # Set resolution
            zoom = dpi / 72  # PyMuPDF default is 72dpi
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat)
            img_filename = os.path.join(self.temp_dir, f"{base_name}_{page_num+1}.png")
            pix.save(img_filename)
            page_images.append({
                'page_number': page_num + 1,
                'image_path': img_filename,
                'image_format': 'PNG',
                'dpi': dpi
            })
        doc.close()
        print(f"PDF to image conversion completed, total {len(page_images)} pages")
        return page_images

    def process_document_with_images(self, docx_path: str) -> Dict[str, Any]:
        """Process document: convert to PDF and generate page screenshots"""
        try:
            pdf_path = self.docx_to_pdf(docx_path)
            
            page_images = self.pdf_to_images(pdf_path)
            
            return {
                'pdf_path': pdf_path,
                'page_images': page_images,
                'total_pages': len(page_images)
            }
            
        except Exception as e:
            print(f"Document processing failed: {e}")
            raise

    def cleanup_temp_files(self, pdf_path: str = None):
        """Clean up temporary files"""
        try:
            if pdf_path and os.path.exists(pdf_path):
                os.remove(pdf_path)
                print(f"Temporary PDF file deleted: {pdf_path}")
        except Exception as e:
            print(f"Failed to clean up temporary files: {e}") 