import os
from PIL import Image
import pymupdf as fitz  # PyMuPDF (Menggunakan import baru agar tidak ada warning)
from docx import Document
from docx.shared import Inches


def convert_image_to_pdf(input_files, output_directory="", status_callback=None):
    """Mengonversi gambar (JPG, JPEG, PNG) ke PDF."""
    total = len(input_files)
    success_count = 0
    last_error = ""

    for idx, img_path in enumerate(input_files, start=1):
        img_path = os.path.abspath(img_path)
        file_name = os.path.basename(img_path)
        file_title, _ = os.path.splitext(file_name)

        if status_callback:
            status_callback(f"Konversi Gambar ({idx}/{total}): {file_name}...")

        target_dir = (
            os.path.abspath(output_directory)
            if output_directory
            else os.path.dirname(img_path)
        )
        output_path = os.path.join(target_dir, f"{file_title}_converted.pdf")

        try:
            image = Image.open(img_path)
            if image.mode != "RGB":
                image = image.convert("RGB")
            
            image.save(output_path, "PDF", resolution=100.0)
            if os.path.exists(output_path):
                success_count += 1
        except Exception as e:
            last_error = str(e)

    return success_count, total, last_error


def convert_pdf_to_image(input_files, output_directory="", status_callback=None, fmt="png"):
    """Mengonversi file PDF ke Gambar (PNG/JPG)."""
    total = len(input_files)
    success_count = 0
    last_error = ""

    for idx, pdf_path in enumerate(input_files, start=1):
        pdf_path = os.path.abspath(pdf_path)
        file_name = os.path.basename(pdf_path)
        file_title, _ = os.path.splitext(file_name)

        target_dir = (
            os.path.abspath(output_directory)
            if output_directory
            else os.path.dirname(pdf_path)
        )

        try:
            doc = fitz.open(pdf_path)
            total_pages = len(doc)
            for page_idx, page in enumerate(doc, start=1):
                if status_callback:
                    status_callback(f"Render PDF -> Gambar: {file_name} (Halaman {page_idx}/{total_pages})...")

                pix = page.get_pixmap(dpi=150)
                out_img_path = os.path.join(target_dir, f"{file_title}_page_{page_idx}.{fmt}")
                pix.save(out_img_path)

            doc.close()
            success_count += 1
        except Exception as e:
            last_error = str(e)

    return success_count, total, last_error


def convert_image_to_word(input_files, output_directory="", status_callback=None):
    """Mengonversi Gambar ke dokumen Word (.docx)."""
    total = len(input_files)
    success_count = 0
    last_error = ""

    for idx, img_path in enumerate(input_files, start=1):
        img_path = os.path.abspath(img_path)
        file_name = os.path.basename(img_path)
        file_title, _ = os.path.splitext(file_name)

        if status_callback:
            status_callback(f"Menyisipkan Gambar ({idx}/{total}): {file_name}...")

        target_dir = (
            os.path.abspath(output_directory)
            if output_directory
            else os.path.dirname(img_path)
        )
        output_path = os.path.join(target_dir, f"{file_title}_converted.docx")

        try:
            doc = Document()
            # Masukkan gambar dengan lebar optimal 6 inci
            doc.add_picture(img_path, width=Inches(6.0))
            doc.save(output_path)
            if os.path.exists(output_path):
                success_count += 1
        except Exception as e:
            last_error = str(e)

    return success_count, total, last_error