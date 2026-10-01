import os
import fitz

def convert_pdf_to_word(file_paths, target_dir):
    try:
        from pdf2docx import Converter
        
        # Sembunyikan warning MuPDF internal
        try:
            fitz.TOOLS.mupdf_display_errors(False)
        except Exception:
            pass

        success = 0
        for pdf_path in file_paths:
            if not os.path.exists(pdf_path):
                continue
                
            base_name = os.path.basename(pdf_path)
            name, _ = os.path.splitext(base_name)
            docx_path = os.path.join(target_dir, f"{name}.docx")

            cv = Converter(pdf_path)
            cv.convert(docx_path, start=0, end=None)
            cv.close()
            success += 1

        return success, len(file_paths), ""
    except Exception as e:
        return 0, len(file_paths), f"Gagal konversi PDF ke Word: {str(e)}"