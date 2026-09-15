import os
from pdf2docx import Converter


def convert_pdf_to_word(input_files, output_directory="", status_callback=None):
    """Fungsi independen untuk mengonversi daftar file PDF ke Word (.docx).

    Returns:
        tuple: (success_count, total_files, last_error_message)
    """
    total = len(input_files)
    success_count = 0
    last_error = ""

    for idx, input_path in enumerate(input_files, start=1):
        input_path = os.path.abspath(input_path)
        file_name = os.path.basename(input_path)
        file_title, _ = os.path.splitext(file_name)

        if status_callback:
            status_callback(f"Memproses ({idx}/{total}): {file_name}...")

        target_dir = (
            os.path.abspath(output_directory)
            if output_directory
            else os.path.dirname(input_path)
        )
        output_path = os.path.join(target_dir, f"{file_title}_converted.docx")

        try:
            cv = Converter(input_path)
            cv.convert(output_path, start=0, end=None)
            cv.close()
            if os.path.exists(output_path):
                success_count += 1
        except Exception as e:
            last_error = str(e)

    return success_count, total, last_error