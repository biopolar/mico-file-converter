import os
from pypdf import PdfReader, PdfWriter


def merge_pdfs(input_files, output_directory="", status_callback=None):
    """Menggabungkan beberapa file PDF menjadi 1 file PDF."""
    total = len(input_files)
    if total < 2:
        return 0, total, "Membutuhkan minimal 2 file PDF untuk digabungkan."

    target_dir = (
        os.path.abspath(output_directory)
        if output_directory
        else os.path.dirname(os.path.abspath(input_files[0]))
    )
    output_path = os.path.join(target_dir, "MiCO_Merged_Result.pdf")

    writer = PdfWriter()
    try:
        for idx, pdf_path in enumerate(input_files, start=1):
            if status_callback:
                status_callback(f"Menggabungkan ({idx}/{total}): {os.path.basename(pdf_path)}...")
            reader = PdfReader(pdf_path)
            for page in reader.pages:
                writer.add_page(page)

        with open(output_path, "wb") as f_out:
            writer.write(f_out)

        return 1, 1, ""
    except Exception as e:
        return 0, 1, str(e)


def split_pdf(input_files, output_directory="", status_callback=None):
    """Memisah tiap halaman PDF menjadi file tersendiri."""
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
            reader = PdfReader(pdf_path)
            total_pages = len(reader.pages)
            for page_idx, page in enumerate(reader.pages, start=1):
                if status_callback:
                    status_callback(f"Memisah {file_name} (Halaman {page_idx}/{total_pages})...")
                
                writer = PdfWriter()
                writer.add_page(page)
                out_page_path = os.path.join(target_dir, f"{file_title}_page_{page_idx}.pdf")
                with open(out_page_path, "wb") as f_out:
                    writer.write(f_out)

            success_count += 1
        except Exception as e:
            last_error = str(e)

    return success_count, total, last_error


def compress_pdf(input_files, output_directory="", status_callback=None):
    """Kompresi ukuran file PDF."""
    total = len(input_files)
    success_count = 0
    last_error = ""

    for idx, pdf_path in enumerate(input_files, start=1):
        pdf_path = os.path.abspath(pdf_path)
        file_name = os.path.basename(pdf_path)
        file_title, _ = os.path.splitext(file_name)

        if status_callback:
            status_callback(f"Mengompres ({idx}/{total}): {file_name}...")

        target_dir = (
            os.path.abspath(output_directory)
            if output_directory
            else os.path.dirname(pdf_path)
        )
        output_path = os.path.join(target_dir, f"{file_title}_compressed.pdf")

        try:
            reader = PdfReader(pdf_path)
            writer = PdfWriter()

            for page in reader.pages:
                page.compress_content_streams()  # Kompres stream halaman
                writer.add_page(page)

            with open(output_path, "wb") as f_out:
                writer.write(f_out)

            if os.path.exists(output_path):
                success_count += 1
        except Exception as e:
            last_error = str(e)

    return success_count, total, last_error