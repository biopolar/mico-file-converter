import os
import re
import pythoncom
import win32com.client
import pdfplumber
import openpyxl


def convert_excel_to_pdf(input_files, output_directory="", status_callback=None):
    """Mengonversi file Excel (.xlsx / .xls) ke PDF menggunakan MS Excel API."""
    pythoncom.CoInitialize()

    total = len(input_files)
    success_count = 0
    last_error = ""

    excel_app = None
    try:
        try:
            excel_app = win32com.client.DispatchEx("Excel.Application")
        except Exception:
            excel_app = win32com.client.Dispatch("Excel.Application")
        
        excel_app.Visible = False
        excel_app.DisplayAlerts = False

        for idx, input_path in enumerate(input_files, start=1):
            input_path = os.path.abspath(input_path)
            file_name = os.path.basename(input_path)
            file_title, _ = os.path.splitext(file_name)

            if status_callback:
                status_callback(f"Memproses Excel ({idx}/{total}): {file_name}...")

            target_dir = (
                os.path.abspath(output_directory)
                if output_directory
                else os.path.dirname(input_path)
            )
            output_path = os.path.join(target_dir, f"{file_title}_converted.pdf")

            wb = None
            try:
                wb = excel_app.Workbooks.Open(input_path, ReadOnly=True)
                # 0 = xlTypePDF
                wb.ExportAsFixedFormat(0, output_path)
                if os.path.exists(output_path):
                    success_count += 1
            except Exception as e:
                last_error = str(e)
            finally:
                if wb:
                    wb.Close(SaveChanges=False)

    except Exception as e:
        last_error = f"Gagal menginisialisasi MS Excel: {e}"
    finally:
        if excel_app:
            try:
                excel_app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

    return success_count, total, last_error


def convert_pdf_to_excel(input_files, output_directory="", status_callback=None):
    """Mengonversi seluruh isi file PDF (teks & tabel utuh) ke file Excel (.xlsx)."""
    total = len(input_files)
    success_count = 0
    last_error = ""

    for idx, input_path in enumerate(input_files, start=1):
        input_path = os.path.abspath(input_path)
        file_name = os.path.basename(input_path)
        file_title, _ = os.path.splitext(file_name)

        if status_callback:
            status_callback(f"Memproses PDF ({idx}/{total}): {file_name}...")

        target_dir = (
            os.path.abspath(output_directory)
            if output_directory
            else os.path.dirname(input_path)
        )
        output_path = os.path.join(target_dir, f"{file_title}_converted.xlsx")

        try:
            wb = openpyxl.Workbook()
            wb.remove(wb.active)  # Hapus sheet default

            with pdfplumber.open(input_path) as pdf:
                for p_idx, page in enumerate(pdf.pages, start=1):
                    ws = wb.create_sheet(title=f"Page_{p_idx}")
                    
                    # Membaca seluruh isi halaman dengan mempertahankan tata letak posisi
                    text = page.extract_text(layout=True)

                    if text:
                        for line in text.split("\n"):
                            line_clean = line.rstrip()
                            if not line_clean:
                                continue
                            
                            # Memisahkan baris menjadi kolom berdasarkan spasi ganda/jarak antar teks
                            columns = [col.strip() for col in re.split(r'\s{2,}', line_clean) if col.strip()]
                            
                            if columns:
                                ws.append(columns)

            wb.save(output_path)
            if os.path.exists(output_path):
                success_count += 1
        except Exception as e:
            last_error = str(e)

    return success_count, total, last_error