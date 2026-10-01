import os
import pythoncom
import win32com.client
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


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
    """Mengonversi seluruh isi PDF ke Excel dengan struktur & styling yang rapi."""
    total = len(input_files)
    success_count = 0
    last_error = ""

    # Styling Excel (menyesuaikan tema Photo 1)
    title_font = Font(name="Calibri", size=14, bold=True, color="1B365D")
    subtitle_font = Font(name="Calibri", size=10, italic=True, color="595959")
    section_font = Font(name="Calibri", size=11, bold=True, color="1B365D")
    bold_font = Font(name="Calibri", size=10, bold=True)
    regular_font = Font(name="Calibri", size=10)

    header_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    total_border = Border(
        top=Side(style="thin", color="000000"),
        bottom=Side(style="double", color="000000"),
    )

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
                    ws.views.sheetView[0].showGridLines = True

                    # 1. Cari area tabel di halaman
                    tables = page.find_tables()
                    table_bboxes = [t.bbox for t in tables]

                    def is_inside_table(bbox):
                        for t_bbox in table_bboxes:
                            if not (
                                bbox[2] < t_bbox[0]
                                or bbox[0] > t_bbox[2]
                                or bbox[3] < t_bbox[1]
                                or bbox[1] > t_bbox[3]
                            ):
                                return True
                        return False

                    # 2. Ambil teks di luar tabel (Judul, Subjudul, Paragraf)
                    elements = []
                    lines = page.extract_text_lines()
                    for line in lines:
                        line_bbox = (line["x0"], line["top"], line["x1"], line["bottom"])
                        if not is_inside_table(line_bbox):
                            elements.append(
                                {
                                    "top": line["top"],
                                    "type": "text",
                                    "text": line["text"].strip(),
                                }
                            )

                    # 3. Masukkan data tabel
                    for t in tables:
                        elements.append(
                            {
                                "top": t.bbox[1],
                                "type": "table",
                                "data": t.extract(),
                            }
                        )

                    # 4. Urutkan elemen dari atas ke bawah sesuai posisi fisik PDF
                    elements = sorted(elements, key=lambda x: x["top"])

                    current_row = 1
                    for elem in elements:
                        if elem["type"] == "text":
                            txt = elem["text"]
                            if not txt:
                                continue

                            cell = ws.cell(row=current_row, column=1, value=txt)

                            # Styling judul / subjudul berdasarkan konteks teks
                            if current_row == 1 or "LAPORAN" in txt.upper():
                                cell.font = title_font
                            elif "PT " in txt or "MANDIRI" in txt.upper():
                                cell.font = subtitle_font
                            elif txt.isupper() and len(txt) < 40:
                                cell.font = section_font
                            else:
                                cell.font = regular_font

                            current_row += 1

                        elif elem["type"] == "table":
                            table_data = elem["data"]
                            if not table_data:
                                continue

                            # Pemisah 1 baris sebelum tabel
                            if current_row > 1 and ws.cell(row=current_row - 1, column=1).value:
                                current_row += 1

                            for r_idx, row in enumerate(table_data):
                                is_header = r_idx == 0
                                is_total_row = any("TOTAL" in str(c).upper() for c in row if c)

                                for c_idx, val in enumerate(row, start=1):
                                    cell_val = val.strip() if val else ""
                                    cell = ws.cell(row=current_row, column=c_idx, value=cell_val)

                                    # Formatting cell tabel
                                    if is_header:
                                        cell.fill = header_fill
                                        cell.font = header_font
                                        cell.alignment = Alignment(
                                            horizontal="center" if c_idx > 1 else "left",
                                            vertical="center",
                                        )
                                    elif is_total_row:
                                        cell.font = bold_font
                                        cell.border = total_border
                                        cell.alignment = Alignment(
                                            horizontal="right" if c_idx > 1 else "left"
                                        )
                                    else:
                                        cell.font = regular_font
                                        cell.border = thin_border
                                        # Perataan angka ke kanan, teks ke kiri
                                        if c_idx > 1 and (
                                            "Rp" in cell_val
                                            or cell_val.replace(".", "").replace(",", "").isdigit()
                                        ):
                                            cell.alignment = Alignment(horizontal="right")
                                        else:
                                            cell.alignment = Alignment(horizontal="left")

                                current_row += 1

                            current_row += 1  # Pemisah setelah tabel

                    # 5. Atur lebar kolom secara otomatis (Auto-fit)
                    for col in ws.columns:
                        max_len = 0
                        col_letter = get_column_letter(col[0].column)
                        for cell in col:
                            if cell.value:
                                val_str = str(cell.value)
                                # Batasi panjang dari teks judul agar tidak merusak lebar kolom A
                                if len(val_str) < 50:
                                    max_len = max(max_len, len(val_str))
                        ws.column_dimensions[col_letter].width = max(max_len + 5, 15)

            wb.save(output_path)
            if os.path.exists(output_path):
                success_count += 1
        except Exception as e:
            last_error = str(e)

    return success_count, total, last_error