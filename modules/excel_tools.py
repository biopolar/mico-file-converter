import os

def convert_excel_to_pdf(file_paths, target_dir):
    try:
        import win32com.client # Import ditaruh di dalam fungsi agar tidak merusak macOS
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        success = 0
        for path in file_paths:
            if os.path.exists(path):
                wb = excel.Workbooks.Open(os.path.abspath(path))
                base = os.path.basename(path)
                name, _ = os.path.splitext(base)
                pdf_path = os.path.join(target_dir, f"{name}.pdf")
                wb.ExportAsFixedFormat(0, pdf_path)
                wb.Close(False)
                success += 1
        excel.Quit()
        return success, len(file_paths), ""
    except Exception as e:
        return 0, len(file_paths), f"Fitur Excel to PDF membutuhkan Windows & Microsoft Excel: {str(e)}"

def convert_pdf_to_excel(file_paths, target_dir):
    try:
        import fitz
        import openpyxl
        
        success = 0
        for pdf_path in file_paths:
            if not os.path.exists(pdf_path):
                continue
            doc = fitz.open(pdf_path)
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Extracted Data"
            
            row_idx = 1
            for page in doc:
                tabs = page.find_tables()
                if tabs and len(tabs.tables) > 0:
                    for table in tabs:
                        for row in table.extract():
                            ws.append(row)
                            row_idx += 1
                else:
                    text = page.get_text("text")
                    for line in text.split("\n"):
                        if line.strip():
                            ws.append([line.strip()])
                            row_idx += 1
                            
            base_name = os.path.basename(pdf_path)
            name, _ = os.path.splitext(base_name)
            out_path = os.path.join(target_dir, f"{name}.xlsx")
            wb.save(out_path)
            doc.close()
            success += 1
            
        return success, len(file_paths), ""
    except Exception as e:
        return 0, len(file_paths), f"Gagal konversi PDF ke Excel: {str(e)}"