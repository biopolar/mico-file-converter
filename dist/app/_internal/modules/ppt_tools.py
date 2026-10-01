import os

def convert_ppt_to_pdf(files, out_dir, progress_callback=None):
    success_count = 0
    error_msg = ""
    for file_path in files:
        try:
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            output_path = os.path.join(out_dir, f"{base_name}.pdf")
            
            # Tempatkan logika library konversi PPT ke PDF di sini
            success_count += 1
        except Exception as e:
            error_msg = str(e)
            
    return success_count, len(files), error_msg

def convert_pdf_to_ppt(files, out_dir, progress_callback=None):
    success_count = 0
    error_msg = ""
    for file_path in files:
        try:
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            output_path = os.path.join(out_dir, f"{base_name}.pptx")
            
            # Tempatkan logika library konversi PDF ke PPT di sini
            success_count += 1
        except Exception as e:
            error_msg = str(e)
            
    return success_count, len(files), error_msg