import os
import win32com.client

def convert_word_to_pdf(file_paths, output_dir):
    # Membuka Word Application 1x di latar belakang
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    
    success_count = 0
    err_msg = ""

    try:
        for file_path in file_paths:
            try:
                doc = word.Documents.Open(os.path.abspath(file_path))
                base_name = os.path.splitext(os.path.basename(file_path))[0]
                out_path = os.path.join(output_dir, f"{base_name}.pdf")
                
                # 17 = wdFormatPDF
                doc.SaveAs(out_path, FileFormat=17)
                doc.Close()
                success_count += 1
            except Exception as e:
                err_msg = str(e)
    finally:
        word.Quit()

    return success_count, len(file_paths), err_msg