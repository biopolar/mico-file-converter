import os
import sys
import shutil
import time
import base64
import tempfile
import subprocess
import webview

def get_resource_path(relative_path):
    """ Mendapatkan path absolut ke resource, kompatibel dengan PyInstaller bundle macOS & Windows """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

# Import fungsi dari modules secara TERPISAH & ISOLATED
try:
    from modules.word_to_pdf import convert_word_to_pdf
except Exception as e:
    def convert_word_to_pdf(*a, **k): return 0, 1, f"Fitur Word to PDF tidak tersedia: {e}"

try:
    from modules.excel_tools import convert_excel_to_pdf, convert_pdf_to_excel
except Exception as e:
    def convert_excel_to_pdf(*a, **k): return 0, 1, f"Error modul Excel: {e}"
    def convert_pdf_to_excel(*a, **k): return 0, 1, f"Error modul PDF to Excel: {e}"

try:
    from modules.ppt_tools import convert_ppt_to_pdf, convert_pdf_to_ppt
except Exception as e:
    def convert_ppt_to_pdf(*a, **k): return 0, 1, f"Fitur PPT to PDF tidak tersedia: {e}"
    def convert_pdf_to_ppt(*a, **k): return 0, 1, f"Error modul PDF to PPT: {e}"

try:
    from modules.pdf_to_word import convert_pdf_to_word
except Exception as e:
    def convert_pdf_to_word(*a, **k): return 0, 1, f"Error modul PDF to Word: {e}"

try:
    from modules.image_tools import convert_image_to_pdf, convert_pdf_to_image
except Exception as e:
    def convert_image_to_pdf(*a, **k): return 0, 1, f"Error modul Image: {e}"
    def convert_pdf_to_image(*a, **k): return 0, 1, f"Error modul PDF to Image: {e}"

try:
    from modules.pdf_tools import merge_pdfs, split_pdf, compress_pdf
except Exception as e:
    def merge_pdfs(*a, **k): return 0, 1, f"Error modul Merge PDF: {e}"
    def split_pdf(*a, **k): return 0, 1, f"Error modul Split PDF: {e}"
    def compress_pdf(*a, **k): return 0, 1, f"Error modul Compress PDF: {e}"

TOOLS_DATA = [
    {"id": "WORD_TO_PDF", "title": "Word to PDF", "desc": "Convert DOC/DOCX documents to PDF format.", "badge": "DOCX", "cat": "Convert to PDF", "ext": ("Word Files (*.docx;*.doc)", "All Files (*.*)"), "target_ext": ".pdf"},
    {"id": "EXCEL_TO_PDF", "title": "Excel to PDF", "desc": "Turn spreadsheets (XLSX/XLS) into PDF files.", "badge": "XLSX", "cat": "Convert to PDF", "ext": ("Excel Files (*.xlsx;*.xls)", "All Files (*.*)"), "target_ext": ".pdf"},
    {"id": "PPT_TO_PDF", "title": "PPT to PDF", "desc": "Export PowerPoint presentations (PPTX/PPT) to PDF.", "badge": "PPTX", "cat": "Convert to PDF", "ext": ("PowerPoint Files (*.pptx;*.ppt)", "All Files (*.*)"), "target_ext": ".pdf"},
    {"id": "PDF_TO_WORD", "title": "PDF to Word", "desc": "Extract editable Word documents from PDF files.", "badge": "PDF", "cat": "Convert from PDF", "ext": ("PDF Files (*.pdf)", "All Files (*.*)"), "target_ext": ".docx"},
    {"id": "PDF_TO_EXCEL", "title": "PDF to Excel", "desc": "Convert tables from PDF files into Excel spreadsheets.", "badge": "PDF", "cat": "Convert from PDF", "ext": ("PDF Files (*.pdf)", "All Files (*.*)"), "target_ext": ".xlsx"},
    {"id": "PDF_TO_PPT", "title": "PDF to PPT", "desc": "Convert PDF slides back into editable Presentations.", "badge": "PDF", "cat": "Convert from PDF", "ext": ("PDF Files (*.pdf)", "All Files (*.*)"), "target_ext": ".pptx"},
    {"id": "MERGE_PDF", "title": "Merge PDF", "desc": "Combine multiple PDF files into a single document.", "badge": "PDF", "cat": "PDF Tools", "ext": ("PDF Files (*.pdf)", "All Files (*.*)"), "target_ext": ".pdf"},
    {"id": "SPLIT_PDF", "title": "Split PDF", "desc": "Separate a PDF into individual pages or custom ranges.", "badge": "PDF", "cat": "PDF Tools", "ext": ("PDF Files (*.pdf)", "All Files (*.*)"), "target_ext": ".pdf"},
    {"id": "COMPRESS_PDF", "title": "Compress PDF", "desc": "Reduce PDF file size without sacrificing quality.", "badge": "PDF", "cat": "PDF Tools", "ext": ("PDF Files (*.pdf)", "All Files (*.*)"), "target_ext": ".pdf"},
    {"id": "IMAGE_TO_PDF", "title": "JPG to PDF", "desc": "Package one or more JPG/PNG images into a PDF.", "badge": "JPG", "cat": "Convert to PDF", "ext": ("Image Files (*.jpg;*.jpeg;*.png)", "All Files (*.*)"), "target_ext": ".pdf"},
    {"id": "PDF_TO_IMAGE", "title": "PDF to JPG", "desc": "Extract high-resolution images from each page of PDF.", "badge": "PDF", "cat": "Convert from PDF", "ext": ("PDF Files (*.pdf)", "All Files (*.*)"), "target_ext": ".jpg"},
    {"id": "PROTECT_PDF", "title": "Protect PDF", "desc": "Add password protection and permissions to PDF.", "badge": "PDF", "cat": "PDF Tools", "ext": ("PDF Files (*.pdf)", "All Files (*.*)"), "target_ext": ".pdf"},
    {"id": "UNLOCK_PDF", "title": "Unlock PDF", "desc": "Remove password protection from PDF files.", "badge": "PDF", "cat": "PDF Tools", "ext": ("PDF Files (*.pdf)", "All Files (*.*)"), "target_ext": ".pdf"},
]

TEMP_DROP_DIR = os.path.join(tempfile.gettempdir(), "MiCO_Temp")
os.makedirs(TEMP_DROP_DIR, exist_ok=True)

class Api:
    def get_tools(self):
        return TOOLS_DATA

    def open_file_dialog(self, tool_id):
        try:
            tool = next((t for t in TOOLS_DATA if t["id"] == tool_id), None)
            file_types = tool["ext"] if tool and "ext" in tool else ("All Files (*.*)",)
            
            window = webview.active_window() or (webview.windows[0] if webview.windows else None)
            if not window:
                return []

            dialog_type = getattr(webview, 'FileDialog', None)
            open_type = dialog_type.OPEN if dialog_type else getattr(webview, 'OPEN_DIALOG', 10)

            result = window.create_file_dialog(
                open_type, 
                allow_multiple=True,
                file_types=file_types
            )

            if result:
                file_list = []
                for path in result:
                    file_list.append({
                        "path": path,
                        "name": os.path.basename(path),
                        "size": round(os.path.getsize(path) / (1024 * 1024), 2) if os.path.exists(path) else 0
                    })
                return file_list
            return []
        except Exception as e:
            print(f"Error opening file dialog: {e}")
            return []

    def save_dropped_file(self, filename, base64_data):
        try:
            file_bytes = base64.b64decode(base64_data)
            out_path = os.path.join(TEMP_DROP_DIR, filename)
            with open(out_path, "wb") as f:
                f.write(file_bytes)
            return {
                "path": out_path,
                "name": filename,
                "size": round(len(file_bytes) / (1024 * 1024), 2)
            }
        except Exception as e:
            print(f"Error saving dropped file: {e}")
            return None

    def select_save_location(self):
        try:
            window = webview.active_window() or (webview.windows[0] if webview.windows else None)
            if not window:
                return None
            dialog_type = getattr(webview, 'FileDialog', None)
            folder_type = dialog_type.FOLDER if dialog_type else getattr(webview, 'FOLDER_DIALOG', 20)

            result = window.create_file_dialog(folder_type)
            if result and len(result) > 0:
                return result[0]
            return None
        except Exception as e:
            print(f"Error selecting folder: {e}")
            return None

    def execute_conversion(self, tool_id, file_paths, target_dir, custom_filename="", password=""):
        if not target_dir or not os.path.exists(target_dir):
            return {"status": "error", "message": "Folder penyimpanan tidak valid."}

        start_time = time.time()
        existing_files = set(os.listdir(target_dir)) if os.path.exists(target_dir) else set()

        try:
            success, total, err = 0, len(file_paths), ""

            if tool_id == "WORD_TO_PDF":
                success, total, err = convert_word_to_pdf(file_paths, target_dir)
            elif tool_id == "EXCEL_TO_PDF":
                success, total, err = convert_excel_to_pdf(file_paths, target_dir)
            elif tool_id == "PPT_TO_PDF":
                success, total, err = convert_ppt_to_pdf(file_paths, target_dir)
            elif tool_id == "PDF_TO_WORD":
                success, total, err = convert_pdf_to_word(file_paths, target_dir)
            elif tool_id == "PDF_TO_EXCEL":
                success, total, err = convert_pdf_to_excel(file_paths, target_dir)
            elif tool_id == "PDF_TO_PPT":
                success, total, err = convert_pdf_to_ppt(file_paths, target_dir)
            elif tool_id == "MERGE_PDF":
                success, total, err = merge_pdfs(file_paths, target_dir)
            elif tool_id == "SPLIT_PDF":
                success, total, err = split_pdf(file_paths, target_dir)
            elif tool_id == "IMAGE_TO_PDF":
                success, total, err = convert_image_to_pdf(file_paths, target_dir)
            elif tool_id == "PDF_TO_IMAGE":
                success, total, err = convert_pdf_to_image(file_paths, target_dir)
            elif tool_id == "COMPRESS_PDF":
                success, total, err = compress_pdf(file_paths, target_dir)
            elif tool_id == "PROTECT_PDF":
                return self.protect_pdf(file_paths, target_dir, password, custom_filename)
            elif tool_id == "UNLOCK_PDF":
                return self.unlock_pdf(file_paths, target_dir, password, custom_filename)
            else:
                return {"status": "error", "message": f"Fitur {tool_id} belum dikonfigurasi."}

            if success == 0 and err:
                return {"status": "error", "message": err}

            new_files = list(set(os.listdir(target_dir)) - existing_files)
            final_file_path = target_dir

            if new_files:
                if custom_filename and len(new_files) == 1:
                    created_file = new_files[0]
                    created_path = os.path.join(target_dir, created_file)
                    
                    _, ext = os.path.splitext(created_file)
                    if not custom_filename.lower().endswith(ext.lower()):
                        custom_filename += ext
                    
                    final_path = os.path.join(target_dir, custom_filename)
                    if created_path != final_path:
                        if os.path.exists(final_path):
                            os.remove(final_path)
                        os.rename(created_path, final_path)
                    final_file_path = final_path
                else:
                    final_file_path = os.path.join(target_dir, new_files[0])

            self.open_and_select_file(final_file_path)

            elapsed = round(time.time() - start_time, 1)
            return {
                "status": "success",
                "tool_name": tool_id.replace("_", " "),
                "processing_time": f"{elapsed}s",
                "size_savings": f"{success}/{total} File Berhasil",
                "out_file": final_file_path
            }
        except Exception as e:
            return {"status": "error", "message": str(e)}

    def protect_pdf(self, file_paths, target_dir, password, custom_filename=""):
        try:
            from pypdf import PdfReader, PdfWriter
            out_path = target_dir
            for file_path in file_paths:
                reader = PdfReader(file_path)
                writer = PdfWriter()
                for page in reader.pages:
                    writer.add_page(page)
                writer.encrypt(password)
                
                base_name = os.path.basename(file_path)
                name, _ = os.path.splitext(base_name)
                
                out_name = custom_filename if (custom_filename and len(file_paths) == 1) else f"{name}_protected.pdf"
                if not out_name.endswith(".pdf"):
                    out_name += ".pdf"
                    
                out_path = os.path.join(target_dir, out_name)
                with open(out_path, "wb") as f:
                    writer.write(f)
                    
            self.open_and_select_file(out_path)
            return {"status": "success", "out_file": out_path, "processing_time": "1s", "size_savings": f"{len(file_paths)} File Protected"}
        except Exception as e:
            return {"status": "error", "message": str(e)}

    def unlock_pdf(self, file_paths, target_dir, password, custom_filename=""):
        try:
            from pypdf import PdfReader, PdfWriter
            out_path = target_dir
            for file_path in file_paths:
                reader = PdfReader(file_path)
                if reader.is_encrypted:
                    if reader.decrypt(password) == 0:
                        return {"status": "error", "message": "Password salah!"}
                writer = PdfWriter()
                for page in reader.pages:
                    writer.add_page(page)
                    
                base_name = os.path.basename(file_path)
                name, _ = os.path.splitext(base_name)
                
                out_name = custom_filename if (custom_filename and len(file_paths) == 1) else f"{name}_unlocked.pdf"
                if not out_name.endswith(".pdf"):
                    out_name += ".pdf"
                    
                out_path = os.path.join(target_dir, out_name)
                with open(out_path, "wb") as f:
                    writer.write(f)
                    
            self.open_and_select_file(out_path)
            return {"status": "success", "out_file": out_path, "processing_time": "1s", "size_savings": f"{len(file_paths)} File Unlocked"}
        except Exception as e:
            return {"status": "error", "message": str(e)}

    def open_and_select_file(self, file_path):
        """ Membuka File Explorer / Finder lintas platform (Windows & macOS) """
        try:
            if file_path and os.path.exists(file_path):
                if sys.platform == "win32":
                    if os.path.isfile(file_path):
                        subprocess.Popen(['explorer', '/select,', os.path.normpath(file_path)])
                    else:
                        os.startfile(file_path)
                elif sys.platform == "darwin": # macOS Finder
                    if os.path.isfile(file_path):
                        subprocess.Popen(['open', '-R', file_path])
                    else:
                        subprocess.Popen(['open', file_path])
        except Exception as e:
            print(f"Error opening file explorer: {e}")

if __name__ == '__main__':
    api = Api()
    html_path = get_resource_path(os.path.join('web', 'index.html'))

    window = webview.create_window(
        'MiCO File Converter',
        html_path,
        js_api=api,
        width=1180,
        height=780,
        resizable=True
    )

    webview.start()