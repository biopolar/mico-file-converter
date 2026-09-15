import os
import time
import threading
from tkinter import messagebox

# Import modul konversi dari folder modules
from modules.word_to_pdf import convert_word_to_pdf
from modules.pdf_to_word import convert_pdf_to_word
from modules.excel_tools import convert_excel_to_pdf, convert_pdf_to_excel
from modules.image_tools import convert_image_to_pdf, convert_pdf_to_image, convert_image_to_word
from modules.pdf_tools import merge_pdfs, split_pdf, compress_pdf
from modules.ppt_tools import convert_ppt_to_pdf, convert_pdf_to_ppt


class AppController:
    def __init__(self, view):
        self.view = view
        self.current_tool = None
        self.files = []

        # Lokasi Output di Documents User (Aman dari WinError 5 Access Denied)
        self.output_dir = os.path.join(os.path.expanduser("~"), "Documents", "MiCO_Output")
        os.makedirs(self.output_dir, exist_ok=True)

    def select_tool(self, tool):
        self.current_tool = tool

    def add_files(self, files):
        self.files = files

    def remove_file(self, path):
        if path in self.files:
            self.files.remove(path)

    def clear_all_files(self):
        self.files = []

    def start_conversion(self):
        if not self.current_tool or not self.files:
            messagebox.showwarning("Peringatan", "Pilih file terlebih dahulu!")
            return

        self.view.set_converting_state(True)

        # Jalankan proses konversi di background thread
        thread = threading.Thread(target=self._process_conversion)
        thread.daemon = True
        thread.start()

    def _show_error_dialog(self, message):
        messagebox.showerror("Error", message)

    def _show_success_dialog(self, summary):
        self.view.show_success_view(summary)

    def _process_conversion(self):
        start_time = time.time()
        tool_id = self.current_tool.get("id")
        os.makedirs(self.output_dir, exist_ok=True)

        success_count = 0
        total = len(self.files)
        last_error = ""

        try:
            # Peta eksekusi modul berdasarkan id tool
            if tool_id == "WORD_TO_PDF":
                success_count, total, last_error = convert_word_to_pdf(self.files, self.output_dir)
            elif tool_id == "PDF_TO_WORD":
                success_count, total, last_error = convert_pdf_to_word(self.files, self.output_dir)
            elif tool_id == "EXCEL_TO_PDF":
                success_count, total, last_error = convert_excel_to_pdf(self.files, self.output_dir)
            elif tool_id == "PDF_TO_EXCEL":
                success_count, total, last_error = convert_pdf_to_excel(self.files, self.output_dir)
            elif tool_id == "PPT_TO_PDF":
                success_count, total, last_error = convert_ppt_to_pdf(self.files, self.output_dir)
            elif tool_id == "PDF_TO_PPT":
                success_count, total, last_error = convert_pdf_to_ppt(self.files, self.output_dir)
            elif tool_id == "IMAGE_TO_PDF":
                success_count, total, last_error = convert_image_to_pdf(self.files, self.output_dir)
            elif tool_id == "PDF_TO_IMAGE":
                success_count, total, last_error = convert_pdf_to_image(self.files, self.output_dir)
            elif tool_id == "IMAGE_TO_WORD":
                success_count, total, last_error = convert_image_to_word(self.files, self.output_dir)
            elif tool_id == "MERGE_PDF":
                success_count, total, last_error = merge_pdfs(self.files, self.output_dir)
            elif tool_id == "SPLIT_PDF":
                success_count, total, last_error = split_pdf(self.files, self.output_dir)
            elif tool_id == "COMPRESS_PDF":
                success_count, total, last_error = compress_pdf(self.files, self.output_dir)
            else:
                last_error = f"Fitur '{tool_id}' belum didukung."

            elapsed_time = round(time.time() - start_time, 1)

            if success_count > 0:
                summary = {
                    "tool_name": self.current_tool.get("title", "Tool"),
                    "output_format": self.current_tool.get("badge", "PDF"),
                    "size_savings": "Optimal",
                    "processing_time": f"{elapsed_time}s",
                    "status": "Success",
                    "out_dir": self.output_dir
                }
                self.view.root.after(0, self._show_success_dialog, summary)
            else:
                err_msg = last_error if last_error else "Konversi gagal dijalankan."
                self.view.root.after(0, self._show_error_dialog, f"Gagal konversi: {err_msg}")

        except Exception as e:
            err_msg = str(e)
            self.view.root.after(0, self._show_error_dialog, f"Terjadi kesalahan: {err_msg}")
        finally:
            self.view.root.after(0, self.view.set_converting_state, False)