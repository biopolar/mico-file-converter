import os
import pythoncom
import win32com.client


def convert_word_to_pdf(input_files, output_directory="", status_callback=None):
    """Fungsi independen untuk mengonversi daftar file Word (.docx) ke PDF.

    Returns:
        tuple: (success_count, total_files, last_error_message)
    """
    pythoncom.CoInitialize()

    total = len(input_files)
    success_count = 0
    last_error = ""

    def get_word_app():
        """Membuat instance Word terpisah & terisolasi khusus untuk konversi"""
        try:
            app = win32com.client.DispatchEx("Word.Application")
        except Exception:
            app = win32com.client.Dispatch("Word.Application")
        app.Visible = False
        app.DisplayAlerts = 0
        return app

    word_app = None

    try:
        word_app = get_word_app()

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
            output_path = os.path.join(target_dir, f"{file_title}_converted.pdf")

            doc = None
            try:
                # Mode ReadOnly agar tidak terpengaruh jika file ditutup/dibuka di MS Word
                doc = word_app.Documents.Open(
                    FileName=input_path,
                    ConfirmConversions=False,
                    ReadOnly=True,
                    AddToRecentFiles=False,
                )
                doc.SaveAs(output_path, FileFormat=17)  # 17 = wdFormatPDF

                if os.path.exists(output_path):
                    success_count += 1
            except Exception as e:
                last_error = str(e)
                # Auto recovery jika terjadi RPC disconnection
                try:
                    word_app = get_word_app()
                except Exception:
                    pass
            finally:
                if doc:
                    try:
                        doc.Close(SaveChanges=False)
                    except Exception:
                        pass

    except Exception as e:
        last_error = f"Gagal menginisialisasi MS Word: {e}"
    finally:
        if word_app:
            try:
                word_app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

    return success_count, total, last_error