import os
import sys
import subprocess

def convert_word_to_pdf(file_paths, target_dir):
    total = len(file_paths)
    if total == 0:
        return 0, 0, "Tidak ada file yang dipilih."
    
    success = 0

    # --- SISTEM OPERASI WINDOWS ---
    if sys.platform == "win32":
        word = None
        co_inited = False
        try:
            import win32com.client
            try:
                import pythoncom
                pythoncom.CoInitialize()
                co_inited = True
            except Exception:
                pass

            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False

            for docx_path in file_paths:
                if not os.path.exists(docx_path):
                    continue
                abs_docx = os.path.abspath(docx_path)
                base_name = os.path.basename(docx_path)
                name, _ = os.path.splitext(base_name)
                pdf_path = os.path.abspath(os.path.join(target_dir, f"{name}.pdf"))
                
                doc = None
                try:
                    doc = word.Documents.Open(abs_docx)
                    doc.SaveAs(pdf_path, FileFormat=17)  # 17 = wdFormatPDF
                    success += 1
                except Exception as doc_err:
                    print(f"Error konversi file {docx_path}: {doc_err}")
                finally:
                    if doc:
                        try:
                            doc.Close(0)  # 0 = wdDoNotSaveChanges
                        except Exception:
                            pass

            return success, total, "" if success > 0 else "Gagal mengonversi file Word ke PDF."

        except Exception as e:
            return 0, total, f"Gagal konversi Word to PDF di Windows: {str(e)}"
        finally:
            if word:
                try:
                    word.Quit()
                except Exception:
                    pass
            if co_inited:
                try:
                    import pythoncom
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    # --- SISTEM OPERASI MACOS ---
    elif sys.platform == "darwin":
        for docx_path in file_paths:
            if not os.path.exists(docx_path):
                continue
            abs_docx = os.path.abspath(docx_path)
            base_name = os.path.basename(docx_path)
            name, _ = os.path.splitext(base_name)
            pdf_path = os.path.abspath(os.path.join(target_dir, f"{name}.pdf"))

            # Escaping tanda petik ganda agar aman di AppleScript
            safe_docx = abs_docx.replace('"', '\\"')
            safe_pdf = pdf_path.replace('"', '\\"')

            converted = False

            # Opsi 1: Microsoft Word Mac via AppleScript
            applescript = f'''
            tell application "Microsoft Word"
                set visibility to false
                open POSIX file "{safe_docx}"
                set theDoc to active document
                save as theDoc file format format PDF file name POSIX file "{safe_pdf}"
                close theDoc saving no
            end tell
            '''
            try:
                res = subprocess.run(["osascript", "-e", applescript], capture_output=True, text=True, timeout=60)
                if res.returncode == 0 and os.path.exists(pdf_path):
                    success += 1
                    converted = True
            except Exception:
                pass

            if converted:
                continue

            # Opsi 2: Fallback LibreOffice jika MS Word tidak ada di Mac
            try:
                lo_res = subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", target_dir, abs_docx], capture_output=True, timeout=60)
                if lo_res.returncode == 0 and os.path.exists(pdf_path):
                    success += 1
                    converted = True
            except Exception:
                pass

        if success > 0:
            return success, total, ""
        else:
            return 0, total, "Konversi Word to PDF di Mac membutuhkan Microsoft Word atau LibreOffice yang terinstall."

    else:
        return 0, total, "Sistem operasi tidak didukung."