import os
import sys
import subprocess

def convert_word_to_pdf(file_paths, target_dir):
    total = len(file_paths)
    if total == 0:
        return 0, 0, "Tidak ada file yang dipilih."
    
    if not os.path.exists(target_dir):
        os.makedirs(target_dir, exist_ok=True)

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
        # Jalur resmi executable LibreOffice jika terinstall di macOS
        libreoffice_path = "soffice"
        mac_lo_app = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
        if os.path.exists(mac_lo_app):
            libreoffice_path = mac_lo_app

        for docx_path in file_paths:
            if not os.path.exists(docx_path):
                continue
            abs_docx = os.path.abspath(docx_path)
            base_name = os.path.basename(docx_path)
            name, _ = os.path.splitext(base_name)
            pdf_path = os.path.abspath(os.path.join(target_dir, f"{name}.pdf"))

            safe_docx = abs_docx.replace('"', '\\"')
            safe_pdf = pdf_path.replace('"', '\\"')

            converted = False

            # Opsi 1: Microsoft Word Mac via AppleScript
            applescript_word = f'''
            tell application "Microsoft Word"
                set display alerts to false
                open (POSIX file "{safe_docx}")
                save as active document file name "{safe_pdf}" file format format PDF
                close active document saving no
            end tell
            '''
            try:
                res = subprocess.run(["osascript", "-e", applescript_word], capture_output=True, text=True, timeout=60)
                if res.returncode == 0 and os.path.exists(pdf_path):
                    success += 1
                    converted = True
            except Exception:
                pass

            if converted:
                continue

            # Opsi 2: Apple Pages (Bawaan macOS) via AppleScript
            applescript_pages = f'''
            tell application "Pages"
                set theDoc to open (POSIX file "{safe_docx}")
                export theDoc to (POSIX file "{safe_pdf}") as PDF
                close theDoc saving no
            end tell
            '''
            try:
                res = subprocess.run(["osascript", "-e", applescript_pages], capture_output=True, text=True, timeout=60)
                if res.returncode == 0 and os.path.exists(pdf_path):
                    success += 1
                    converted = True
            except Exception:
                pass

            if converted:
                continue

            # Opsi 3: LibreOffice di macOS
            try:
                lo_res = subprocess.run([libreoffice_path, "--headless", "--convert-to", "pdf", "--outdir", target_dir, abs_docx], capture_output=True, timeout=60)
                if lo_res.returncode == 0 and os.path.exists(pdf_path):
                    success += 1
                    converted = True
            except Exception:
                pass

        if success > 0:
            return success, total, ""
        else:
            return 0, total, "Gagal mengonversi file Word ke PDF di Mac. Pastikan Microsoft Word atau Apple Pages dapat dibuka."

    else:
        return 0, total, "Sistem operasi tidak didukung."   