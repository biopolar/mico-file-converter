import os
import io
import shutil
import fitz  # PyMuPDF
from PIL import Image
from pypdf import PdfReader, PdfWriter


def merge_pdfs(file_paths, output_dir):
    try:
        writer = PdfWriter()
        for file_path in file_paths:
            reader = PdfReader(file_path)
            for page in reader.pages:
                writer.add_page(page)

        out_path = os.path.join(output_dir, "Merged_Document.pdf")
        with open(out_path, "wb") as f:
            writer.write(f)

        return 1, 1, ""
    except Exception as e:
        return 0, len(file_paths), str(e)


def split_pdf(file_paths, output_dir):
    success_count = 0
    try:
        for file_path in file_paths:
            reader = PdfReader(file_path)
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            
            for i, page in enumerate(reader.pages):
                writer = PdfWriter()
                writer.add_page(page)
                out_path = os.path.join(output_dir, f"{base_name}_page_{i+1}.pdf")
                with open(out_path, "wb") as f:
                    writer.write(f)
            
            success_count += 1
        return success_count, len(file_paths), ""
    except Exception as e:
        return success_count, len(file_paths), str(e)


def compress_pdf(file_paths, output_dir, image_quality=65, max_dimension=1800):
    """
    Kompresi Universal PDF dengan Garansi Ukuran File:
    - Tidak memproses gambar yang ukurannya sudah sangat kecil (< 15 KB).
    - Hanya mengganti gambar jika byte hasil kompresi terbukti lebih kecil.
    - Garansi: Jika file hasil kompresi lebih besar dari file awal, otomatis gunakan file awal.
    """
    success_count = 0
    err_msg = ""

    for file_path in file_paths:
        try:
            original_size = os.path.getsize(file_path)
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            out_path = os.path.join(output_dir, f"{base_name}_compressed.pdf")

            doc = fitz.open(file_path)
            
            # 1. Optimalisasi Gambar Internal
            for page in doc:
                image_list = page.get_images(full=True)
                for img in image_list:
                    xref = img[0]
                    try:
                        base_image = doc.extract_image(xref)
                        image_bytes = base_image["image"]
                        
                        # Lewati gambar ikon/elemen kecil yang sudah < 15 KB
                        if len(image_bytes) < 15 * 1024:
                            continue

                        pil_img = Image.open(io.BytesIO(image_bytes))
                        
                        if pil_img.mode in ("RGBA", "P", "LA"):
                            pil_img = pil_img.convert("RGB")
                        
                        # Downscale jika ukuran piksel lebih dari max_dimension
                        w, h = pil_img.size
                        if w > max_dimension or h > max_dimension:
                            pil_img.thumbnail((max_dimension, max_dimension), Image.Resampling.LANCZOS)
                        
                        buffer = io.BytesIO()
                        pil_img.save(buffer, format="JPEG", quality=image_quality, optimize=True)
                        compressed_bytes = buffer.getvalue()
                        
                        # Hanya ganti jika ukuran gambar baru LEBIH KECIL
                        if len(compressed_bytes) < len(image_bytes):
                            page.replace_image(xref, stream=compressed_bytes)
                    except Exception:
                        continue

            # 2. Simpan dengan Kompresi Struktur PDF
            doc.save(
                out_path,
                garbage=4,
                deflate=True,
                deflate_images=True,
                deflate_fonts=True
            )
            doc.close()

            # 3. GARANSI UKURAN: Jika hasil kompresi justru membengkak, pakai file asli
            if os.path.exists(out_path) and os.path.getsize(out_path) >= original_size:
                shutil.copyfile(file_path, out_path)

            success_count += 1
        except Exception as e:
            err_msg = str(e)

    return success_count, len(file_paths), err_msg