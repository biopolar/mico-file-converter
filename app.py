import warnings
warnings.filterwarnings("ignore")

import sys
import types
import tkinter as tk
import ctypes
import os
import shutil
import time
import subprocess
import customtkinter as ctk
from tkinter import filedialog, messagebox

def get_resource_path(relative_path):
    """Mendapatkan path absolut untuk resource, berlaku saat run script atau PyInstaller exe."""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

def open_folder(path):
    try:
        if sys.platform == "win32":
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.call(["open", path])
        else:
            subprocess.call(["xdg-open", path])
    except Exception as e:
        print(f"Gagal membuka folder: {e}")

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDpiAware()
    except Exception:
        pass

if not hasattr(tk, "tix"):
    dummy_tix = types.ModuleType("tix")
    dummy_tix.Tk = tk.Tk
    sys.modules["tkinter.tix"] = dummy_tix
    sys.modules["Tix"] = dummy_tix
    setattr(tk, "tix", dummy_tix)

from tkinterdnd2 import DND_FILES, TkinterDnD

try:
    from controllers.app_controller import AppController
except Exception as e:
    print(f"Error loading AppController: {e}")
    AppController = None

ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

TOOL_FILTER_MAP = {
    "WORD_TO_PDF": ([("Word Documents", "*.docx *.doc")], [".docx", ".doc"]),
    "EXCEL_TO_PDF": ([("Excel Spreadsheets", "*.xlsx *.xls *.csv")], [".xlsx", ".xls", ".csv"]),
    "PPT_TO_PDF": ([("PowerPoint Presentations", "*.pptx *.ppt")], [".pptx", ".ppt"]),
    "PDF_TO_WORD": ([("PDF Documents", "*.pdf")], [".pdf"]),
    "PDF_TO_EXCEL": ([("PDF Documents", "*.pdf")], [".pdf"]),
    "PDF_TO_PPT": ([("PDF Documents", "*.pdf")], [".pdf"]),
    "MERGE_PDF": ([("PDF Documents", "*.pdf")], [".pdf"]),
    "SPLIT_PDF": ([("PDF Documents", "*.pdf")], [".pdf"]),
    "COMPRESS_PDF": ([("PDF Documents", "*.pdf")], [".pdf"]),
    "IMAGE_TO_PDF": ([("Image Files", "*.jpg *.jpeg *.png *.bmp *.webp")], [".jpg", ".jpeg", ".png", ".bmp", ".webp"]),
    "PDF_TO_IMAGE": ([("PDF Documents", "*.pdf")], [".pdf"]),
    "IMAGE_TO_WORD": ([("Image Files", "*.jpg *.jpeg *.png *.bmp *.webp")], [".jpg", ".jpeg", ".png", ".bmp", ".webp"]),
    "PROTECT_PDF": ([("PDF Documents", "*.pdf")], [".pdf"]),
    "UNLOCK_PDF": ([("PDF Documents", "*.pdf")], [".pdf"]),
}

TOOLS_DATA = [
    {"id": "WORD_TO_PDF", "title": "Word to PDF", "desc": "Convert Word documents to PDF quickly while preserving layout.", "badge": "DOCX", "cat": "Convert to PDF"},
    {"id": "EXCEL_TO_PDF", "title": "Excel to PDF", "desc": "Turn spreadsheets into PDF files with precise column fidelity.", "badge": "XLSX", "cat": "Convert to PDF"},
    {"id": "PPT_TO_PDF", "title": "PowerPoint to PDF", "desc": "Export your presentations to PDF with slide transitions preserved.", "badge": "PPTX", "cat": "Convert to PDF"},
    {"id": "PDF_TO_WORD", "title": "PDF to Word", "desc": "Extract editable Word documents from any PDF with high accuracy.", "badge": "PDF", "cat": "Convert from PDF"},
    {"id": "PDF_TO_EXCEL", "title": "PDF to Excel", "desc": "Convert tables from PDF files into fully editable Excel spreadsheets.", "badge": "PDF", "cat": "Convert from PDF"},
    {"id": "PDF_TO_PPT", "title": "PDF to PowerPoint", "desc": "Convert PDF slides back into editable PowerPoint presentations.", "badge": "PDF", "cat": "Convert from PDF"},
    {"id": "MERGE_PDF", "title": "Merge PDF", "desc": "Combine multiple PDF files into a single document in seconds.", "badge": "PDF", "cat": "PDF Tools"},
    {"id": "SPLIT_PDF", "title": "Split PDF", "desc": "Separate a PDF into individual pages or custom page ranges.", "badge": "PDF", "cat": "PDF Tools"},
    {"id": "COMPRESS_PDF", "title": "Compress PDF", "desc": "Reduce PDF file size without sacrificing readability.", "badge": "PDF", "cat": "PDF Tools"},
    {"id": "IMAGE_TO_PDF", "title": "JPG to PDF", "desc": "Package one or more JPG images into a clean PDF file.", "badge": "JPG", "cat": "Convert to PDF"},
    {"id": "PDF_TO_IMAGE", "title": "PDF to JPG", "desc": "Extract high-resolution images from each page of any PDF.", "badge": "PDF", "cat": "Convert from PDF"},
    {"id": "IMAGE_TO_WORD", "title": "JPG to Word", "desc": "Convert JPG images into editable Word documents using OCR.", "badge": "JPG", "cat": "Convert from PDF"},
    {"id": "PROTECT_PDF", "title": "Protect PDF", "desc": "Add password protection and permissions to keep PDFs secure.", "badge": "PDF", "cat": "PDF Tools"},
    {"id": "UNLOCK_PDF", "title": "Unlock PDF", "desc": "Remove password protection and restrictions from PDF files.", "badge": "PDF", "cat": "PDF Tools"},
]

class MiCOApp:
    def __init__(self, root):
        self.root = root
        self.root.title("MiCO File Converter Dashboard")
        self.root.configure(bg="#fcface")
        
        # Set ikon jendela aplikasi
        self._set_app_icon()

        self.root.geometry("1180x780")
        self.root.minsize(1020, 700)
        self._center_window()
        
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        
        self.output_dir = os.path.join(os.path.expanduser("~"), "Documents", "MiCO_Output")
        os.makedirs(self.output_dir, exist_ok=True)

        self.controller = AppController(self) if AppController else None

        self.active_category = "All"
        self.current_tool = None
        self.selected_files = []
        self.current_cols = 3
        self.current_filtered_tools = TOOLS_DATA
        
        self._resize_timer = None
        self._last_width = 1180
        self._last_state = self.root.state()
        self.is_converting = False

        self.container = ctk.CTkFrame(self.root, fg_color="#fcface")
        self.container.pack(fill="both", expand=True)

        self.dashboard_frame = ctk.CTkFrame(self.container, fg_color="transparent")
        self.tool_frame = ctk.CTkFrame(self.container, fg_color="transparent")
        self.success_frame = ctk.CTkFrame(self.container, fg_color="transparent")

        self._build_dashboard_ui()
        self._build_tool_ui()
        self.show_dashboard_view()

        self.root.bind("<Configure>", self._on_window_resize)

    def _set_app_icon(self):
        icon_ico = get_resource_path("app_icon.ico")
        icon_png = get_resource_path("app_icon.png")

        try:
            if sys.platform == "win32" and os.path.exists(icon_ico):
                self.root.iconbitmap(icon_ico)
            elif os.path.exists(icon_png):
                img = tk.PhotoImage(file=icon_png)
                self.root.iconphoto(True, img)
            elif os.path.exists(icon_ico):
                self.root.iconbitmap(icon_ico)
        except Exception as e:
            print(f"Gagal memuat ikon aplikasi: {e}")

    def _center_window(self):
        self.root.update_idletasks()
        width = 1180
        height = 780
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.root.geometry(f"{width}x{height}+{max(0, x)}+{max(0, y)}")

    def _on_close(self):
        if getattr(self, "is_converting", False):
            confirm = messagebox.askyesno(
                "Proses Berjalan",
                "Proses konversi sedang berlangsung. Yakin ingin keluar dan membatalkan proses?"
            )
            if not confirm:
                return

        self.root.quit()
        self.root.destroy()
        sys.exit(0)

    def hide_all_frames(self):
        self.dashboard_frame.pack_forget()
        self.tool_frame.pack_forget()
        self.success_frame.pack_forget()

    def _on_window_resize(self, event):
        if event.widget != self.root:
            return

        try:
            if self.root.state() == "iconic" or not self.root.winfo_viewable():
                return
        except Exception:
            pass

        current_state = self.root.state()
        state_changed = current_state != getattr(self, "_last_state", None)
        self._last_state = current_state

        current_w = event.width

        if state_changed:
            if getattr(self, "_resize_timer", None) is not None:
                self.root.after_cancel(self._resize_timer)
            self._last_width = current_w
            self._process_resize()
            return

        if abs(current_w - self._last_width) < 20:
            return
        
        self._last_width = current_w

        if getattr(self, "_resize_timer", None) is not None:
            self.root.after_cancel(self._resize_timer)

        self._resize_timer = self.root.after(100, self._process_resize)

    def _process_resize(self):
        width = self.root.winfo_width()
        is_zoomed_or_wide = (self.root.state() == "zoomed") or (width >= 1300)
        target_cols = 5 if is_zoomed_or_wide else 3

        if target_cols != getattr(self, "current_cols", 3):
            self.current_cols = target_cols
            self._regrid_cards()

    def _regrid_cards(self):
        cols = getattr(self, "current_cols", 3)
        for c in range(5):
            self.grid_scroll.grid_columnconfigure(c, weight=1 if c < cols else 0)

        for idx, card in enumerate(self.grid_scroll.winfo_children()):
            row = idx // cols
            col = idx % cols
            card.grid(row=row, column=col, padx=8, pady=8, sticky="ew")

    def _build_dashboard_ui(self):
        header = ctk.CTkFrame(self.dashboard_frame, fg_color="transparent")
        header.pack(fill="x", pady=(20, 5), padx=40)
        ctk.CTkLabel(header, text="All-in-one document conversion tools", font=ctk.CTkFont(size=26, weight="bold"), text_color="#0F172A").pack()
        ctk.CTkLabel(header, text="Convert, compress, and manage your files — fast, free, and beautifully simple.", text_color="#64748B", font=ctk.CTkFont(size=13)).pack(pady=2)

        search_frame = ctk.CTkFrame(self.dashboard_frame, fg_color="transparent")
        search_frame.pack(fill="x", padx=120, pady=10)
        self.search_entry = ctk.CTkEntry(search_frame, placeholder_text="🔍 Search conversion tools...", height=40, corner_radius=20, fg_color="#f5f5f2", font=ctk.CTkFont(size=13))
        self.search_entry.pack(fill="x")
        self.search_entry.bind("<KeyRelease>", self._filter_tools)

        tabs_frame = ctk.CTkFrame(self.dashboard_frame, fg_color="transparent")
        tabs_frame.pack(pady=5)

        self.tab_buttons = {}
        categories = ["All", "Convert to PDF", "Convert from PDF", "PDF Tools"]
        for cat in categories:
            btn = ctk.CTkButton(
                tabs_frame, text=cat,
                fg_color="#2563EB" if cat == self.active_category else "transparent",
                text_color="white" if cat == self.active_category else "#64748B",
                hover_color="#1D4ED8" if cat == self.active_category else "#E2E8F0",
                height=32, corner_radius=16,
                font=ctk.CTkFont(size=13, weight="bold" if cat == self.active_category else "normal"),
                command=lambda c=cat: self._select_category(c)
            )
            btn.pack(side="left", padx=4)
            self.tab_buttons[cat] = btn

        self.grid_scroll = ctk.CTkScrollableFrame(self.dashboard_frame, fg_color="transparent")
        self.grid_scroll.pack(fill="both", expand=True, padx=40, pady=10)

    def show_dashboard_view(self):
        self.hide_all_frames()
        self._filter_tools()
        self.dashboard_frame.pack(fill="both", expand=True)

    def _select_category(self, cat):
        self.active_category = cat
        for c, btn in self.tab_buttons.items():
            if c == cat:
                btn.configure(fg_color="#2563EB", text_color="white", hover_color="#1D4ED8", font=ctk.CTkFont(size=13, weight="bold"))
            else:
                btn.configure(fg_color="transparent", text_color="#64748B", hover_color="#E2E8F0", font=ctk.CTkFont(size=13, weight="normal"))
        self._filter_tools()

    def _filter_tools(self, event=None):
        query = self.search_entry.get().lower() if hasattr(self, 'search_entry') else ""
        filtered = []
        for t in TOOLS_DATA:
            match_cat = (self.active_category == "All") or (t["cat"] == self.active_category)
            match_query = query in t["title"].lower() or query in t["desc"].lower()
            if match_cat and match_query:
                filtered.append(t)
        self.current_filtered_tools = filtered
        self._render_cards(filtered)

    def _render_cards(self, tools):
        for child in self.grid_scroll.winfo_children():
            child.destroy()

        cols = getattr(self, "current_cols", 3)
        for c in range(5):
            self.grid_scroll.grid_columnconfigure(c, weight=1 if c < cols else 0)

        for idx, tool in enumerate(tools):
            row = idx // cols
            col = idx % cols

            card = tk.Frame(
                self.grid_scroll,
                bg="#f5f5f2",
                highlightthickness=1,
                highlightbackground="#E2E8F0",
                height=165,
                cursor="hand2"
            )
            card.grid(row=row, column=col, padx=8, pady=8, sticky="ew")
            card.pack_propagate(False)

            top_f = tk.Frame(card, bg="#f5f5f2")
            top_f.pack(fill="x", padx=14, pady=(12, 4))
            
            lbl_title = tk.Label(
                top_f, 
                text=tool["title"], 
                font=("Segoe UI", 11, "bold"), 
                bg="#f5f5f2",
                fg="#0F172A",
                anchor="w"
            )
            lbl_title.pack(side="left", fill="x", expand=True)

            is_pdf = (tool["badge"] == "PDF")
            badge_bg = "#FEE2E2" if is_pdf else "#E0F2FE"
            badge_fg = "#B91C1C" if is_pdf else "#0369A1"
            hover_border_color = "#EF4444" if is_pdf else "#2563EB"

            badge = tk.Label(
                top_f, 
                text=tool["badge"], 
                bg=badge_bg, 
                fg=badge_fg, 
                font=("Segoe UI", 8, "bold"), 
                padx=6, 
                pady=2
            )
            badge.pack(side="right")

            desc_label = tk.Label(
                card, 
                text=tool["desc"], 
                font=("Segoe UI", 9), 
                bg="#f5f5f2",
                fg="#64748B", 
                justify="left",
                anchor="nw",
                wraplength=200
            )
            desc_label.pack(fill="both", expand=True, padx=14, pady=(2, 8))

            def on_enter(e, c=card, h_color=hover_border_color):
                c.configure(highlightbackground=h_color, highlightthickness=2)

            def on_leave(e, c=card):
                c.configure(highlightbackground="#E2E8F0", highlightthickness=1)

            click_action = lambda e, t=tool: self.open_tool_page(t)

            for widget in (card, top_f, lbl_title, badge, desc_label):
                widget.bind("<Enter>", on_enter)
                widget.bind("<Leave>", on_leave)
                widget.bind("<Button-1>", click_action)

    def _build_tool_ui(self):
        top = ctk.CTkFrame(self.tool_frame, fg_color="transparent")
        top.pack(fill="x", padx=40, pady=15)
        ctk.CTkButton(top, text="‹ Back", fg_color="transparent", text_color="#64748B", hover_color="#E2E8F0", width=60, font=ctk.CTkFont(size=13, weight="bold"), command=self.show_dashboard_view).pack(side="left")
        self.nav_title_label = ctk.CTkLabel(top, text="  |   Tool", font=ctk.CTkFont(size=15, weight="bold"), text_color="#0F172A")
        self.nav_title_label.pack(side="left")

        self.tool_scroll = ctk.CTkScrollableFrame(self.tool_frame, fg_color="transparent")
        self.tool_scroll.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        self.center_wrapper = ctk.CTkFrame(self.tool_scroll, fg_color="transparent")
        self.center_wrapper.pack(fill="x", padx=80, pady=5)

        self.tool_main_title = ctk.CTkLabel(self.center_wrapper, text="Tool Title", font=ctk.CTkFont(size=26, weight="bold"), text_color="#0F172A")
        self.tool_main_title.pack(pady=(5, 2))
        
        self.tool_main_desc = ctk.CTkLabel(self.center_wrapper, text="Tool Description", text_color="#64748B", font=ctk.CTkFont(size=13))
        self.tool_main_desc.pack(pady=(0, 10))

        self.drop_box = ctk.CTkFrame(self.center_wrapper, fg_color="#F0F7FF", border_color="#93C5FD", border_width=2, corner_radius=12, height=170)
        self.drop_box.pack(fill="x", pady=5)
        self.drop_box.pack_propagate(False)

        ctk.CTkLabel(self.drop_box, text="☁️", font=ctk.CTkFont(size=28)).pack(pady=(12, 2))
        ctk.CTkLabel(self.drop_box, text="Drop files here or Browse", font=ctk.CTkFont(size=16, weight="bold"), text_color="#1E3A8A").pack()
        self.supported_ext_label = ctk.CTkLabel(self.drop_box, text="Select files to convert", font=ctk.CTkFont(size=13), text_color="#64748B")
        self.supported_ext_label.pack(pady=2)

        ctk.CTkButton(self.drop_box, text="Browse Files", fg_color="white", text_color="#2563EB", border_color="#2563EB", border_width=1, hover_color="#EFF6FF", width=140, height=34, font=ctk.CTkFont(size=13, weight="bold"), command=self._browse_files).pack(pady=6)

        try:
            self.drop_box._canvas.drop_target_register(DND_FILES)
            self.drop_box._canvas.dnd_bind("<<Drop>>", self._on_drag_drop_files)
        except Exception:
            pass

        self.file_list_frame = ctk.CTkScrollableFrame(self.center_wrapper, fg_color="transparent", height=340)

        self.btn_convert = ctk.CTkButton(
            self.center_wrapper,
            text="Select Files Above to Convert",
            fg_color="#CBD5E1",
            hover_color="#1D4ED8",
            text_color="white",
            text_color_disabled="#64748B",
            height=46,
            width=360,
            corner_radius=10,
            font=ctk.CTkFont(size=15, weight="bold"),
            command=self._execute_conversion
        )
        self.btn_convert.pack(pady=(20, 40))

    def open_tool_page(self, tool):
        self.current_tool = tool
        self.selected_files = []

        self.nav_title_label.configure(text=f"  |   {tool['title']}")
        self.tool_main_title.configure(text=tool['title'])
        self.tool_main_desc.configure(text=tool['desc'])

        tool_id = tool.get("id", "")
        if tool_id in TOOL_FILTER_MAP:
            ext_str = ", ".join([e.upper().replace(".", "") for e in TOOL_FILTER_MAP[tool_id][1]])
            self.supported_ext_label.configure(text=f"Supports {ext_str} files")
        else:
            self.supported_ext_label.configure(text="Select valid files")

        self.update_file_list_ui()
        self.hide_all_frames()
        self.tool_frame.pack(fill="both", expand=True)

    def _browse_files(self):
        tool_id = self.current_tool.get("id", "") if self.current_tool else ""
        filter_data = TOOL_FILTER_MAP.get(tool_id, ([("All Files", "*.*")], None))
        filetypes_arg = filter_data[0]

        files = filedialog.askopenfilenames(title="Select File", filetypes=filetypes_arg)
        if files:
            self._add_valid_files(files)

    def _on_drag_drop_files(self, event):
        try:
            files = self.root.tk.splitlist(event.data)
            if files:
                self._add_valid_files(files)
        except Exception:
            pass

    def _add_valid_files(self, files):
        tool_id = self.current_tool.get("id", "") if self.current_tool else ""
        valid_exts = TOOL_FILTER_MAP.get(tool_id, (None, None))[1]

        added_count = 0
        for f in files:
            ext = os.path.splitext(f)[1].lower()
            if valid_exts and ext not in valid_exts:
                continue
            if f not in self.selected_files:
                self.selected_files.append(f)
                added_count += 1

        if added_count == 0 and valid_exts:
            ext_formatted = ", ".join([e.upper() for e in valid_exts])
            messagebox.showwarning("File Tidak Sesuai", f"Fitur {self.current_tool['title']} hanya menerima file format: {ext_formatted}")

        self.update_file_list_ui()

    def update_file_list_ui(self):
        for child in self.file_list_frame.winfo_children():
            child.destroy()

        if not self.selected_files:
            self.file_list_frame.pack_forget()
            self.btn_convert.configure(state="disabled", text="Select Files Above to Convert", fg_color="#CBD5E1")
            return

        self.file_list_frame.pack(fill="x", pady=10)

        self.btn_convert.pack_forget()
        self.btn_convert.pack(pady=(20, 40))

        self.btn_convert.configure(state="normal", text=f"Convert Now ({len(self.selected_files)} File)", fg_color="#2563EB")

        top_info = ctk.CTkFrame(self.file_list_frame, fg_color="transparent")
        top_info.pack(fill="x", pady=(0, 5))
        ctk.CTkLabel(top_info, text=f"{len(self.selected_files)} file(s) selected", font=ctk.CTkFont(size=13, weight="bold"), text_color="#0F172A").pack(side="left")
        ctk.CTkButton(top_info, text="Clear all", fg_color="transparent", text_color="#EF4444", hover_color="#FEE2E2", width=60, font=ctk.CTkFont(size=12), command=self._clear_all_files).pack(side="right")

        for f in self.selected_files:
            f_card = ctk.CTkFrame(self.file_list_frame, fg_color="#f5f5f2", corner_radius=8, border_width=1, border_color="#E2E8F0", height=42)
            f_card.pack(fill="x", pady=3)
            f_card.pack_propagate(False)

            size_mb = round(os.path.getsize(f) / (1024 * 1024), 2) if os.path.exists(f) else 0
            ctk.CTkLabel(f_card, text=f"📄  {os.path.basename(f)}", font=ctk.CTkFont(size=13, weight="bold"), text_color="#0F172A").pack(side="left", padx=12)
            ctk.CTkLabel(f_card, text=f"{size_mb} MB", font=ctk.CTkFont(size=12), text_color="#64748B").pack(side="left", padx=10)
            ctk.CTkButton(f_card, text="✕", fg_color="transparent", text_color="#94A3B8", hover_color="#F1F5F9", width=30, command=lambda path=f: self._remove_single_file(path)).pack(side="right", padx=8)

    def _remove_single_file(self, path):
        if path in self.selected_files:
            self.selected_files.remove(path)
        self.update_file_list_ui()

    def _clear_all_files(self):
        self.selected_files = []
        self.update_file_list_ui()

    def _execute_conversion(self):
        if not self.selected_files:
            messagebox.showwarning("Warning", "Silakan pilih file terlebih dahulu!")
            return

        tool_id = self.current_tool.get("id", "") if self.current_tool else ""

        if tool_id == "PROTECT_PDF":
            dialog = ctk.CTkInputDialog(text="Masukkan kata sandi untuk melindungi PDF:", title="Protect PDF Password")
            password = dialog.get_input()
            if not password:
                return
            self._process_protect_pdf(password)
            return
        elif tool_id == "UNLOCK_PDF":
            dialog = ctk.CTkInputDialog(text="Masukkan kata sandi PDF untuk membuka kunci:", title="Unlock PDF Password")
            password = dialog.get_input()
            if not password:
                return
            self._process_unlock_pdf(password)
            return
        elif tool_id == "COMPRESS_PDF":
            self._process_compress_pdf()
            return

        if not self.controller:
            messagebox.showerror("Error", "AppController gagal dimuat!")
            return

        try:
            self.controller.select_tool(self.current_tool)
            self.controller.add_files(self.selected_files)
            self.controller.start_conversion()
        except Exception as e:
            messagebox.showerror("Error", f"Gagal memulai konversi: {e}")

    def _process_compress_pdf(self):
        self.set_converting_state(True)
        self.root.update()

        files_to_process = list(self.selected_files)

        import threading
        threading.Thread(target=self._worker_compress_pdf, args=(files_to_process,), daemon=True).start()

    def _worker_compress_pdf(self, files_to_process):
        start_time = time.time()
        try:
            import pymupdf as fitz
        except ImportError:
            try:
                import fitz
            except ImportError:
                self.root.after(0, lambda: self._handle_compress_error(
                    "Modul 'pymupdf' belum terinstal.\n\nJalankan perintah ini di terminal:\npip install pymupdf"
                ))
                return

        out_path = None
        try:
            out_files = []
            total_orig_size = 0
            total_new_size = 0

            for file_path in files_to_process:
                orig_size = os.path.getsize(file_path)
                total_orig_size += orig_size

                base_name = os.path.basename(file_path)
                name, ext = os.path.splitext(base_name)
                out_path = os.path.join(self.output_dir, f"{name}_compressed.pdf")

                doc = fitz.open(file_path)
                try:
                    for page in doc:
                        for img in page.get_images():
                            xref = img[0]
                            try:
                                old_stream = doc.xref_stream(xref)
                                old_len = len(old_stream) if old_stream else 0

                                pix = fitz.Pixmap(doc, xref)
                                if pix.n > 4:
                                    pix = fitz.Pixmap(fitz.csRGB, pix)
                                
                                img_bytes = pix.tobytes("jpeg", jpg_quality=60)
                                
                                if old_len == 0 or len(img_bytes) < old_len:
                                    page.replace_image(xref, stream=img_bytes)
                            except Exception:
                                pass

                    doc.save(
                        out_path,
                        garbage=3,
                        deflate=True,
                        deflate_images=True,
                        deflate_fonts=True
                    )
                finally:
                    doc.close()

                new_size = os.path.getsize(out_path)

                if new_size > orig_size:
                    shutil.copyfile(file_path, out_path)
                    new_size = orig_size

                total_new_size += new_size
                out_files.append(out_path)

            elapsed = round(time.time() - start_time, 1)

            orig_mb = round(total_orig_size / (1024 * 1024), 2)
            new_mb = round(total_new_size / (1024 * 1024), 2)
            saved_mb = round((total_orig_size - total_new_size) / (1024 * 1024), 2)

            size_display = f"{orig_mb} MB ➔ {new_mb} MB"
            saved_display = f"{saved_mb} MB" if saved_mb > 0 else "0.0 MB (Optimal)"

            summary = {
                "tool_name": "Compress PDF",
                "output_format": "PDF (Compressed)",
                "size_savings": size_display,
                "saved_amount": saved_display,
                "processing_time": f"{elapsed}s",
                "status": "Success",
                "out_dir": self.output_dir
            }

            self.root.after(0, lambda: self._handle_compress_success(summary))

        except Exception as e:
            if out_path and os.path.exists(out_path):
                try:
                    os.remove(out_path)
                except Exception:
                    pass

            err_msg = str(e)
            self.root.after(0, lambda msg=err_msg: self._handle_compress_error(f"Gagal mengompres PDF: {msg}"))

    def _handle_compress_success(self, summary):
        self.set_converting_state(False)
        self.show_success_view(summary)

    def _handle_compress_error(self, err_msg):
        self.set_converting_state(False)
        messagebox.showerror("Error", err_msg)

    def _process_protect_pdf(self, password):
        self.set_converting_state(True)
        self.root.update()

        start_time = time.time()
        try:
            try:
                from pypdf import PdfReader, PdfWriter
            except ImportError:
                from PyPDF2 import PdfReader, PdfWriter
        except ImportError:
            self.set_converting_state(False)
            messagebox.showerror("Error Library", "Modul 'pypdf' belum terinstal.\n\nJalankan perintah ini di terminal:\npip install pypdf")
            return

        try:
            out_files = []
            for file_path in self.selected_files:
                reader = PdfReader(file_path)
                writer = PdfWriter()

                for page in reader.pages:
                    writer.add_page(page)

                writer.encrypt(password)

                base_name = os.path.basename(file_path)
                name, ext = os.path.splitext(base_name)
                out_path = os.path.join(self.output_dir, f"{name}_protected.pdf")

                with open(out_path, "wb") as f:
                    writer.write(f)

                out_files.append(out_path)

            elapsed = round(time.time() - start_time, 1)

            summary = {
                "tool_name": "Protect PDF",
                "output_format": "PDF (Encrypted)",
                "size_savings": "Protected",
                "saved_amount": "Protected",
                "processing_time": f"{elapsed}s",
                "status": "Success",
                "out_dir": self.output_dir
            }
            self.set_converting_state(False)
            self.show_success_view(summary)
        except Exception as e:
            self.set_converting_state(False)
            messagebox.showerror("Error", f"Gagal memproteksi PDF: {e}")

    def _process_unlock_pdf(self, password):
        self.set_converting_state(True)
        self.root.update()

        start_time = time.time()
        try:
            try:
                from pypdf import PdfReader, PdfWriter
            except ImportError:
                from PyPDF2 import PdfReader, PdfWriter
        except ImportError:
            self.set_converting_state(False)
            messagebox.showerror("Error Library", "Modul 'pypdf' belum terinstal.\n\nJalankan perintah ini di terminal:\npip install pypdf")
            return

        try:
            out_files = []
            for file_path in self.selected_files:
                reader = PdfReader(file_path)
                
                if reader.is_encrypted:
                    decrypt_result = reader.decrypt(password)
                    if decrypt_result == 0:
                        raise ValueError("Kata sandi yang dimasukkan salah!")

                writer = PdfWriter()
                for page in reader.pages:
                    writer.add_page(page)

                base_name = os.path.basename(file_path)
                name, ext = os.path.splitext(base_name)
                out_path = os.path.join(self.output_dir, f"{name}_unlocked.pdf")

                with open(out_path, "wb") as f:
                    writer.write(f)

                out_files.append(out_path)

            elapsed = round(time.time() - start_time, 1)

            summary = {
                "tool_name": "Unlock PDF",
                "output_format": "PDF (Unlocked)",
                "size_savings": "Unlocked",
                "saved_amount": "Unlocked",
                "processing_time": f"{elapsed}s",
                "status": "Success",
                "out_dir": self.output_dir
            }
            self.set_converting_state(False)
            self.show_success_view(summary)
        except Exception as e:
            self.set_converting_state(False)
            messagebox.showerror("Error", f"Gagal membuka kunci PDF: {e}")

    def set_converting_state(self, is_processing):
        self.is_converting = is_processing
        if is_processing:
            self.btn_convert.configure(state="disabled", text="Converting...", fg_color="#93C5FD")
        else:
            self.btn_convert.configure(state="normal", text=f"Convert Now ({len(self.selected_files)} File)", fg_color="#2563EB")

    def show_success_view(self, summary):
        self.hide_all_frames()

        for child in self.success_frame.winfo_children():
            child.destroy()

        self.success_frame.pack(fill="both", expand=True)

        ctk.CTkLabel(self.success_frame, text="✓", font=ctk.CTkFont(size=32, weight="bold"), text_color="#10B981", fg_color="#D1FAE5", width=64, height=64, corner_radius=32).pack(pady=(25, 10))
        ctk.CTkLabel(self.success_frame, text="Conversion Complete!", font=ctk.CTkFont(size=24, weight="bold"), text_color="#0F172A").pack()
        ctk.CTkLabel(self.success_frame, text="Your file is ready to download.", text_color="#64748B", font=ctk.CTkFont(size=13)).pack(pady=(2, 12))

        center_container = ctk.CTkFrame(self.success_frame, fg_color="transparent", width=540)
        center_container.pack(pady=5)

        card = ctk.CTkFrame(center_container, fg_color="#f5f5f2", corner_radius=12, border_width=1, border_color="#E2E8F0", width=540)
        card.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(card, text="CONVERSION SUMMARY", font=ctk.CTkFont(size=12, weight="bold"), text_color="#94A3B8").pack(anchor="w", padx=24, pady=(16, 10))

        items = [
            ("Tool Used", summary.get("tool_name", "N/A")),
            ("Output Format", summary.get("output_format", "N/A")),
            ("File Size Savings", summary.get("size_savings", "0 KB")),
            ("Processing Time", summary.get("processing_time", "0s")),
            ("Status", summary.get("status", "Success"))
        ]

        for label, val in items:
            row = ctk.CTkFrame(card, fg_color="transparent")
            row.pack(fill="x", padx=24, pady=4)
            ctk.CTkLabel(row, text=label, text_color="#64748B", font=ctk.CTkFont(size=13)).pack(side="left")
            val_color = "#10B981" if label == "Status" else "#0F172A"
            ctk.CTkLabel(row, text=val, text_color=val_color, font=ctk.CTkFont(size=13, weight="bold")).pack(side="right")

        ctk.CTkFrame(card, fg_color="transparent", height=12).pack()

        badge = ctk.CTkFrame(center_container, fg_color="#ECFDF5", border_color="#A7F3D0", border_width=1, corner_radius=8, width=540)
        badge.pack(fill="x", pady=(0, 15))
        
        saved_info = summary.get("saved_amount", summary.get("size_savings", "Protected"))
        ctk.CTkLabel(badge, text=f"★ Saved {saved_info} — process complete!", text_color="#047857", font=ctk.CTkFont(size=13, weight="bold")).pack(pady=8)

        btn_box = ctk.CTkFrame(self.success_frame, fg_color="transparent")
        btn_box.pack(pady=5)

        out_folder = summary.get("out_dir", self.output_dir)
        ctk.CTkButton(
            btn_box, 
            text="Download File", 
            fg_color="#2563EB", 
            hover_color="#1D4ED8", 
            text_color="#FFFFFF", 
            height=42, 
            width=170, 
            font=ctk.CTkFont(size=13, weight="bold"), 
            command=lambda: self.download_and_save_file(out_folder)
        ).pack(side="left", padx=8)

        ctk.CTkButton(
            btn_box, 
            text="Convert Another", 
            fg_color="#f5f5f2", 
            text_color="#0F172A", 
            border_color="#CBD5E1", 
            border_width=1, 
            hover_color="#E2E8F0", 
            height=42, 
            width=170, 
            font=ctk.CTkFont(size=13, weight="bold"), 
            command=self.show_dashboard_view
        ).pack(side="left", padx=8)

        ctk.CTkButton(
            self.success_frame, 
            text="← Back to all tools", 
            fg_color="transparent", 
            text_color="#64748B", 
            hover_color="#E2E8F0", 
            font=ctk.CTkFont(size=13), 
            command=self.show_dashboard_view
        ).pack(pady=(10, 10))

    def download_and_save_file(self, source_dir):
        target_folder = filedialog.askdirectory(title="Pilih Folder Penyimpanan File")
        if target_folder:
            source_abs = os.path.abspath(source_dir)
            target_abs = os.path.abspath(target_folder)

            if source_abs == target_abs:
                messagebox.showinfo("Informasi", f"File sudah berada di folder ini:\n{target_folder}")
                open_folder(target_folder)
                return

            try:
                files = [f for f in os.listdir(source_dir) if os.path.isfile(os.path.join(source_dir, f))]
                if not files:
                    messagebox.showwarning("Peringatan", "Tidak ada file yang dapat diunduh.")
                    return

                for file_name in files:
                    src_path = os.path.join(source_dir, file_name)
                    dest_path = os.path.join(target_folder, file_name)

                    if os.path.abspath(src_path) == os.path.abspath(dest_path):
                        continue

                    shutil.copy2(src_path, dest_path)

                messagebox.showinfo("Berhasil", f"File berhasil disimpan ke:\n{target_folder}")
                open_folder(target_folder)
            except Exception as e:
                messagebox.showerror("Error", f"Gagal menyimpan file: {e}\n\nPastikan file tidak sedang dibuka di aplikasi lain.")

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = MiCOApp(root)
    root.mainloop()