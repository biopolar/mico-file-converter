import customtkinter as ctk

class AnimatedAccordion(ctk.CTkFrame):
    def __init__(self, master, title, items, select_callback, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.select_callback = select_callback
        self.raw_title = title
        self.is_open = False
        self.animating = False

        self.btn_header = ctk.CTkButton(
            self, text=f"{self.raw_title}  ▾", anchor="w",
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="transparent", text_color=("gray15", "gray85"),
            hover_color=("gray75", "gray25"), height=34, command=self.toggle
        )
        self.btn_header.pack(fill="x")

        self.content_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.content_frame.pack_propagate(False)
        self.content_frame.configure(height=0)
        self.content_frame.pack(fill="x")

        self.sub_buttons = {}
        for label, mode_key in items:
            btn = ctk.CTkButton(
                self.content_frame, text=f"  •  {label}", anchor="w", height=30,
                font=ctk.CTkFont(size=12), fg_color="transparent",
                text_color=("gray30", "gray70"), hover_color=("gray70", "gray30"),
                command=lambda m=mode_key, l=label: self.on_item_click(m, l)
            )
            btn.pack(fill="x", padx=(10, 0), pady=2)
            self.sub_buttons[mode_key] = btn

        self.target_height = len(items) * 34
        self.current_height = 0

    def toggle(self):
        if self.animating: return
        self.is_open = not self.is_open
        arrow = "▴" if self.is_open else "▾"
        self.btn_header.configure(text=f"{self.raw_title}  {arrow}")
        self.animate()

    def animate(self):
        self.animating = True
        step = 6
        if self.is_open:
            if self.current_height < self.target_height:
                self.current_height = min(self.current_height + step, self.target_height)
                self.content_frame.configure(height=self.current_height)
                self.after(10, self.animate)
            else: self.animating = False
        else:
            if self.current_height > 0:
                self.current_height = max(self.current_height - step, 0)
                self.content_frame.configure(height=self.current_height)
                self.after(10, self.animate)
            else: self.animating = False

    def on_item_click(self, mode_key, label):
        self.select_callback(mode_key, label, self)

    def highlight_button(self, mode_key):
        for key, btn in self.sub_buttons.items():
            if key == mode_key:
                btn.configure(fg_color=["#3B8ED0", "#1F6AA5"], text_color="white")
            else:
                btn.configure(fg_color="transparent", text_color=("gray30", "gray70"))

    def reset_highlights(self):
        for btn in self.sub_buttons.values():
            btn.configure(fg_color="transparent", text_color=("gray30", "gray70"))