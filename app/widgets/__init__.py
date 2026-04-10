import customtkinter as ctk
from tkinter import filedialog

class FilePickerWidget(ctk.CTkFrame):
    def __init__(self, master, label_text, **kwargs):
        super().__init__(master, **kwargs)
        self.path_var = ctk.StringVar()
        
        self.label = ctk.CTkLabel(self, text=label_text, width=150, anchor="w")
        self.label.pack(side="left", padx=5)
        
        self.entry = ctk.CTkEntry(self, textvariable=self.path_var, state="disabled")
        self.entry.pack(side="left", padx=5, fill="x", expand=True)
        
        self.btn = ctk.CTkButton(self, text="Browse...", command=self._browse, width=80)
        self.btn.pack(side="left", padx=5)

    def _browse(self):
        filename = filedialog.askopenfilename()
        if filename:
            self.path_var.set(filename)
            # Re-enable briefly to update content if it was disabled in a weird state
            self.entry.configure(state="normal")
            self.entry.delete(0, "end")
            self.entry.insert(0, filename)
            self.entry.configure(state="disabled")

    def get_path(self):
        return self.path_var.get()

class ProgressWidget(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        
        self.label_var = ctk.StringVar(value="Ready")
        self.label = ctk.CTkLabel(self, textvariable=self.label_var)
        self.label.pack(anchor="w", padx=5, pady=(2, 0))
        
        self.progressbar = ctk.CTkProgressBar(self)
        self.progressbar.pack(fill="x", padx=5, pady=(2, 5))
        self.progressbar.set(0)

    def update_progress(self, message: str, progress: float):
        self.label_var.set(message)
        self.progressbar.set(progress)
        self.update_idletasks()
