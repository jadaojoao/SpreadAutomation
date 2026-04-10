import customtkinter as ctk
from app.screens.screen_1a import Screen1A

class SpreadApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("Spread Automation - Mode 1A")
        self.geometry("850x650")
        
        # Setup UI Layout
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        # We only have one screen (Mode 1A) for Phase 1.
        self.screen_1a = Screen1A(self)
        self.screen_1a.grid(row=0, column=0, sticky="nsew")

if __name__ == "__main__":
    app = SpreadApp()
    app.mainloop()
