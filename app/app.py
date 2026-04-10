import customtkinter as ctk
from pathlib import Path

# Apply Design System theme before any widget is created.
_THEME_PATH = Path(__file__).parent.parent / "themes" / "ctk_theme.json"
ctk.set_default_color_theme(str(_THEME_PATH))
ctk.set_appearance_mode("light")

from app.screens.screen_1a import Screen1A


class SpreadApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Spread Automation")
        self.geometry("920x700")
        self.minsize(760, 560)

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.screen_1a = Screen1A(self)
        self.screen_1a.grid(row=0, column=0, sticky="nsew")


if __name__ == "__main__":
    app = SpreadApp()
    app.mainloop()
