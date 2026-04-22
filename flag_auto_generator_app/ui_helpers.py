import tkinter as tk
from tkinter import filedialog, ttk


def pick_save_path(title: str, defaultextension: str, filetypes, parent=None):
    if parent is None:
        root = tk.Tk()
        root.withdraw()
        path = filedialog.asksaveasfilename(
            title=title,
            defaultextension=defaultextension,
            filetypes=filetypes,
        )
        root.destroy()
        return path
    return filedialog.asksaveasfilename(
        parent=parent,
        title=title,
        defaultextension=defaultextension,
        filetypes=filetypes,
    )


class LoadingDialog(tk.Toplevel):
    """ローディング表示用のダイアログ。"""

    def __init__(self, parent, title="処理中...", message="お待ちください..."):
        super().__init__(parent)
        self.title(title)
        self.transient(parent)
        self.grab_set()

        window_width = 300
        window_height = 120
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        center_x = int(screen_width / 2 - window_width / 2)
        center_y = int(screen_height / 2 - window_height / 2)
        self.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")

        self.protocol("WM_DELETE_WINDOW", lambda: None)

        frame = ttk.Frame(self, padding=20)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text=message, font=("", 10)).pack(pady=(10, 15))

        self.progress = ttk.Progressbar(
            frame,
            mode="indeterminate",
            length=250,
        )
        self.progress.pack(pady=10)
        self.progress.start(10)

        self.update()

    def close(self):
        self.progress.stop()
        self.grab_release()
        self.destroy()
