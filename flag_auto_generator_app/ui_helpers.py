# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import filedialog, ttk

import ttkbootstrap as tb
from ttkbootstrap.constants import SECONDARY

from .ui_theme import PANEL_BG, TEXT_SOFT, font_ui, place_toplevel_center


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


class LoadingDialog(tb.Toplevel):
    """ローディング表示。静かな配色と広めの余白で視認性を保つ。"""

    def __init__(self, parent, title="処理中...", message="お待ちください..."):
        super().__init__(parent)
        self.title(title)
        self.transient(parent)
        self.grab_set()
        self.configure(bg=PANEL_BG)

        window_width = 520
        window_height = 168

        self.protocol("WM_DELETE_WINDOW", lambda: None)

        frame = ttk.Frame(self, padding=28)
        frame.pack(fill=tk.BOTH, expand=True)

        tb.Label(
            frame,
            text=message,
            font=font_ui(10),
            bootstyle=SECONDARY,
            foreground=TEXT_SOFT,
            wraplength=460,
            justify=tk.LEFT,
        ).pack(anchor=tk.W, pady=(0, 14))

        self.progress = ttk.Progressbar(
            frame,
            mode="indeterminate",
            length=440,
        )
        self.progress.pack(anchor=tk.W)
        self.progress.start(10)

        self.update()
        self.update_idletasks()
        place_toplevel_center(self, window_width, window_height)

    def close(self):
        self.progress.stop()
        self.grab_release()
        self.destroy()
