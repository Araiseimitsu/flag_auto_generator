# -*- coding: utf-8 -*-
"""
落ち着いたライトテーマ用の共通スタイル。
彩度を抑えつつ、情報階層が伝わる見た目をまとめる。
"""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk

# sandstone: 白飛びしにくい落ち着いたトーンのライトテーマ
THEME_NAME = "sandstone"

# 全体のトーンを抑えるための固定色
APP_BG = "#efece6"
PANEL_BG = "#f6f4ef"
PANEL_BORDER = "#d3cec4"
PANEL_HEADER = "#5a6066"
TEXT_SOFT = "#6b6f73"
TEXT_STRONG = "#283038"
SCROLL_STRIP_BG = "#e4e0d8"
TREE_HEADER_BG = "#ece8e0"
STEP_BADGE_BG = "#dde2e4"

# 本文は Windows 標準 UI 系で可読性を確保
FONT_UI = "Segoe UI"

# 補足説明の折り返し幅（px）
HINT_WRAPLENGTH = 700


def font_ui(size: int = 10, weight: str = "normal") -> tuple[str, int] | tuple[str, int, str]:
    if weight in ("", "normal"):
        return (FONT_UI, size)
    return (FONT_UI, size, weight)


def apply_app_style(root) -> None:
    """画面全体のフレーム・カード・ラベル・表を静かな業務UIに揃える。"""
    style: ttk.Style = ttk.Style(root)

    root.configure(bg=APP_BG)

    style.configure(".", font=font_ui(10))
    style.configure("TFrame", background=APP_BG)
    style.configure("Surface.TFrame", background=PANEL_BG)
    style.configure("Toolbar.TFrame", background=APP_BG)
    style.configure("Inline.TFrame", background=PANEL_BG)

    style.configure(
        "HeroEyebrow.TLabel",
        background=APP_BG,
        foreground=TEXT_SOFT,
        font=font_ui(9, "bold"),
    )
    style.configure(
        "HeroTitle.TLabel",
        background=APP_BG,
        foreground=TEXT_STRONG,
        font=font_ui(18, "bold"),
    )
    style.configure(
        "HeroBody.TLabel",
        background=APP_BG,
        foreground=TEXT_SOFT,
        font=font_ui(10),
    )
    style.configure(
        "Chip.TLabel",
        background=STEP_BADGE_BG,
        foreground=TEXT_STRONG,
        font=font_ui(8, "bold"),
        padding=(10, 4),
    )
    style.configure(
        "StepBadge.TLabel",
        background=STEP_BADGE_BG,
        foreground=TEXT_STRONG,
        font=font_ui(8, "bold"),
        padding=(8, 3),
    )
    style.configure(
        "StepTitle.TLabel",
        background=APP_BG,
        foreground=TEXT_STRONG,
        font=font_ui(12, "bold"),
    )
    style.configure(
        "StepBody.TLabel",
        background=APP_BG,
        foreground=TEXT_SOFT,
        font=font_ui(9),
    )
    style.configure(
        "CardTitle.TLabel",
        background=PANEL_BG,
        foreground=TEXT_STRONG,
        font=font_ui(9, "bold"),
    )
    style.configure(
        "CardNote.TLabel",
        background=PANEL_BG,
        foreground=TEXT_SOFT,
        font=font_ui(9),
    )

    style.configure(
        "AppCard.TLabelframe",
        background=PANEL_BG,
        borderwidth=1,
        relief="solid",
        bordercolor=PANEL_BORDER,
        lightcolor=PANEL_BORDER,
        darkcolor=PANEL_BORDER,
    )
    style.configure(
        "AppCard.TLabelframe.Label",
        background=PANEL_BG,
        foreground=PANEL_HEADER,
        font=font_ui(9, "bold"),
    )

    style.configure(
        "Data.Treeview",
        rowheight=28,
        borderwidth=0,
        fieldbackground=PANEL_BG,
        background=PANEL_BG,
        foreground=TEXT_STRONG,
    )
    style.configure(
        "Data.Treeview.Heading",
        background=TREE_HEADER_BG,
        foreground=TEXT_STRONG,
        relief="flat",
        font=font_ui(9, "bold"),
        padding=(8, 7),
    )
    style.map(
        "Data.Treeview",
        background=[("selected", "#cfd6d9")],
        foreground=[("selected", TEXT_STRONG)],
    )


def apply_preview_treeview_style(root) -> None:
    """プレビュー用 Treeview。落ち着いたニュートラル配色で一覧性を高める。"""
    style: ttk.Style = ttk.Style(root)
    name = "Preview.Treeview"
    style.configure(
        name,
        rowheight=28,
        borderwidth=0,
        fieldbackground=PANEL_BG,
        background=PANEL_BG,
        foreground=TEXT_STRONG,
    )
    style.configure(
        f"{name}.Heading",
        background=TREE_HEADER_BG,
        foreground=TEXT_STRONG,
        relief="flat",
        font=font_ui(9, "bold"),
        padding=(8, 7),
    )
    style.map(
        name,
        background=[("selected", "#cfd6d9")],
        foreground=[("selected", TEXT_STRONG)],
    )


def make_step_caption(
    parent: tk.Widget,
    step_num: int,
    title: str,
    blurb: str | None = None,
) -> None:
    """ステップ番号＋見出し＋補足。カード手前に置く。"""
    frame = ttk.Frame(parent, style="Toolbar.TFrame")
    frame.pack(anchor="w", fill="x", pady=(0, 8))

    head = ttk.Frame(frame, style="Toolbar.TFrame")
    head.pack(anchor="w")
    ttk.Label(head, text=f"Step {step_num}", style="StepBadge.TLabel").pack(side="left")
    ttk.Label(head, text=title, style="StepTitle.TLabel").pack(side="left", padx=(10, 0))

    if blurb:
        ttk.Label(
            frame,
            text=blurb,
            style="StepBody.TLabel",
            wraplength=HINT_WRAPLENGTH,
            justify="left",
        ).pack(anchor="w", pady=(6, 0))


def section_separator(parent: tk.Widget) -> ttk.Separator:
    sep = ttk.Separator(parent, orient=tk.HORIZONTAL)
    sep.pack(fill=tk.X, pady=18, padx=0)
    return sep


def place_toplevel_center(window: tk.Toplevel, width: int, height: int) -> None:
    """
    モーダル等を画面中央に。ジオメトリを設定してから必ず呼ぶ（pack 後推奨で update_idletasks を含む）。
    """
    window.update_idletasks()
    sw = window.winfo_screenwidth()
    sh = window.winfo_screenheight()
    x = max(0, (sw - width) // 2)
    y = max(0, (sh - height) // 2)
    window.geometry(f"{width}x{height}+{x}+{y}")


def scrollstrip_background() -> str:
    """主窓スクロール帯。テーマ白より控えめな色で目の負担を下げる。"""
    return SCROLL_STRIP_BG
