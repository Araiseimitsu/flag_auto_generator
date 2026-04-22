# -*- coding: utf-8 -*-
"""使い方モーダル。主画面と同じニュートラルトーンで表示する。"""
import tkinter as tk
from tkinter import ttk

import ttkbootstrap as tb
from ttkbootstrap.constants import SECONDARY

from .ui_theme import APP_BG, TEXT_SOFT, TEXT_STRONG, font_ui, place_toplevel_center

# 大きく・中央。最小サイズは小さいディスプレイ用
_HELP_W = 1000
_HELP_H = 720
_HELP_MIN_W = 820
_HELP_MIN_H = 560

# Text 用（テーマ白より一段暗め。眩しさを抑える）
_TEXT_BG = "#ede9e2"
_TEXT_FG = TEXT_STRONG
_TEXT_SEL_BG = "#cfd6d9"


def open_help_window(parent) -> None:
    win = tb.Toplevel(parent)
    win.title("使い方")
    win.transient(parent)
    win.grab_set()
    win.minsize(_HELP_MIN_W, _HELP_MIN_H)
    win.configure(bg=APP_BG)

    head = ttk.Frame(win, padding=(30, 22, 30, 14))
    head.pack(fill=tk.X)
    tk.Label(
        head,
        text="検査シート設定ツール",
        font=font_ui(14, "bold"),
        foreground=TEXT_STRONG,
        background=APP_BG,
    ).pack(anchor=tk.W)
    tk.Label(
        head,
        text="使い方",
        font=font_ui(9, "bold"),
        foreground=TEXT_SOFT,
        background=APP_BG,
    ).pack(anchor=tk.W, pady=(4, 0))
    tk.Label(
        head,
        text="画面の上から順に入力し、最後に「この内容で Excel を保存・生成」で保存先を選びます。",
        font=font_ui(10),
        foreground=TEXT_SOFT,
        background=APP_BG,
        wraplength=900,
        justify=tk.LEFT,
    ).pack(anchor=tk.W, pady=(8, 0))

    mid = ttk.Frame(win, padding=(24, 0, 24, 8))
    mid.pack(fill=tk.BOTH, expand=True)

    scroll = ttk.Scrollbar(mid, orient=tk.VERTICAL)
    body = tk.Text(
        mid,
        wrap=tk.WORD,
        yscrollcommand=scroll.set,
        font=font_ui(10),
        height=1,
        width=1,
        padx=20,
        pady=18,
        highlightthickness=0,
        relief=tk.FLAT,
        background=_TEXT_BG,
        foreground=_TEXT_FG,
        selectbackground=_TEXT_SEL_BG,
        cursor="arrow",
    )
    scroll.config(command=body.yview)
    body.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scroll.pack(side=tk.RIGHT, fill=tk.Y)

    body.tag_configure(
        "sec",
        font=font_ui(11, "bold"),
        foreground=TEXT_STRONG,
        lmargin1=0,
        lmargin2=0,
        spacing1=6,
    )
    body.tag_configure("sub", font=font_ui(10, "bold"), foreground="#52585d", spacing1=10)
    body.tag_configure("tx", font=font_ui(10), foreground=_TEXT_FG, lmargin1=6, lmargin2=12)
    body.tag_configure("note", font=font_ui(9), foreground=TEXT_SOFT, lmargin1=6, lmargin2=12)
    body.tag_configure(
        "mono",
        font=("Consolas", 9) if _font_available("Consolas") else font_ui(9),
        foreground="#333333",
    )
    body.tag_configure("foot", font=font_ui(9), foreground=TEXT_SOFT)

    _append_help_content(body)
    body.configure(state=tk.DISABLED)

    foot = ttk.Frame(win, padding=(24, 8, 24, 20))
    foot.pack(fill=tk.X)
    tk.Label(
        foot,
        text="Version 0.7.0  ·  2026-02-26  ·  DIP Dpertment / A・T",
        font=font_ui(9),
        foreground=TEXT_SOFT,
        background=APP_BG,
    ).pack(side=tk.LEFT)
    tb.Button(foot, text="閉じる", command=win.destroy, bootstyle=SECONDARY).pack(side=tk.RIGHT)

    win.update_idletasks()
    place_toplevel_center(win, _HELP_W, _HELP_H)
    win.focus_set()


def _font_available(name: str) -> bool:
    try:
        import tkinter.font as tkfont

        font = tkfont.Font(family=name, size=9)
        return font.actual()["family"].lower() == name.lower()
    except Exception:
        return False


def _append_help_content(body: tk.Text) -> None:
    def sec(text: str) -> None:
        body.insert(tk.END, text + "\n", "sec")

    def sub(text: str) -> None:
        body.insert(tk.END, text + "\n", "sub")

    def tx(text: str) -> None:
        body.insert(tk.END, text + "\n", "tx")

    def note(text: str) -> None:
        body.insert(tk.END, text + "\n", "note")

    sec("このツールの目的")
    tx(
        "元の工程内検査シート（.xlsx）に、L〜SR 列向けの式（依頼の集約判定、自動測定の参照、測定不要で「-」）を一括で書き込むための補助です。"
    )
    body.insert(tk.END, "\n", "tx")

    sec("作業の流れ（上から順）")
    sub("1. ファイル")
    tx("「参照…」で雛形を指定し、プレビュー表で中身の見え方を確認します。シート名を変えたら「表を再表示」。")
    sub("2. 基本・測定不要（任意）")
    tx("普段使う「シート名」に合わせます。E 列に「測定不要」と入る行番号を入れると、測定行の終わり・工具の開始行などは内側で計算します。")
    tx("L〜SR に「-」を入れたい測定 No は、1 件ずつ「追加」で一覧に登録します。")
    sub("3. 工具・自動測定")
    tx("工具名と、その工具が扱う測定 No（カンマ区切り）を表に登録します。自動測定ブロック向けに「測定 No → データの順番」の対応も入れます。")
    sub("4. 保存")
    tx("「この内容で Excel を保存・生成」で、開いていない保存先名を指定して出力します。元のファイルは上書きしません。")
    body.insert(tk.END, "\n", "tx")

    sec("先頭 1〜3 行目の集計式（L 列の例）")
    tx("シート上段の集計式は、次のような想定に合わせて L〜SN 列の 1〜3 行目を補正することがあります。")
    body.insert(tk.END, "=SUMPRODUCT(--(L11:L119<>\"\"),--(MOD(ROW(L11:L119)-ROW(L11),3)=0))", "mono")
    body.insert(tk.END, "\n", "tx")
    note("2 行目は L12:L120、3 行目は L13:L121 のように、参照する行が 1 行ずつ下がります。")
    note(
        "例の 119 行目（L119 や L11:L119 の範囲の終端に相当）は、測定項目の数（データ行の取り扱い範囲）に応じて変えてください。式内で同じ範囲を参照する箇所は、すべて同じ終端行に揃えます。"
    )
    body.insert(tk.END, "\n", "tx")

    sec("知っておきたいこと")
    note("・ 元ファイルと、保存先にする .xlsx は、保存して閉じた状態にしてください。")
    note("・ 測定 No は原則として半角の整数で指定します。")
    note("・ 行番号の前提（測定行の幅など）が、実物のブックと違うと、式の対象行がずれます。")
    note("・ 生成直後、Excel から修復の確認が出る場合は「はい」で開いてかまいません。書式が大きく崩れる場合は、必要な式だけ既存のブックへ貼り付けて使う方法もあります。")
    body.insert(tk.END, "\n", "foot")
    body.insert(
        tk.END,
        "本ヘルプの内容はアプリのバージョンに合わせて更新します。不具合や分かりづらい箇所は、担当までお知らせください。\n",
        "foot",
    )
