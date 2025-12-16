"""
PyInstallerでexe化するためのビルドスクリプト
"""
import os
import subprocess
import sys
from pathlib import Path

try:
    from PIL import Image, ImageDraw, ImageFont
except ImportError:
    print("Pillowがインストールされていません。インストール中...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "Pillow"])
    from PIL import Image, ImageDraw, ImageFont


def create_icon(output_path: str = "app_icon.ico"):
    """簡単なアプリアイコンを作成"""
    # 256x256の画像を作成
    size = 256
    img = Image.new("RGBA", (size, size), (70, 130, 180, 255))  # スチールブルー背景
    
    draw = ImageDraw.Draw(img)
    
    # 円を描画
    margin = 20
    draw.ellipse(
        [margin, margin, size - margin, size - margin],
        fill=(255, 255, 255, 255),
        outline=(50, 100, 150, 255),
        width=5
    )
    
    # 中央に「E」の文字を描画（簡易的なアイコン）
    try:
        # システムフォントを使用
        font_size = 150
        font = ImageFont.truetype("arial.ttf", font_size)
    except:
        # フォントが見つからない場合はデフォルトフォント
        font = ImageFont.load_default()
        font_size = 100
    
    text = "E"
    bbox = draw.textbbox((0, 0), text, font=font)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]
    
    position = ((size - text_width) // 2, (size - text_height) // 2 - 20)
    draw.text(position, text, fill=(70, 130, 180, 255), font=font)
    
    # 複数のサイズを含むICOファイルとして保存
    sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
    images = []
    for s in sizes:
        resized = img.resize(s, Image.Resampling.LANCZOS)
        images.append(resized)
    
    # ICO形式で保存（最初の画像に複数サイズの情報を含める）
    # Pillowは自動的に複数サイズを処理します
    img.save(output_path, format="ICO", sizes=[s for s in sizes])
    print(f"アイコンファイルを作成しました: {output_path}")


def build_exe():
    """PyInstallerでexe化"""
    script_name = "auto_mesure_order_request.py"
    icon_path = "app_icon.ico"
    
    # アイコンファイルが存在しない場合は作成
    if not os.path.exists(icon_path):
        print("アイコンファイルが見つかりません。作成中...")
        create_icon(icon_path)
    
    # PyInstallerがインストールされているか確認
    try:
        import PyInstaller
    except ImportError:
        print("PyInstallerがインストールされていません。インストール中...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
    
    # PyInstallerコマンドを構築（Windowsでも確実に動作するようにpython -m PyInstallerを使用）
    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--onefile",  # 単一ファイルにまとめる
        "--windowed",  # コンソールウィンドウを表示しない（GUIアプリ用）
        f"--icon={icon_path}",  # アイコンを指定
        "--name=AutoMesureOrderRequest",  # exeファイル名
        "--clean",  # ビルド前に一時ファイルをクリーンアップ
        script_name
    ]
    
    print("PyInstallerでexe化を開始します...")
    print(f"実行コマンド: {' '.join(cmd)}")
    
    try:
        subprocess.check_call(cmd)
        print("\n✓ ビルドが完了しました！")
        print(f"exeファイルは dist/AutoMesureOrderRequest.exe に生成されました。")
    except subprocess.CalledProcessError as e:
        print(f"\n✗ ビルドに失敗しました: {e}")
        sys.exit(1)


if __name__ == "__main__":
    build_exe()

