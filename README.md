# flag_auto_generator

Excel検査シート向けに、以下を自動で設定するWindows向けGUIツールです。

- 工具と測定No対応に基づく「依頼」式の生成（L〜SR列）
- 自動測定結果の反映式生成（測定No → データ順番）
- 測定不要設定の書き込み（E列の「測定不要」およびL〜SR列の「-」式）

## 動作環境

- OS: Windows
- Python: 3.12 推奨
- Excel: 再計算処理に使用（`pywin32` 経由）

## セットアップ手順

1. 仮想環境を作成

```powershell
py -3.12 -m venv .venv
```

2. 仮想環境を有効化

```powershell
.\.venv\Scripts\Activate.ps1
```

3. 依存パッケージをインストール

```powershell
pip install -r requirements.txt
```

## 実行方法

### GUIアプリ起動

```powershell
py -3.12 .\flag_auto_generator.py
```

### EXEビルド（任意）

```powershell
py -3.12 .\build.py
```

## 使い方（要点）

1. 「Excelを選択」で元ファイルを読み込み、プレビューを確認
2. 基本設定（シート名、測定No列、行範囲、工具開始行など）を設定
3. 「工具と測定No対応」を登録
4. 必要に応じて「自動測定データ対応（測定No → データ順番）」を登録
5. 必要に応じて「測定不要書き込み設定」にNoを入力
6. 「この設定でExcel生成」で出力

## 必要な環境変数

- 必須の環境変数はありません

## 注意事項

- 元Excelおよび出力先Excelを開いたまま実行すると保存に失敗することがあります
- 測定Noは整数で入力してください
- 出力列はL〜SR固定です
- `pywin32` が未導入の場合、Excel強制再計算はスキップされます（ファイル出力自体は実行されます）
