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

内部実装は以下のように分割しています。

- 起動入口: `flag_auto_generator.py`
- GUI本体: `flag_auto_generator_app/gui.py`
- Excel処理: `flag_auto_generator_app/excel_ops.py`
- GUI共通部品: `flag_auto_generator_app/ui_helpers.py`

### EXEビルド（任意）

```powershell
py -3.12 .\build.py
```

## 使い方（要点）

1. 「Excelを選択」で元ファイルを読み込み、プレビューを確認
2. 基本設定は「シート名」を必要に応じて設定
3. 「工具と測定No対応」を登録
4. 必要に応じて「自動測定データ対応（測定No → データ順番）」を登録
5. 必要に応じて「測定不要」の測定 No を「追加」で1件ずつ登録
6. 「この内容で Excel を保存・生成」で出力

### 基本設定の補足

- 「測定不要書き込み設定の行」に入力した値を基準に内部設定を自動計算します
- 測定行(max): `入力値 - 1`
- 工具開始行: `入力値 + 3`
- 自動測定データ開始行: `入力値 + (工具数 * 3) + 6`
- そのほかの基本設定は固定値として内部で扱います

### 固定設定

以下の項目は固定値で、GUIからは編集できません。

- 測定No列: A
- 測定行(min): 11
- 測定行ステップ: 3
- 集計行(min): 11
- 集計行(max): 測定行(max) と連動
- 集計行ステップ: 3
- 数式区切り: ,
- 工具名列: E
- 工具行ステップ: 3

### 先頭集計式

- Excel の 1〜3 行目には集計式を設定します
- 1〜3 行目の基準式は例として次のとおりです（L 列の例。2・3 行目は開始行が 1 行ずつ下がる）  
  `=SUMPRODUCT(--(L11:L119<>""),--(MOD(ROW(L11:L119)-ROW(L11),3)=0))` など
- 1 行目がこの形式でない場合、L〜SN 列の 1〜3 行目を同パターンの式で補正します

## 必要な環境変数

- 必須の環境変数はありません

## 注意事項

- 元Excelおよび出力先Excelを開いたまま実行すると保存に失敗することがあります
- 測定Noは整数で入力してください
- 出力列はL〜SR固定です
- 工具行の同じ列に値が入ると、10行目も「依頼」表示になります
- `openpyxl` 由来の `UserWarning: Data Validation extension is not supported and will be removed` は警告表示のみで、処理停止ではありません
- Excel COM を使った強制再計算は既定で無効です。必要な場合のみ環境変数 `FLAG_AUTO_GENERATOR_FORCE_EXCEL_RECALC=1` を付けて起動してください
