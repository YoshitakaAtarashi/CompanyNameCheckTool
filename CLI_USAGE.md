# PowerPoint キーワード検出ツール (CLI版)

## 概要
Windowsコマンドライン上で動作する、PowerPointファイル内のキーワード検出専用ツールです。
指定したディレクトリ内のPPTファイルを検索し、キーワードの出現箇所をリスト表示します。

## 特徴
- ✅ コマンドラインで単体動作
- ✅ ディレクトリ指定で一括検査
- ✅ 再帰的なサブディレクトリ検索対応
- ✅ マスタースライド（複数グループ対応）
- ✅ 検出結果のテキスト出力
- ✅ 設定ファイル（config.json）共有

## 必要環境
- Python 3.7以上
- python-pptxライブラリ

## インストール
```powershell
# 既存の環境を使用（requirements.txtから）
pip install -r requirements.txt
```

## 使用方法

### 基本的な使い方
```powershell
# ディレクトリを指定して検査
python detect_keywords_cli.py C:\Documents\Presentations

# サブディレクトリを含めずに検査
python detect_keywords_cli.py C:\Documents\Presentations --no-recursive

# 特定のキーワードで検査
python detect_keywords_cli.py C:\Documents\Presentations --keywords "OldCompany" "旧社名"

# 結果をファイルに保存
python detect_keywords_cli.py C:\Documents\Presentations --output results.txt

# 検出数0のファイルも全て表示
python detect_keywords_cli.py C:\Documents\Presentations --show-all
```

### コマンドラインオプション

| オプション | 短縮形 | 説明 |
|-----------|--------|------|
| `directory` | - | 検索対象のディレクトリパス（必須） |
| `--keywords` | `-k` | 検索キーワード（スペース区切り） |
| `--no-recursive` | `-n` | サブディレクトリを検索しない |
| `--output` | `-o` | 結果を保存するファイル名 |
| `--show-all` | `-a` | 検出数0のファイルも含めて全ファイルを表示 |

### 使用例

#### 例1: 基本的な検索
```powershell
python detect_keywords_cli.py D:\Documents\PPT_Files
```

出力例:
```
================================================================================
PowerPoint キーワード検出ツール (CLI版)
================================================================================
検索ディレクトリ: D:\Documents\PPT_Files
検索キーワード: OldCompany, 旧社名, Old Company Name
再帰検索: はい
--------------------------------------------------------------------------------

PPTファイルを検索中...
15 件のPPTファイルが見つかりました。

[1/15] 検査中: presentation1.pptx ... ✓ 3 箇所で検出
[2/15] 検査中: presentation2.pptx ... 検出なし
[3/15] 検査中: document.pptx ... ✓ 5 箇所で検出
...


D:\Documents\PPT_Files\presentation1.pptx	3
D:\Documents\PPT_Files\document.pptx	5
D:\Documents\PPT_Files\slides.pptx	2

================================================================================
対象ディレクトリ: D:\Documents\PPT_Files
検出ファイル数: 3/15
実施日時: 2025-11-21 10:30:45
================================================================================

※ デフォルトでは検出があったファイルのみ表示されます。
※ 全ファイルを表示するには `--show-all` オプションを使用してください。
```

#### 例2: カスタムキーワードで検索
```powershell
python detect_keywords_cli.py D:\Documents --keywords "重要" "機密" --output check_results.txt
```

#### 例3: 直下のファイルのみ検索
```powershell
python detect_keywords_cli.py D:\Documents\PPT_Files --no-recursive
```

#### 例4: 全ファイルを表示（検出0のファイルも含む）
```powershell
python detect_keywords_cli.py D:\Documents\PPT_Files --show-all
```

出力例（--show-all使用時）:
```
D:\Documents\PPT_Files\presentation1.pptx	3
D:\Documents\PPT_Files\presentation2.pptx	0
D:\Documents\PPT_Files\document.pptx	5
D:\Documents\PPT_Files\report.pptx	0
D:\Documents\PPT_Files\slides.pptx	2

================================================================================
対象ディレクトリ: D:\Documents\PPT_Files
検出ファイル数: 3/5
実施日時: 2025-11-21 10:35:00
================================================================================
```

## 出力形式

検出結果は以下の形式で出力されます：

**デフォルト（検出があったファイルのみ）:**
```
[ファイルフルパス][タブ][検出箇所数]
[ファイルフルパス][タブ][検出箇所数]
...

================================================================================
対象ディレクトリ: [指定したディレクトリ]
検出ファイル数: [キーワードが見つかったファイル数]/[総ファイル数]
実施日時: [実行日時]
================================================================================
```

**`--show-all` オプション使用時（全ファイル表示）:**
- 検出数0のファイルも含めて全ファイルを表示

- 各行はファイルのフルパスとキーワード検出箇所数をタブ区切りで表示
- デフォルトでは検出があったファイルのみ表示
- 最後にサマリー情報（対象ディレクトリ、検出ファイル数、実施日時）を出力

## 設定ファイル
`config.json` を使用してデフォルト設定を管理します。
Webツール（app.py）と設定を共有します。

```json
{
  "default_keywords": [
    "OldCompany",
    "旧社名",
    "Old Company Name"
  ],
  "allowed_extensions": ["pptx", "ppt"]
}
```

## ファイル出力

`--output` オプションで結果をテキストファイルに保存できます：

```powershell
python detect_keywords_cli.py C:\Documents --output results.txt
```

保存されるファイルも同じ形式で、Excelなどでタブ区切りとして開くことができます。

## 検出対象
- 通常スライド内のテキスト
- マスタースライド（複数のマスターグループ対応）
- スライドレイアウト内のテキスト
- すべてのテキストシェイプ

## 制限事項
- 画像内のテキスト（OCR）は検出不可
- 検出のみ（置換・削除は不可）
- .pptx, .ppt形式のみ対応

## トラブルシューティング

### エラー: "ディレクトリが存在しません"
- パスが正しいか確認してください
- Windowsのパス区切りは `\` または `/` 両方使用可能

### エラー: "PPTファイルが見つかりませんでした"
- 指定ディレクトリに.pptxまたは.pptファイルがあるか確認
- `--no-recursive` オプションを外してサブディレクトリも検索

### ファイルが読み込めない
- ファイルが破損していないか確認
- ファイルが他のプログラムで開かれていないか確認
- PowerPoint形式が古い場合は.pptxに変換

## Webツールとの違い

| 機能 | CLI版 | Web版（app.py） |
|-----|-------|----------------|
| 実行環境 | コマンドライン | ブラウザ |
| キーワード検出 | ✅ | ✅ |
| キーワード置換 | ❌ | ✅ |
| キーワード削除 | ❌ | ✅ |
| バッチ処理 | ✅ | ✅ |
| GUI | ❌ | ✅ |
| 出力形式 | テキスト | JSON/ダウンロード |

## ライセンス
既存ツールと同じライセンスを適用
