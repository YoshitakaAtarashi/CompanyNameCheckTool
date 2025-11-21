@echo off
chcp 65001 > nul
echo ================================================================================
echo PowerPoint キーワード検出ツール (CLI版) - クイックスタート
echo ================================================================================
echo.

REM Pythonがインストールされているか確認
python --version >nul 2>&1
if errorlevel 1 (
    echo エラー: Pythonがインストールされていません。
    echo Python 3.7以上をインストールしてください。
    pause
    exit /b 1
)

echo Pythonが見つかりました。
echo.

REM 検索ディレクトリを入力
set /p TARGET_DIR="検索対象のディレクトリパスを入力してください: "

if not exist "%TARGET_DIR%" (
    echo.
    echo エラー: 指定されたディレクトリが存在しません: %TARGET_DIR%
    pause
    exit /b 1
)

echo.
echo 検索中のディレクトリ: %TARGET_DIR%
echo.

REM 再帰検索の確認
set /p RECURSIVE="サブディレクトリも検索しますか？ (Y/N, デフォルト: Y): "
if /i "%RECURSIVE%"=="N" (
    set RECURSIVE_OPT=--no-recursive
) else (
    set RECURSIVE_OPT=
)

REM 出力ファイルの確認
set /p SAVE_FILE="結果をファイルに保存しますか？ (ファイル名を入力、不要なら Enter): "

if "%SAVE_FILE%"=="" (
    set OUTPUT_OPT=
) else (
    set OUTPUT_OPT=--output "%SAVE_FILE%"
)

echo.
echo ================================================================================
echo 検査を開始します...
echo ================================================================================
echo.

REM ツールを実行
python detect_keywords_cli.py "%TARGET_DIR%" %RECURSIVE_OPT% %OUTPUT_OPT%

echo.
echo ================================================================================
echo 完了しました。
echo ================================================================================
pause
