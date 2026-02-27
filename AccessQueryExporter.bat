@echo off
setlocal enabledelayedexpansion

REM ================================================
REM   Access クエリエクスポーター（ローカル版）
REM ================================================
REM   使い方:
REM     1. .accdb / .mdb ファイルをこのバッチファイルにドラッグ＆ドロップ
REM     2. または、ダブルクリックしてファイル選択ダイアログから選択
REM ================================================

set "SCRIPT_DIR=%~dp0"
set "VBS_PATH=%SCRIPT_DIR%ExtractToJSON.vbs"
set "VIEWER_PATH=%SCRIPT_DIR%viewer.html"

REM VBSファイルの存在チェック
if not exist "%VBS_PATH%" (
    echo エラー: ExtractToJSON.vbs が見つかりません。
    echo このバッチファイルと同じフォルダに配置してください。
    pause
    exit /b 1
)

REM 引数チェック
if "%~1"=="" (
    REM 引数なし → ファイル選択ダイアログ
    echo.
    echo ========================================
    echo   Access クエリエクスポーター
    echo ========================================
    echo.
    echo ファイル選択ダイアログを開いています...
    echo.

    REM PowerShellでファイル選択ダイアログを表示
    for /f "delims=" %%F in ('powershell -NoProfile -Command "Add-Type -AssemblyName System.Windows.Forms; $d = New-Object System.Windows.Forms.OpenFileDialog; $d.Filter = 'Access Database (*.accdb;*.mdb)|*.accdb;*.mdb'; $d.Title = 'Access ファイルを選択してください'; if ($d.ShowDialog() -eq 'OK') { $d.FileName }"') do set "DB_PATH=%%F"

    if "!DB_PATH!"=="" (
        echo キャンセルされました。
        timeout /t 2 >nul
        exit /b 0
    )
) else (
    set "DB_PATH=%~1"
)

REM 拡張子チェック
set "EXT=%DB_PATH:~-6%"
echo !EXT! | findstr /i ".accdb" >nul 2>&1
if errorlevel 1 (
    set "EXT4=%DB_PATH:~-4%"
    echo !EXT4! | findstr /i ".mdb" >nul 2>&1
    if errorlevel 1 (
        echo.
        echo エラー: 対応していないファイル形式です。
        echo .accdb または .mdb ファイルを指定してください。
        echo.
        pause
        exit /b 1
    )
)

echo.
echo ========================================
echo   Access クエリエクスポーター
echo ========================================
echo.
echo 解析中: %DB_PATH%
echo.

REM VBScript実行
cscript //nologo "%VBS_PATH%" "%DB_PATH%"

if errorlevel 1 (
    echo.
    echo エラーが発生しました。
    pause
    exit /b 1
)

REM JSON結果ファイルパスを生成
for %%A in ("%DB_PATH%") do set "BASENAME=%%~nA"
set "JSON_PATH=%SCRIPT_DIR%%BASENAME%_result.json"

REM viewer.htmlをブラウザで開く（JSONファイルパスをクエリパラメータとして渡す）
if exist "%VIEWER_PATH%" (
    echo.
    echo ブラウザでビューアを開いています...
    start "" "%VIEWER_PATH%"
) else (
    echo.
    echo 注意: viewer.html が見つかりません。テキスト出力のみ完了しました。
)

echo.
echo ========================================
echo   完了！
echo ========================================
echo.
echo テキスト出力: %SCRIPT_DIR%%BASENAME%_クエリ一覧.txt
echo JSON出力: %JSON_PATH%
echo.
timeout /t 5 >nul
