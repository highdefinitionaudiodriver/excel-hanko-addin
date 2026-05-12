@echo off
chcp 65001 >nul 2>&1
title Excel Hanko Add-in - 電子印鑑サーバー

echo ============================================================
echo   Excel Hanko Add-in - 電子印鑑（データ印）
echo   ローカル HTTPS サーバー起動ツール
echo ============================================================
echo.

:: Node.js の存在確認
where node >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo [エラー] Node.js が見つかりません。
    echo https://nodejs.org/ からインストールしてください（v18以上推奨）。
    echo.
    pause
    exit /b 1
)

:: バージョン表示
echo [情報] Node.js バージョン:
node --version
echo.

:: node_modules の存在確認 → なければ npm install
if not exist "node_modules" (
    echo [初回セットアップ] npm install を実行しています...
    echo （初回のみ数十秒かかります）
    echo.
    call npm install
    if %ERRORLEVEL% neq 0 (
        echo.
        echo [エラー] npm install に失敗しました。
        pause
        exit /b 1
    )
    echo.
    echo [完了] パッケージのインストールが完了しました。
    echo.
)

:: SSL証明書の生成
echo [準備] SSL証明書を確認しています...
call npx office-addin-dev-certs install
echo.

:: 証明書パス
set CERT_DIR=%USERPROFILE%\.office-addin-dev-certs
set CERT=%CERT_DIR%\localhost.crt
set KEY=%CERT_DIR%\localhost.key

if not exist "%CERT%" (
    echo [エラー] SSL証明書が見つかりません: %CERT%
    echo office-addin-dev-certs のインストールに失敗した可能性があります。
    pause
    exit /b 1
)

:: サーバー起動
echo ============================================================
echo   HTTPS サーバーを起動します
echo   URL: https://localhost:3000/taskpane.html
echo.
echo   ★ 次の手順:
echo     1. ブラウザで上記URLにアクセスして動作確認
echo     2. Excel Online で manifest.xml をサイドロード
echo     3. 終了するには Ctrl+C を押してください
echo ============================================================
echo.

npx http-server src -p 3000 -S -C "%CERT%" -K "%KEY%" --cors

pause
