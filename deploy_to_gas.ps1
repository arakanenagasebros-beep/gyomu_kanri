# GAS自動デプロイスクリプト
# 使い方: PowerShellで右クリック → 「PowerShellで実行」

$ScriptId = "15wOroyFkYN_4LUgEVurK9ZTIykTbAKaWoeWVmo7gfsykEtLHbzI64Lbw"
# PSScriptRootが空の場合（右クリック実行時）はスクリプト自身のパスから取得
if ($PSScriptRoot -and (Test-Path $PSScriptRoot)) {
    $ProjectDir = $PSScriptRoot
} else {
    $ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
}
if (-not $ProjectDir) {
    $ProjectDir = "C:\Users\admin\Documents\Claude code\moc_staff"
}

Write-Host "=== GAS自動デプロイ ===" -ForegroundColor Cyan
Write-Host "プロジェクトフォルダ: $ProjectDir"

# Node.js確認
if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
    Write-Host "[エラー] Node.jsがインストールされていません。" -ForegroundColor Red
    Write-Host "https://nodejs.org からインストールしてください。"
    Read-Host "Enterキーで終了"
    exit 1
}

# clasp確認・インストール
if (-not (Get-Command clasp -ErrorAction SilentlyContinue)) {
    Write-Host "claspをインストール中..." -ForegroundColor Yellow
    npm install -g @google/clasp
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[エラー] claspのインストールに失敗しました。" -ForegroundColor Red
        Read-Host "Enterキーで終了"
        exit 1
    }
    Write-Host "claspインストール完了!" -ForegroundColor Green
}

# .clasp.json 作成
$ClaspConfig = @{
    scriptId = $ScriptId
    rootDir = $ProjectDir
} | ConvertTo-Json

$ClaspConfigPath = Join-Path $ProjectDir ".clasp.json"
$ClaspConfig | Set-Content $ClaspConfigPath -Encoding UTF8
Write-Host ".clasp.json を作成しました。" -ForegroundColor Green

# appsscript.json 確認・作成
$AppScriptPath = Join-Path $ProjectDir "appsscript.json"
if (-not (Test-Path $AppScriptPath)) {
    $AppScriptJson = @{
        timeZone = "Asia/Tokyo"
        dependencies = @{}
        exceptionLogging = "STACKDRIVER"
        runtimeVersion = "V8"
    } | ConvertTo-Json
    $AppScriptJson | Set-Content $AppScriptPath -Encoding UTF8
    Write-Host "appsscript.json を作成しました。" -ForegroundColor Green
}

# Google認証
Write-Host ""
Write-Host "=== Google認証 ===" -ForegroundColor Cyan
Write-Host "ブラウザが開きます。Googleアカウントでログインしてください。"
Write-Host "(既にログイン済みの場合はスキップされます)"
clasp login

if ($LASTEXITCODE -ne 0) {
    Write-Host "[エラー] 認証に失敗しました。" -ForegroundColor Red
    Read-Host "Enterキーで終了"
    exit 1
}

# プッシュ実行
Write-Host ""
Write-Host "=== GASへプッシュ中... ===" -ForegroundColor Cyan
Set-Location $ProjectDir
clasp push --force

if ($LASTEXITCODE -eq 0) {
    Write-Host ""
    Write-Host "=== プッシュ成功! ===" -ForegroundColor Green
    Write-Host ""
    Write-Host "次のステップ: resetAdminCredsToSuugaku を実行してください" -ForegroundColor Yellow
    Write-Host "1. GASエディタを開く: https://script.google.com/home/projects/$ScriptId/edit"
    Write-Host "2. 関数ドロップダウンで 'resetAdminCredsToSuugaku' を選択"
    Write-Host "3. ▶ 実行ボタンをクリック"
    Start-Process "https://script.google.com/home/projects/$ScriptId/edit?hl=ja"
} else {
    Write-Host "[エラー] プッシュに失敗しました。" -ForegroundColor Red
}

Read-Host "Enterキーで終了"
