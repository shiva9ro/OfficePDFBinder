# ========================================================
#  Office PDF Binder - installer rebuild script
#  Reuses the existing OfficePDFBinder_Main.dist directory.
# ========================================================

$IssFile        = "setup_office_binder.iss"
$SelfScriptName = "build_installer_only.ps1"
$VersionFile    = "version.py"
if (-not (Test-Path $VersionFile)) {
    Write-Host "[ERROR] $VersionFile がありません" -ForegroundColor Red
    exit 1
}
$VersionContent = Get-Content $VersionFile -Raw
if ($VersionContent -notmatch 'APP_VERSION\s*=\s*"([^"]+)"') {
    Write-Host "[ERROR] $VersionFile から APP_VERSION を読み取れません。" -ForegroundColor Red
    exit 1
}
$AppVersion     = $Matches[1]

# 現在の conda 環境を優先。未設定の場合はパスから python を検索。
if ($Env:CONDA_PREFIX -and (Test-Path (Join-Path $Env:CONDA_PREFIX "python.exe"))) {
    $PythonExe = Join-Path $Env:CONDA_PREFIX "python.exe"
} else {
    $PythonExe = (Get-Command python).Source
}

Write-Host "`n========================================================" -ForegroundColor Cyan
Write-Host "[1/3] クリーンアップ（dist は削除しません）..."
Write-Host "========================================================"

# 既存の Nuitka 配布フォルダとポータブル成果物は残し、インストーラーだけ更新します。
if (Test-Path "Output") {
    Get-ChildItem -LiteralPath "Output" -Filter "OfficePDFBinder_Setup_*.exe" -File |
        Remove-Item -Force
}
if (Test-Path "source.zip"){ Remove-Item -Force "source.zip" }

Write-Host "`n========================================================" -ForegroundColor Cyan
Write-Host "[2/3] 日英README → HTMLマニュアル変換 & source.zip 作成..."
Write-Host "========================================================"

if (Test-Path "convert_readme.py") {
    & $PythonExe "convert_readme.py"
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[ERROR] README の HTML 変換に失敗しました。" -ForegroundColor Red
        exit $LASTEXITCODE
    }
    Write-Host " - README.html 生成完了" -ForegroundColor Green
} else {
    Write-Host "[WARN] convert_readme.py が見つからないため、README.html の再生成をスキップします。" -ForegroundColor Yellow
}

& ".\create_source_archive.ps1"
if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERROR] source.zip の作成に失敗しました。" -ForegroundColor Red
    exit $LASTEXITCODE
}
Write-Host " - source.zip 作成完了" -ForegroundColor Green

Write-Host "`n========================================================" -ForegroundColor Cyan
Write-Host "[3/3] Inno Setup でインストーラーのみ作成..."
Write-Host "========================================================"

$IsccExe  = "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
$IsccArgs = @("/DMyAppVersion=$AppVersion", $IssFile)

Write-Host "実行: $IsccExe $IsccArgs"
& $IsccExe $IsccArgs

if ($LASTEXITCODE -eq 0) {
    Write-Host "`n[SUCCESS] インストーラーの再生成が完了しました。Output フォルダを確認してください。" -ForegroundColor Green
} else {
    Write-Host "`n[ERROR] インストーラー作成に失敗しました。" -ForegroundColor Red
}

