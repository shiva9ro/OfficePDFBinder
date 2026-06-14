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

# 既存の Nuitka 配布フォルダは残し、インストーラー出力だけ更新します。
if (Test-Path "Output")    { Remove-Item -Recurse -Force "Output" }
if (Test-Path "source.zip"){ Remove-Item -Force "source.zip" }

Write-Host "`n========================================================" -ForegroundColor Cyan
Write-Host "[2/3] README.md → README.html 変換 & source.zip 作成..."
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

$SourceFiles = @(
    "OfficePDFBinder_Main.py",
    $IssFile,
    "app.ico",
    $VersionFile,
    "LICENSE.txt",
    "NOTICE.txt",
    "build_installer.ps1",
    $SelfScriptName,
    "convert_readme.py",
    "README.md",
    "README.html",
    "docs\images"
)

$Missing = $false
foreach ($f in $SourceFiles) {
    if (-not (Test-Path $f)) {
        Write-Host "[ERROR] $f がありません" -ForegroundColor Red
        $Missing = $true
    }
}
if ($Missing) { exit 1 }

Compress-Archive -Path $SourceFiles -DestinationPath "source.zip" -Force
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

Read-Host "Enterキーを押して終了..."


