# ========================================================
#  Office PDF Binder - full build script
# ========================================================

# 現在の conda 環境を優先。未設定の場合はパスから python を検索。
if ($Env:CONDA_PREFIX -and (Test-Path (Join-Path $Env:CONDA_PREFIX "python.exe"))) {
    $PythonExe = Join-Path $Env:CONDA_PREFIX "python.exe"
} else {
    $PythonExe = (Get-Command python).Source
}

$ScriptName = "OfficePDFBinder_Main.py"
$IssFile = "setup_office_binder.iss"
$IconFile = "app.ico"
$SelfScriptName = "build_installer.ps1"

$ProductName = "Office PDF Binder"
$InternalName = "OfficePDFBinder_Main"
$VersionFile = "version.py"
if (-not (Test-Path $VersionFile)) {
    Write-Host "[ERROR] $VersionFile がありません" -ForegroundColor Red
    exit 1
}
$VersionContent = Get-Content $VersionFile -Raw
if ($VersionContent -notmatch 'APP_VERSION\s*=\s*"([^"]+)"') {
    Write-Host "[ERROR] $VersionFile から APP_VERSION を読み取れません。" -ForegroundColor Red
    exit 1
}
$AppVersion = $Matches[1]
$CompanyName = "Takeshi Kashiwagi"

Write-Host "`n========================================================" -ForegroundColor Cyan
Write-Host "[1/4] クリーンアップ..."
Write-Host "========================================================"

if (Test-Path "Output") { Remove-Item -Recurse -Force "Output" }
if (Test-Path "source.zip") { Remove-Item -Force "source.zip" }

# Nuitka の中間生成物と配布フォルダを削除し、毎回クリーンにビルドします。
if (Test-Path "$InternalName.build") { 
    Write-Host " - 古い中間ビルドフォルダを削除中..." 
    Remove-Item -Recurse -Force "$InternalName.build" 
}
if (Test-Path "$InternalName.dist") { 
    Write-Host " - 古いビルドフォルダを削除中..." 
    Remove-Item -Recurse -Force "$InternalName.dist" 
}


Write-Host "`n========================================================" -ForegroundColor Cyan
Write-Host "[2.5/4] README.md → README.html 変換..."
Write-Host "========================================================"

& $PythonExe "convert_readme.py"
if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERROR] README の HTML 変換に失敗しました。" -ForegroundColor Red
    exit $LASTEXITCODE
}
Write-Host " - README.html 生成完了" -ForegroundColor Green

Write-Host "`n========================================================" -ForegroundColor Cyan
Write-Host "[2/4] source.zip 作成..."
Write-Host "========================================================"

$SourceFiles = @(
    $ScriptName,
    $IssFile,
    $IconFile,
    $VersionFile,
    "LICENSE.txt",
    "NOTICE.txt",
    $SelfScriptName,
    "convert_readme.py",
    "README.md",
    "README.html",
    "docs\images"
)
$Missing = $false
foreach ($f in $SourceFiles) {
    if (-not (Test-Path $f)) { Write-Host "[ERROR] $f がありません" -ForegroundColor Red; $Missing=$true }
}
if ($Missing) { exit 1 }

Compress-Archive -Path $SourceFiles -DestinationPath "source.zip" -Force
Write-Host " - OK" -ForegroundColor Green

# source.zip はインストーラー同梱用に作成し、ローカルバックアップも残します。
$BackupDir = "Backup"
if (-not (Test-Path $BackupDir)) {
    New-Item -ItemType Directory -Path $BackupDir | Out-Null
}
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$BackupFileName = "${Timestamp}_source.zip"
$BackupPath = Join-Path $BackupDir $BackupFileName
Copy-Item -Path "source.zip" -Destination $BackupPath -Force
Write-Host " - バックアップ作成: $BackupPath" -ForegroundColor Green


Write-Host "`n========================================================" -ForegroundColor Cyan
Write-Host "[3/4] Nuitka ビルド (Ver $AppVersion)..."
Write-Host "========================================================"

$NuitkaArgs = @(
    "-m", "nuitka",
    "--standalone",
    "--enable-plugin=pyside6",
    "--windows-console-mode=disable",
    "--lto=no",
    "--output-dir=.",
    "--windows-icon-from-ico=$IconFile",
    "--product-name=$ProductName",
    "--company-name=$CompanyName",
    "--file-version=$AppVersion",
    "--product-version=$AppVersion",
    "--file-description=Office PDF Binder Tool",
    "--no-pyi-file",
    "--nofollow-import-to=tkinter",
    "--nofollow-import-to=unittest",
    "--nofollow-import-to=numpy",
    "--nofollow-import-to=pandas",
    "--nofollow-import-to=pyarrow",
    "--nofollow-import-to=openpyxl",
    "--nofollow-import-to=PySide6.QtDataVisualization",
    "--nofollow-import-to=PySide6.QtPdf",
    "--nofollow-import-to=PySide6.QtQuick",
    "--nofollow-import-to=PySide6.QtQml",
    $ScriptName
)

& $PythonExe $NuitkaArgs

if ($LASTEXITCODE -ne 0) {
    Write-Host "`n[ERROR] Nuitka ビルドに失敗しました。" -ForegroundColor Red
    exit $LASTEXITCODE
}
Write-Host " - ビルド完了" -ForegroundColor Green


Write-Host "`n========================================================" -ForegroundColor Cyan
Write-Host "[4/4] インストーラー作成..."
Write-Host "========================================================"

$IsccExe = "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
$IsccArgs = @("/DMyAppVersion=$AppVersion", $IssFile)

Write-Host "実行: $IsccExe $IsccArgs"

& $IsccExe $IsccArgs

if ($LASTEXITCODE -eq 0) {
    Write-Host "`n[SUCCESS] 完了しました！ Outputフォルダを確認してください。" -ForegroundColor Green
} else {
    Write-Host "`n[ERROR] インストーラー作成に失敗しました。" -ForegroundColor Red
}

Read-Host "Enterキーを押して終了..."
