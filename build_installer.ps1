# ========================================================
#  Office PDF Binder - full build script
# ========================================================

param(
    [switch]$Fast
)

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
$PortableScriptName = "build_portable.ps1"
$TranslationSource = "translations\OfficePDFBinder_en.ts"
$TranslationBinary = "translations\OfficePDFBinder_en.qm"

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
Write-Host "[1/5] クリーンアップ..."
Write-Host "========================================================"

if (Test-Path "Output") { Remove-Item -Recurse -Force "Output" }
if (Test-Path "source.zip") { Remove-Item -Force "source.zip" }

# 高速ビルドではNuitkaの中間生成物を再利用する。
# distは古いDLL等が残らないよう、どちらのモードでも作り直す。
if (-not $Fast -and (Test-Path "$InternalName.build")) {
    Write-Host " - 古い中間ビルドフォルダを削除中..."
    Remove-Item -Recurse -Force "$InternalName.build"
} elseif ($Fast -and (Test-Path "$InternalName.build")) {
    Write-Host " - 高速ビルド: 中間ビルドフォルダを再利用します。" -ForegroundColor Yellow
} elseif ($Fast) {
    Write-Host " - 高速ビルド: 再利用できる中間生成物がないため、初回は通常速度です。" -ForegroundColor Yellow
}
if (Test-Path "$InternalName.dist") { 
    Write-Host " - 古いビルドフォルダを削除中..." 
    Remove-Item -Recurse -Force "$InternalName.dist" 
}

Write-Host "`n========================================================" -ForegroundColor Cyan
Write-Host "[1.5/5] Qt翻訳ファイルを更新・コンパイル..."
Write-Host "========================================================"

$PythonDir = Split-Path -Parent $PythonExe
$LUpdateExe = Join-Path $PythonDir "Scripts\pyside6-lupdate.exe"
$LReleaseExe = Join-Path $PythonDir "Scripts\pyside6-lrelease.exe"
foreach ($Tool in @($LUpdateExe, $LReleaseExe)) {
    if (-not (Test-Path -LiteralPath $Tool -PathType Leaf)) {
        Write-Host "[ERROR] Qt翻訳ツールがありません: $Tool" -ForegroundColor Red
        exit 1
    }
}

& $LUpdateExe $ScriptName -ts $TranslationSource -source-language ja_JP -target-language en_US
if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERROR] 翻訳対象文字列の更新に失敗しました。" -ForegroundColor Red
    exit $LASTEXITCODE
}
& $LReleaseExe $TranslationSource -qm $TranslationBinary -nounfinished
if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERROR] Qt翻訳ファイルのコンパイルに失敗しました。" -ForegroundColor Red
    exit $LASTEXITCODE
}


Write-Host "`n========================================================" -ForegroundColor Cyan
Write-Host "[2/5] README.md → README.html 変換..."
Write-Host "========================================================"

& $PythonExe "convert_readme.py"
if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERROR] README の HTML 変換に失敗しました。" -ForegroundColor Red
    exit $LASTEXITCODE
}
Write-Host " - README.html 生成完了" -ForegroundColor Green

Write-Host "`n========================================================" -ForegroundColor Cyan
Write-Host "[3/5] source.zip 作成..."
Write-Host "========================================================"

$SourceFiles = @(
    $ScriptName,
    "i18n.py",
    $IssFile,
    $IconFile,
    $VersionFile,
    "LICENSE.txt",
    "NOTICE.txt",
    "build.ps1",
    $SelfScriptName,
    $PortableScriptName,
    "convert_readme.py",
    "README.md",
    "README.html",
    "README.en.md",
    "README.en.html",
    $TranslationSource,
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
$BuildMode = if ($Fast) { "高速" } else { "クリーン" }
Write-Host "[4/5] Nuitka $BuildMode ビルド (Ver $AppVersion)..."
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
    "--include-data-file=$TranslationBinary=$TranslationBinary",
    $ScriptName
)

& $PythonExe $NuitkaArgs

if ($LASTEXITCODE -ne 0) {
    Write-Host "`n[ERROR] Nuitka ビルドに失敗しました。" -ForegroundColor Red
    exit $LASTEXITCODE
}
Write-Host " - ビルド完了" -ForegroundColor Green


Write-Host "`n========================================================" -ForegroundColor Cyan
Write-Host "[4.5/5] 制限ポータブル版作成..."
Write-Host "========================================================"

& ".\$PortableScriptName"
if ($LASTEXITCODE -ne 0) {
    Write-Host "`n[ERROR] 制限ポータブル版の作成に失敗しました。" -ForegroundColor Red
    exit $LASTEXITCODE
}


Write-Host "`n========================================================" -ForegroundColor Cyan
Write-Host "[5/5] インストーラー作成..."
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
