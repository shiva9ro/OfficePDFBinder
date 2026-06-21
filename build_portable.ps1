param(
    [string]$DistDir = "OfficePDFBinder_Main.dist",
    [string]$OutputDir = "Output"
)

# ========================================================
#  Office PDF Binder - restricted portable package
#  Reuses the same Nuitka dist as the installer build.
# ========================================================

$ErrorActionPreference = "Stop"
$VersionFile = "version.py"
$MarkerFile = "OfficePDFBinder.restricted-portable"

if (-not (Test-Path -LiteralPath $VersionFile -PathType Leaf)) {
    Write-Host "[ERROR] $VersionFile がありません。" -ForegroundColor Red
    exit 1
}

$VersionContent = Get-Content -LiteralPath $VersionFile -Raw
if ($VersionContent -notmatch 'APP_VERSION\s*=\s*"([^"]+)"') {
    Write-Host "[ERROR] $VersionFile から APP_VERSION を読み取れません。" -ForegroundColor Red
    exit 1
}
$AppVersion = $Matches[1]

if (-not (Test-Path -LiteralPath $DistDir -PathType Container)) {
    Write-Host "[ERROR] Nuitka配布フォルダーがありません: $DistDir" -ForegroundColor Red
    Write-Host "先に build_installer.ps1 でビルドしてください。" -ForegroundColor Yellow
    exit 1
}

& python ".\convert_readme.py"
if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERROR] README の HTML 変換に失敗しました。" -ForegroundColor Red
    exit $LASTEXITCODE
}

& ".\create_source_archive.ps1"
if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERROR] source.zip の作成に失敗しました。" -ForegroundColor Red
    exit $LASTEXITCODE
}

$RequiredFiles = @(
    "README.md",
    "README.ja.md",
    "README.html",
    "README.ja.html",
    "source.zip",
    "LICENSE.txt",
    "NOTICE.txt"
)
foreach ($File in $RequiredFiles) {
    if (-not (Test-Path -LiteralPath $File -PathType Leaf)) {
        Write-Host "[ERROR] 配布に必要なファイルがありません: $File" -ForegroundColor Red
        exit 1
    }
}

New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
$PackageName = "OfficePDFBinder_Portable_$AppVersion"
$PackageDir = Join-Path $OutputDir $PackageName
$ZipPath = Join-Path $OutputDir "$PackageName.zip"

if (Test-Path -LiteralPath $PackageDir) {
    Remove-Item -LiteralPath $PackageDir -Recurse -Force
}
if (Test-Path -LiteralPath $ZipPath) {
    Remove-Item -LiteralPath $ZipPath -Force
}

Copy-Item -LiteralPath $DistDir -Destination $PackageDir -Recurse
Set-Content -LiteralPath (Join-Path $PackageDir $MarkerFile) -Value "restricted-portable" -Encoding ascii
foreach ($File in $RequiredFiles) {
    Copy-Item -LiteralPath $File -Destination $PackageDir -Force
}

Compress-Archive -LiteralPath $PackageDir -DestinationPath $ZipPath -Force

Write-Host "[SUCCESS] 制限ポータブル版を作成しました。" -ForegroundColor Green
Write-Host " - フォルダー: $PackageDir"
Write-Host " - ZIP: $ZipPath"
