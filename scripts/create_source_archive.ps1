param(
    [string]$DestinationPath = "source.zip"
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..")).Path
$DestinationFullPath = if ([IO.Path]::IsPathRooted($DestinationPath)) {
    $DestinationPath
} else {
    Join-Path $ProjectRoot $DestinationPath
}

$SourceFiles = [Collections.Generic.HashSet[string]]::new(
    [StringComparer]::OrdinalIgnoreCase
)

$RequiredFiles = @(
    ".gitignore",
    "OfficePDFBinder_Main.py",
    "i18n.py",
    "version.py",
    "packaging\setup_office_binder.iss",
    "app.ico",
    "LICENSE.txt",
    "NOTICE.txt",
    "build.ps1",
    "scripts\build_installer.ps1",
    "scripts\build_installer_only.ps1",
    "scripts\build_portable.ps1",
    "scripts\create_source_archive.ps1",
    "scripts\convert_readme.py",
    "requirements.txt",
    "requirements-dev.txt",
    "pytest.ini",
    "TESTING.md",
    "README.md",
    "README.ja.md",
    "README.html",
    "README.ja.html",
    "translations\OfficePDFBinder_en.ts"
)

foreach ($RelativePath in $RequiredFiles) {
    $FullPath = Join-Path $ProjectRoot $RelativePath
    if (-not (Test-Path -LiteralPath $FullPath -PathType Leaf)) {
        throw "Source archive file is missing: $RelativePath"
    }
    [void]$SourceFiles.Add($RelativePath)
}

foreach ($Tree in @("tests", ".github", "tools", "docs\images")) {
    $TreePath = Join-Path $ProjectRoot $Tree
    if (-not (Test-Path -LiteralPath $TreePath -PathType Container)) {
        throw "Source archive directory is missing: $Tree"
    }
    foreach ($File in Get-ChildItem -LiteralPath $TreePath -File -Recurse) {
        $RelativePath = $File.FullName.Substring($ProjectRoot.Length).TrimStart("\", "/")
        & git -C $ProjectRoot check-ignore --quiet -- $RelativePath
        if ($LASTEXITCODE -ne 0) {
            [void]$SourceFiles.Add($RelativePath)
        }
    }
}

if (Test-Path -LiteralPath $DestinationFullPath) {
    Remove-Item -LiteralPath $DestinationFullPath -Force
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$Stream = [IO.File]::Open($DestinationFullPath, [IO.FileMode]::CreateNew)
$Archive = [IO.Compression.ZipArchive]::new(
    $Stream,
    [IO.Compression.ZipArchiveMode]::Create
)
try {
    foreach ($RelativePath in ($SourceFiles | Sort-Object)) {
        $FullPath = Join-Path $ProjectRoot $RelativePath
        $EntryName = $RelativePath.Replace("\", "/")
        [IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
            $Archive,
            $FullPath,
            $EntryName,
            [IO.Compression.CompressionLevel]::Optimal
        ) | Out-Null
    }
} finally {
    $Archive.Dispose()
    $Stream.Dispose()
}

Write-Host "[SUCCESS] source.zipを作成しました（$($SourceFiles.Count)ファイル）。" -ForegroundColor Green
exit 0
