param(
    [string]$OutputDirectory = "native\shell_bridge\out.build",
    [ValidateSet('x64', 'arm64')][string]$Architecture = 'x64',
    [switch]$Tests,
    [switch]$IntegrationServer,
    [switch]$ExplorerCommand
)
$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$VsWhere = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
if (-not (Test-Path -LiteralPath $VsWhere)) { throw "Visual Studio C++ Build Tools が必要です。" }
$VsPath = & $VsWhere -latest -products '*' -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath
if (-not $VsPath) { throw "MSVC x64 と Windows SDK をインストールしてください。" }
# Import the toolchain environment without constructing shell command strings.
Import-Module (Join-Path $VsPath "Common7\Tools\Microsoft.VisualStudio.DevShell.dll")
Enter-VsDevShell -VsInstallPath $VsPath -SkipAutomaticLocation -Arch $Architecture -HostArch x64 | Out-Null
if ($env:VSCMD_ARG_TGT_ARCH -ne $Architecture) {
    throw "MSVC target architecture mismatch: expected $Architecture, got $env:VSCMD_ARG_TGT_ARCH"
}

# Resolve the compiler explicitly rather than reusing cl.exe from PATH.
$Compiler = Join-Path $env:VCToolsInstallDir "bin\Hostx64\$Architecture\cl.exe"
if (-not (Test-Path -LiteralPath $Compiler -PathType Leaf)) {
    throw "MSVC compiler not found: $Compiler"
}

function Assert-NativeArchitecture([string]$Path, [string]$ExpectedArchitecture) {
    $Bytes = [IO.File]::ReadAllBytes($Path)
    if ($Bytes.Length -lt 64 -or $Bytes[0] -ne 0x4D -or $Bytes[1] -ne 0x5A) {
        throw "Invalid PE file: $Path"
    }
    $Offset = [BitConverter]::ToInt32($Bytes, 0x3C)
    if ($Offset -lt 0 -or $Offset -gt $Bytes.Length - 6 -or
        [BitConverter]::ToUInt32($Bytes, $Offset) -ne 0x00004550) {
        throw "Invalid PE header: $Path"
    }
    $Machine = [BitConverter]::ToUInt16($Bytes, $Offset + 4)
    $ExpectedMachine = if ($ExpectedArchitecture -eq 'arm64') { 0xAA64 } else { 0x8664 }
    if ($Machine -ne $ExpectedMachine) {
        throw ("Native architecture mismatch: {0}; expected {1}, got 0x{2:X4}" -f $Path, $ExpectedArchitecture, $Machine)
    }
}
$OutputPath = if ([IO.Path]::IsPathRooted($OutputDirectory)) { $OutputDirectory } else { Join-Path $ProjectRoot $OutputDirectory }
New-Item -ItemType Directory -Force -Path $OutputPath | Out-Null
$SourcePath = Join-Path $ProjectRoot "native\shell_bridge"
Push-Location $OutputPath
try {
    if ($ExplorerCommand) {
        & $Compiler /nologo /LD /std:c++17 /EHsc /W4 /WX /O2 /MT /utf-8 /DUNICODE /D_UNICODE /DWIN32_LEAN_AND_MEAN /DNOMINMAX `
            (Join-Path $SourcePath 'explorer_command.cpp') /Fe:OfficePDFBinder_Explorer.dll `
            /link /DYNAMICBASE /NXCOMPAT ole32.lib shell32.lib shlwapi.lib uuid.lib
        if ($LASTEXITCODE -ne 0) { throw 'ExplorerCommand DLL のビルドに失敗しました。' }
        Assert-NativeArchitecture (Join-Path $OutputPath 'OfficePDFBinder_Explorer.dll') $Architecture
        return
    }
    $Entry = if ($Tests) { "test_driver.cpp" } else { "main.cpp" }
    $Exe = if ($Tests) { "shell_bridge_test.exe" } else { "OfficePDFBinder_Shell.exe" }
    $Defines = @()
    if ($IntegrationServer) { $Exe = 'shell_bridge_integration.exe'; $Defines += '/DBINDER_INTEGRATION_TEST' }
    $Subsystem = if ($Tests) { '/SUBSYSTEM:CONSOLE' } else { '/SUBSYSTEM:WINDOWS' }
    & $Compiler /nologo /std:c++17 /EHsc /W4 /WX /O2 /MT /utf-8 /DUNICODE /D_UNICODE /DWIN32_LEAN_AND_MEAN /DNOMINMAX @Defines `
        (Join-Path $SourcePath $Entry) (Join-Path $SourcePath 'delivery.cpp') "/Fe:$Exe" `
        /link $Subsystem /DYNAMICBASE /NXCOMPAT ole32.lib shell32.lib uuid.lib user32.lib advapi32.lib
    if ($LASTEXITCODE -ne 0) { throw "Shell bridge のビルドに失敗しました。" }
    Assert-NativeArchitecture (Join-Path $OutputPath $Exe) $Architecture
} finally { Pop-Location }
