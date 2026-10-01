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
Enter-VsDevShell -VsInstallPath $VsPath -SkipAutomaticLocation -DevCmdArguments "-arch=$Architecture -host_arch=x64" | Out-Null
$OutputPath = if ([IO.Path]::IsPathRooted($OutputDirectory)) { $OutputDirectory } else { Join-Path $ProjectRoot $OutputDirectory }
New-Item -ItemType Directory -Force -Path $OutputPath | Out-Null
$SourcePath = Join-Path $ProjectRoot "native\shell_bridge"
Push-Location $OutputPath
try {
    if ($ExplorerCommand) {
        & cl.exe /nologo /LD /std:c++17 /EHsc /W4 /WX /O2 /MT /utf-8 /DUNICODE /D_UNICODE /DWIN32_LEAN_AND_MEAN /DNOMINMAX `
            (Join-Path $SourcePath 'explorer_command.cpp') /Fe:OfficePDFBinder_Explorer.dll `
            /link /DYNAMICBASE /NXCOMPAT ole32.lib shell32.lib shlwapi.lib uuid.lib
        if ($LASTEXITCODE -ne 0) { throw 'ExplorerCommand DLL のビルドに失敗しました。' }
        return
    }
    $Entry = if ($Tests) { "test_driver.cpp" } else { "main.cpp" }
    $Exe = if ($Tests) { "shell_bridge_test.exe" } else { "OfficePDFBinder_Shell.exe" }
    $Defines = @()
    if ($IntegrationServer) { $Exe = 'shell_bridge_integration.exe'; $Defines += '/DBINDER_INTEGRATION_TEST' }
    $Subsystem = if ($Tests) { '/SUBSYSTEM:CONSOLE' } else { '/SUBSYSTEM:WINDOWS' }
    & cl.exe /nologo /std:c++17 /EHsc /W4 /WX /O2 /MT /utf-8 /DUNICODE /D_UNICODE /DWIN32_LEAN_AND_MEAN /DNOMINMAX @Defines `
        (Join-Path $SourcePath $Entry) (Join-Path $SourcePath 'delivery.cpp') "/Fe:$Exe" `
        /link $Subsystem /DYNAMICBASE /NXCOMPAT ole32.lib shell32.lib uuid.lib user32.lib advapi32.lib
    if ($LASTEXITCODE -ne 0) { throw "Shell bridge のビルドに失敗しました。" }
} finally { Pop-Location }
