# ========================================================
#  Office PDF Binder - build entry point
# ========================================================

param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("Package", "Fast", "Release")]
    [string]$Mode
)

$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot

$PythonExe = if (
    $Env:CONDA_PREFIX -and
    (Test-Path -LiteralPath (Join-Path $Env:CONDA_PREFIX "python.exe") -PathType Leaf)
) {
    Join-Path $Env:CONDA_PREFIX "python.exe"
} else {
    (Get-Command python).Source
}

$BuildStartedAt = Get-Date
$BuildStopwatch = [Diagnostics.Stopwatch]::StartNew()
$BuildExitCode = 0
$BuildErrorMessage = $null

try {
    switch ($Mode) {
        "Package" {
            Write-Host "[BUILD] Package: existing dist -> installer" -ForegroundColor Cyan
            & ".\scripts\build_installer_only.ps1"
        }
        "Fast" {
            Write-Host "[BUILD] Fast (Clang): incremental Nuitka build -> portable + installer" -ForegroundColor Cyan
            & ".\scripts\build_installer.ps1" -Fast
        }
        "Release" {
            Write-Host "[BUILD] Release (Clang): clean Nuitka build -> portable + installer" -ForegroundColor Cyan
            & ".\scripts\build_installer.ps1"
        }
    }

    if ($LASTEXITCODE -ne 0) {
        $BuildExitCode = $LASTEXITCODE
    }
} catch {
    $BuildExitCode = 1
    $BuildErrorMessage = $_.Exception.Message
} finally {
    $BuildStopwatch.Stop()
    $BuildFinishedAt = Get-Date

    $PythonVersion = (& $PythonExe --version 2>&1) -join " "
    $NuitkaVersion = (
        & $PythonExe -c "import importlib.metadata as m; print(m.version('Nuitka'))"
    ) -join " "
    $PyMuPDFVersion = (
        & $PythonExe -c "import importlib.metadata as m; print(m.version('PyMuPDF'))"
    ) -join " "
    $MetricsDirectory = Join-Path $PSScriptRoot "Backup\build-metrics"
    New-Item -ItemType Directory -Path $MetricsDirectory -Force | Out-Null
    $MetricsPath = Join-Path $MetricsDirectory (
        "build_{0}_{1}.json" -f $BuildStartedAt.ToString("yyyyMMdd_HHmmss"), $Mode
    )

    [pscustomobject]@{
        mode = $Mode
        status = if ($BuildExitCode -eq 0) { "Success" } else { "Failed" }
        exit_code = $BuildExitCode
        started_at = $BuildStartedAt.ToString("o")
        finished_at = $BuildFinishedAt.ToString("o")
        elapsed_seconds = [Math]::Round($BuildStopwatch.Elapsed.TotalSeconds, 3)
        elapsed = $BuildStopwatch.Elapsed.ToString("c")
        python_executable = $PythonExe
        python_version = $PythonVersion
        nuitka_version = $NuitkaVersion
        pymupdf_version = $PyMuPDFVersion
        c_compiler = if ($Mode -eq "Package") { $null } else { "clang-cl" }
        error = $BuildErrorMessage
    } |
        ConvertTo-Json |
        Set-Content -LiteralPath $MetricsPath -Encoding utf8

    Write-Host (
        "[BUILD] Elapsed: {0:hh\:mm\:ss\.fff}" -f $BuildStopwatch.Elapsed
    ) -ForegroundColor Cyan
    Write-Host "[BUILD] Metrics: $MetricsPath" -ForegroundColor Cyan
}

if ($BuildErrorMessage) {
    Write-Host "[ERROR] $BuildErrorMessage" -ForegroundColor Red
}
if ($BuildExitCode -ne 0) {
    exit $BuildExitCode
}
