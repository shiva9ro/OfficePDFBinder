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

switch ($Mode) {
    "Package" {
        Write-Host "[BUILD] Package: existing dist -> installer" -ForegroundColor Cyan
        & ".\scripts\build_installer_only.ps1"
    }
    "Fast" {
        Write-Host "[BUILD] Fast: incremental Nuitka build -> portable + installer" -ForegroundColor Cyan
        & ".\scripts\build_installer.ps1" -Fast
    }
    "Release" {
        Write-Host "[BUILD] Release: clean Nuitka build -> portable + installer" -ForegroundColor Cyan
        & ".\scripts\build_installer.ps1"
    }
}

if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}
