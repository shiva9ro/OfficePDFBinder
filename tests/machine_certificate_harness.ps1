param([string]$Script, [string]$AppDir, [string]$Scenario)
# Dot sourcing defines functions only. All external state operations are mocked.
. $Script -ApplicationDirectory $AppDir
$script:Calls = [Collections.Generic.List[string]]::new()
function Get-ShellIntegrationInfo {
    [pscustomobject]@{ CertificatePath='mock.cer'; Thumbprint='TEST' }
}
function Test-Path {
    param($LiteralPath)
    return $Scenario -ne 'machine-install' -and $Scenario -ne 'machine-absent-uninstall'
}
function Import-Certificate {
    param($FilePath, $CertStoreLocation)
    $script:Calls.Add("trust:$CertStoreLocation")
}
function Remove-Item {
    param($LiteralPath, [switch]$Force)
    $script:Calls.Add("remove:$LiteralPath")
    if ($Scenario -eq 'machine-delete-error') { throw 'Simulated certificate deletion failure' }
}
function Get-AppxPackage {
    [CmdletBinding()]
    param([switch]$AllUsers, $Name)
    if (-not $AllUsers) { throw 'Expected an all-users query' }
    $script:Calls.Add('query:all-users')
    if ($Scenario -eq 'machine-query-error') {
        Write-Error 'Simulated package query failure'
        return
    }
    if ($Scenario -eq 'machine-shared-uninstall') {
        [pscustomobject]@{ PackageFullName='another-user-package' }
    }
}
$Failure = $null
try {
    if ($Scenario -in @('machine-install', 'machine-pretrusted-install')) {
        Install-MachineCertificate
    } else {
        Uninstall-MachineCertificate
    }
} catch { $Failure = $_.Exception.Message }
ConvertTo-Json -InputObject @{ calls=@($script:Calls); error=$Failure } -Compress
