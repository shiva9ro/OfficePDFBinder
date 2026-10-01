param([string]$Script, [string]$AppDir, [string]$Scenario)
# Never touch real registry, certificates or package registrations in this harness.
. $Script -ApplicationDirectory $AppDir
$script:Calls = [Collections.Generic.List[string]]::new()
$script:Removed = $false
$cert = [Security.Cryptography.X509Certificates.X509Certificate2]::new(
    (Join-Path $AppDir 'shell-integration\OfficePDFBinder.ContextMenu.cer'))
$script:Thumbprint = $cert.Thumbprint
function Get-ItemProperty {
    param($LiteralPath, $ErrorAction)
    if ($Scenario -like '*uninstall*') {
        return [pscustomobject]@{
            InstallationDirectory=$(if ($Scenario -eq 'other-uninstall') {'C:\Other'} else {$AppDir})
            OwnedCertificate=$(if ($Scenario -ne 'pretrusted-uninstall') {$script:Thumbprint})
        }
    }
}
function Test-Path {
    param($LiteralPath, $Path)
    $p = if ($LiteralPath) {$LiteralPath} else {$Path}
    return ($p -like 'Cert:*' -and ($Scenario -like '*uninstall*' -or $Scenario -eq 'pretrusted-install'))
}
function New-Item { param($Path,[switch]$Force) }
function Set-ItemProperty { param($LiteralPath,$Name,$Value); $script:Calls.Add("set:$Name=$Value") }
function Import-Certificate { param($FilePath,$CertStoreLocation); $script:Calls.Add("trust:$CertStoreLocation") }
function Remove-Item { param($LiteralPath,[switch]$Recurse,[switch]$Force); $script:Calls.Add("remove:$LiteralPath") }
function Add-AppxPackage {
    param($Path,$ExternalLocation,[switch]$ForceUpdateFromAnyVersion)
    $script:Calls.Add("add:$(Split-Path $Path -Leaf)")
}
function Get-AppxPackage {
    param($Name)
    if (-not $script:Removed) { [pscustomobject]@{PackageFullName='owned-package'} }
}
function Remove-AppxPackage {
    param([Parameter(ValueFromPipeline)]$Package)
    process { $script:Calls.Add('unregister'); $script:Removed=$true }
}
$Unregister = $Scenario -like '*uninstall*'
Invoke-ShellRegistration
ConvertTo-Json -InputObject @($script:Calls) -Compress
