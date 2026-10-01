param(
    [Parameter(Mandatory)][string]$ApplicationDirectory,
    [switch]$Unregister
)
$ErrorActionPreference = 'Stop'
$PackageName = 'OfficePDFBinder.ContextMenu'
$OwnerKey = 'HKCU:\Software\OfficePDFBinder\ShellIntegration'

function Remove-LegacyMenu {
    foreach ($Extension in @('.pdf','.doc','.docx','.docm','.xls','.xlsx','.xlsm','.ppt','.pptx','.pptm',
        '.png','.jpg','.jpeg','.bmp','.webp','.tif','.tiff','.heic','.heif','.hif','.svg')) {
        $Key = "HKCU:\Software\Classes\SystemFileAssociations\$Extension\shell\OfficePDFBinder"
        if (Test-Path -LiteralPath $Key) { Remove-Item -LiteralPath $Key -Recurse -Force }
    }
    # The delivery COM server is registered in the package now.
    foreach ($Kind in @('CLSID','AppID')) {
        $Key = "HKCU:\Software\Classes\$Kind\{F9869918-6F6D-4FC9-A8DF-70F43498A802}"
        if (Test-Path -LiteralPath $Key) { Remove-Item -LiteralPath $Key -Recurse -Force }
    }
}

function Invoke-ShellRegistration {
    $AppDir = (Resolve-Path -LiteralPath $ApplicationDirectory).Path.TrimEnd('\')
    $Support = Join-Path $AppDir 'shell-integration'
    $Certificate = [Security.Cryptography.X509Certificates.X509Certificate2]::new(
        (Join-Path $Support 'OfficePDFBinder.ContextMenu.cer'))
    if ($Certificate.Subject -ne 'CN=OfficePDFBinder.ContextMenu') { throw 'Unexpected certificate publisher' }
    $Thumbprint = $Certificate.Thumbprint
    $Store = 'Cert:\CurrentUser\TrustedPeople'
    $Owner = Get-ItemProperty -LiteralPath $OwnerKey -ErrorAction SilentlyContinue
    if ($Unregister) {
        if ($Owner -and $Owner.InstallationDirectory -ne $AppDir) { return }
        Get-AppxPackage -Name $PackageName | Remove-AppxPackage
        if ($Owner -and $Owner.OwnedCertificate -eq $Thumbprint -and
            @(Get-AppxPackage -Name $PackageName).Count -eq 0 -and (Test-Path "$Store\$Thumbprint")) {
            Remove-Item -LiteralPath "$Store\$Thumbprint"
        }
        if ($Owner) { Remove-Item -LiteralPath $OwnerKey -Recurse -Force }
        Remove-LegacyMenu
        return
    }
    Remove-LegacyMenu
    if ([Environment]::OSVersion.Version.Build -lt 22000) { return }
    New-Item -Path $OwnerKey -Force | Out-Null
    Set-ItemProperty -LiteralPath $OwnerKey -Name InstallationDirectory -Value $AppDir
    if (-not (Test-Path "$Store\$Thumbprint")) {
        Import-Certificate -FilePath (Join-Path $Support 'OfficePDFBinder.ContextMenu.cer') `
            -CertStoreLocation $Store | Out-Null
        Set-ItemProperty -LiteralPath $OwnerKey -Name OwnedCertificate -Value $Thumbprint
    }
    Add-AppxPackage -Path (Join-Path $Support 'OfficePDFBinder.ContextMenu.msix') `
        -ExternalLocation $AppDir -ForceUpdateFromAnyVersion
}

if ($MyInvocation.InvocationName -ne '.') {
    $LogFile = Join-Path $env:TEMP 'OfficePDFBinder-shell.log'
    try {
        "$(Get-Date -Format o) Unregister=$Unregister" | Add-Content $LogFile
        Invoke-ShellRegistration
        'SUCCESS' | Add-Content $LogFile
        exit 0
    }
    catch { $_ | Out-String | Add-Content $LogFile; Write-Error $_ -ErrorAction Continue; exit 1 }
}
