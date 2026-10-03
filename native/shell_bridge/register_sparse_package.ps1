param(
    [Parameter(Mandatory)][string]$ApplicationDirectory,
    [switch]$Unregister,
    [switch]$InstallMachineCertificate,
    [switch]$UninstallMachineCertificate
)
$ErrorActionPreference = 'Stop'
$PackageName = 'OfficePDFBinder.ContextMenu'
$OwnerKey = 'HKCU:\Software\OfficePDFBinder\ShellIntegration'
$CurrentUserStore = 'Cert:\CurrentUser\TrustedPeople'
$MachineStore = 'Cert:\LocalMachine\TrustedPeople'

function Get-ShellIntegrationInfo {
    $AppDir = (Resolve-Path -LiteralPath $ApplicationDirectory).Path.TrimEnd('\')
    $Support = Join-Path $AppDir 'shell-integration'
    $CertificatePath = Join-Path $Support 'OfficePDFBinder.ContextMenu.cer'
    $Certificate = [Security.Cryptography.X509Certificates.X509Certificate2]::new($CertificatePath)
    if ($Certificate.Subject -ne 'CN=OfficePDFBinder.ContextMenu') {
        throw 'Unexpected certificate publisher'
    }

    [pscustomobject]@{
        AppDir          = $AppDir
        Support         = $Support
        CertificatePath = $CertificatePath
        Thumbprint      = $Certificate.Thumbprint
    }
}

function Install-MachineCertificate {
    $Info = Get-ShellIntegrationInfo
    if (-not (Test-Path -LiteralPath "$MachineStore\$($Info.Thumbprint)")) {
        Import-Certificate -FilePath $Info.CertificatePath `
            -CertStoreLocation $MachineStore | Out-Null
    }
}

function Uninstall-MachineCertificate {
    $Info = Get-ShellIntegrationInfo

    $RemainingPackages = @(
        Get-AppxPackage -AllUsers -Name $PackageName -ErrorAction Stop
    )
    if ($RemainingPackages.Count -ne 0) {
        return
    }

    $CertificatePath = "$MachineStore\$($Info.Thumbprint)"
    if (Test-Path -LiteralPath $CertificatePath) {
        Remove-Item -LiteralPath $CertificatePath -Force
    }
}

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
    $Info = Get-ShellIntegrationInfo
    $AppDir = $Info.AppDir
    $Support = $Info.Support
    $Thumbprint = $Info.Thumbprint
    $Owner = Get-ItemProperty -LiteralPath $OwnerKey -ErrorAction SilentlyContinue

    if ($Unregister) {
        if ($Owner -and $Owner.InstallationDirectory -ne $AppDir) { return }

        Get-AppxPackage -Name $PackageName | Remove-AppxPackage

        # Cleanup for installations made by older builds that owned a CurrentUser certificate.
        if ($Owner -and $Owner.OwnedCertificate -eq $Thumbprint -and
            @(Get-AppxPackage -Name $PackageName).Count -eq 0 -and
            (Test-Path -LiteralPath "$CurrentUserStore\$Thumbprint")) {
            Remove-Item -LiteralPath "$CurrentUserStore\$Thumbprint" -Force
        }

        if ($Owner) { Remove-Item -LiteralPath $OwnerKey -Recurse -Force }
        Remove-LegacyMenu
        return
    }

    Remove-LegacyMenu
    if ([Environment]::OSVersion.Version.Build -lt 22000) { return }

    New-Item -Path $OwnerKey -Force | Out-Null
    Set-ItemProperty -LiteralPath $OwnerKey -Name InstallationDirectory -Value $AppDir

    Add-AppxPackage -Path (Join-Path $Support 'OfficePDFBinder.ContextMenu.msix') `
        -ExternalLocation $AppDir -ForceUpdateFromAnyVersion
}

if ($MyInvocation.InvocationName -ne '.') {
    $LogFile = Join-Path $env:TEMP 'OfficePDFBinder-shell.log'
    try {
        "$(Get-Date -Format o) Unregister=$Unregister" | Add-Content $LogFile

        $ModeCount = @(
            [bool]$Unregister,
            [bool]$InstallMachineCertificate,
            [bool]$UninstallMachineCertificate
        ).Where({ $_ }).Count
        if ($ModeCount -gt 1) {
            throw 'Only one operation switch may be specified'
        }

        if ($InstallMachineCertificate) {
            Install-MachineCertificate
        }
        elseif ($UninstallMachineCertificate) {
            Uninstall-MachineCertificate
        }
        else {
            Invoke-ShellRegistration
        }

        'SUCCESS' | Add-Content $LogFile
        exit 0
    }
    catch {
        $_ | Out-String | Add-Content $LogFile
        Write-Error $_ -ErrorAction Continue
        exit 1
    }
}
