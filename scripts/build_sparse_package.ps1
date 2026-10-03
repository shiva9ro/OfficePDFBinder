param(
    [string]$CertificateThumbprint,
    [ValidateSet('x64', 'arm64')][string[]]$Architectures = @('x64', 'arm64')
)
$ErrorActionPreference = 'Stop'
$Root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$Output = Join-Path $Root 'native\shell_bridge\sparse.build'
New-Item -ItemType Directory -Force $Output | Out-Null
$SdkRoot = Join-Path ${env:ProgramFiles(x86)} 'Windows Kits\10\bin'
$Sdk = Get-ChildItem $SdkRoot -Directory | Where-Object {
    Test-Path (Join-Path $_.FullName 'x64\makeappx.exe')
} | Sort-Object Name -Descending | Select-Object -First 1
if (-not $Sdk) { throw 'Windows SDK の MakeAppx と SignTool が必要です。' }
$MakeAppx = Join-Path $Sdk.FullName 'x64\makeappx.exe'
$SignTool = Join-Path $Sdk.FullName 'x64\signtool.exe'
$Subject = 'CN=OfficePDFBinder.ContextMenu'
if ($CertificateThumbprint) {
    $Cert = Get-Item "Cert:\CurrentUser\My\$CertificateThumbprint"
} else {
    $Cert = Get-ChildItem Cert:\CurrentUser\My | Where-Object {
        $_.Subject -eq $Subject -and $_.HasPrivateKey -and $_.NotAfter -gt (Get-Date).AddDays(30)
    } | Sort-Object NotAfter -Descending | Select-Object -First 1
    if (-not $Cert) {
        $Cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject $Subject `
            -CertStoreLocation Cert:\CurrentUser\My -KeyAlgorithm RSA -KeyLength 2048 `
            -KeyExportPolicy NonExportable -HashAlgorithm SHA256 -NotAfter (Get-Date).AddYears(3)
    }
}
if ($Cert.Subject -ne $Subject -or -not $Cert.HasPrivateKey -or $Cert.NotAfter -le (Get-Date)) {
    throw '署名証明書のSubject・秘密鍵・有効期限を確認してください。'
}
Export-Certificate -Cert $Cert -FilePath (Join-Path $Output 'OfficePDFBinder.ContextMenu.cer') -Force | Out-Null
$VersionText = Get-Content (Join-Path $Root 'version.py') -Raw
if ($VersionText -notmatch 'APP_VERSION\s*=\s*"(\d+\.\d+\.\d+)"') { throw 'Version format error' }
$Version = "$($Matches[1]).1"
# One registration for a mixed selection; the command filters supported extensions.
$Verbs = '<desktop5:ItemType Type="*"><desktop5:Verb Id="OfficePDFBinderOpen" Clsid="F9869918-6F6D-4FC9-A8DF-70F43498A804" /></desktop5:ItemType>'
Add-Type -AssemblyName System.Drawing
foreach ($Arch in $Architectures) {
    $ArchOutput = Join-Path $Output $Arch
    & (Join-Path $PSScriptRoot 'build_shell_bridge.ps1') -ExplorerCommand -Architecture $Arch -OutputDirectory $ArchOutput
    $Stage = Join-Path $ArchOutput 'stage'
    New-Item -ItemType Directory -Force (Join-Path $Stage 'Assets') | Out-Null
    $Icon = [Drawing.Icon]::new((Join-Path $Root 'app.ico'))
    $Bitmap = $Icon.ToBitmap()
    try {
        foreach ($Size in @(44, 50, 150)) {
            $Scaled = [Drawing.Bitmap]::new($Bitmap, [Drawing.Size]::new($Size, $Size))
            try { $Scaled.Save((Join-Path $Stage "Assets\Logo$Size.png"), [Drawing.Imaging.ImageFormat]::Png) }
            finally { $Scaled.Dispose() }
        }
    } finally { $Bitmap.Dispose(); $Icon.Dispose() }
    @"
<?xml version="1.0" encoding="utf-8"?>
<Package xmlns="http://schemas.microsoft.com/appx/manifest/foundation/windows10"
 xmlns:uap="http://schemas.microsoft.com/appx/manifest/uap/windows10"
 xmlns:uap10="http://schemas.microsoft.com/appx/manifest/uap/windows10/10"
 xmlns:rescap="http://schemas.microsoft.com/appx/manifest/foundation/windows10/restrictedcapabilities"
 xmlns:com="http://schemas.microsoft.com/appx/manifest/com/windows10"
 xmlns:desktop4="http://schemas.microsoft.com/appx/manifest/desktop/windows10/4"
 xmlns:desktop5="http://schemas.microsoft.com/appx/manifest/desktop/windows10/5"
 IgnorableNamespaces="uap uap10 rescap com desktop4 desktop5">
 <Identity Name="OfficePDFBinder.ContextMenu" Publisher="$Subject" Version="$Version" ProcessorArchitecture="$Arch" />
 <Properties><DisplayName>Office PDF Binder</DisplayName><PublisherDisplayName>Office PDF Binder</PublisherDisplayName><Logo>Assets\Logo50.png</Logo><uap10:AllowExternalContent>true</uap10:AllowExternalContent></Properties>
 <Resources><Resource Language="ja-jp" /><Resource Language="en-us" /></Resources>
 <Dependencies><TargetDeviceFamily Name="Windows.Desktop" MinVersion="10.0.22000.0" MaxVersionTested="10.0.26100.0" /></Dependencies>
 <Applications><Application Id="OfficePDFBinder" Executable="OfficePDFBinder_Main.exe" uap10:TrustLevel="mediumIL" uap10:RuntimeBehavior="win32App">
 <uap:VisualElements AppListEntry="none" DisplayName="Office PDF Binder" Description="Office PDF Binder" BackgroundColor="transparent" Square150x150Logo="Assets\Logo150.png" Square44x44Logo="Assets\Logo44.png" />
 <Extensions><com:Extension Category="windows.comServer"><com:ComServer><com:ExeServer Executable="OfficePDFBinder_Shell.exe" DisplayName="Office PDF Binder Shell"><com:Class Id="F9869918-6F6D-4FC9-A8DF-70F43498A802" /></com:ExeServer><com:SurrogateServer DisplayName="Office PDF Binder"><com:Class Id="F9869918-6F6D-4FC9-A8DF-70F43498A804" Path="OfficePDFBinder_Explorer.dll" ThreadingModel="STA" /></com:SurrogateServer></com:ComServer></com:Extension>
 <desktop4:Extension Category="windows.fileExplorerContextMenus"><desktop4:FileExplorerContextMenus>$Verbs</desktop4:FileExplorerContextMenus></desktop4:Extension></Extensions>
 </Application></Applications><Capabilities><rescap:Capability Name="runFullTrust" /><rescap:Capability Name="unvirtualizedResources" /></Capabilities>
</Package>
"@ | Set-Content (Join-Path $Stage 'AppxManifest.xml') -Encoding utf8
    $Package = Join-Path $ArchOutput 'OfficePDFBinder.ContextMenu.msix'
    # External EXE/DLL references are intentionally outside the sparse package.
    & $MakeAppx pack /d $Stage /p $Package /o /nv
    if ($LASTEXITCODE -ne 0) { throw "MakeAppx failed ($Arch)" }
    & $SignTool sign /fd SHA256 /sha1 $Cert.Thumbprint /s My $Package
    if ($LASTEXITCODE -ne 0) { throw "SignTool failed ($Arch)" }

}
# Only the public certificate is distributed. The private key stays in the signing user's store.
Write-Host "Sparse packages: $Output"
