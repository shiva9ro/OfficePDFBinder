#ifndef MyAppVersion
#define MyAppVersion "1.5.4"
#endif

[Setup]
SourceDir=..
AppName=Office PDF Binder
AppId={{85651C7D-2D19-4AD3-A127-173365C70370}
AppVersion={#MyAppVersion}
AppPublisher=Takeshi Kashiwagi
AppCopyright=Takeshi Kashiwagi

PrivilegesRequired=lowest

DefaultDirName={localappdata}\Programs\Office PDF Binder
DefaultGroupName=Office PDF Binder

OutputDir=Output
OutputBaseFilename=OfficePDFBinder_Setup_{#MyAppVersion}

Compression=lzma2
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64compatible
ChangesAssociations=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"

[CustomMessages]
english.OpenWith=Open with Office PDF Binder
japanese.OpenWith=Office PDF Binder で開く
english.LegacyInstallPresent=An older version of Office PDF Binder is installed.
japanese.LegacyInstallPresent=旧バージョンの Office PDF Binder がインストールされています。
english.UninstallBeforeUpgrade=Uninstall the older version, and then run this installer again.
japanese.UninstallBeforeUpgrade=先に旧バージョンをアンインストールしてから、このインストーラーを再実行してください。
english.InstallationAborted=Installation will be canceled.
japanese.InstallationAborted=インストールを中止します。
english.NetworkInstallDenied=Installation from a network location is not allowed.
japanese.NetworkInstallDenied=ネットワーク経由でのインストールは許可されていません。
english.CopyInstallerLocally=Copy the installer to a local drive, and then run it again.
japanese.CopyInstallerLocally=インストーラーをローカルドライブにコピーしてから実行してください。
english.ModernMenu=Show in the Windows 11 context menu
japanese.ModernMenu=Windows 11の右クリックメニューに表示する
english.ModernMenuFailed=Windows 11 context menu registration failed. See shell-integration in the installation folder to retry.
japanese.ModernMenuFailed=Windows 11のメニュー登録に失敗しました。再登録の手順はインストール先のshell-integrationをご確認ください。

english.MenuUninstallFailed=Context menu removal failed. Uninstallation was stopped to preserve its files. Close Office PDF Binder and retry. See %TEMP%\OfficePDFBinder-shell.log.
japanese.MenuUninstallFailed=右クリックメニューの解除に失敗したため、アンインストールを中止しました。Office PDF Binderを閉じて再実行してください。ログ: %TEMP%\OfficePDFBinder-shell.log
english.CertificateCleanupFailed=The context-menu signing certificate could not be removed. Uninstallation of the application will continue. The certificate may remain on this PC.
japanese.CertificateCleanupFailed=右クリックメニュー用の署名証明書を削除できませんでした。アプリ本体のアンインストールは続行します。証明書がPCに残る場合があります。

english.MachineInstallPresent=An all-users installation exists. Uninstall it first, then run this installer normally. This version installs for the current user only.
japanese.MachineInstallPresent=全ユーザー向けのOffice PDF Binderが既にインストールされています。先にその版をアンインストールし、このインストーラーを通常起動してください。今後は現在のユーザー専用でインストールします。

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"
Name: "modernmenu"; Description: "{cm:ModernMenu}"; MinVersion: 10.0.22000

[Files]
Source: "OfficePDFBinder_Main.dist\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "native\shell_bridge\out.build\OfficePDFBinder_Shell.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "native\shell_bridge\sparse.build\x64\OfficePDFBinder_Explorer.dll"; DestDir: "{app}"; Flags: ignoreversion; Check: not IsArm64
Source: "native\shell_bridge\sparse.build\arm64\OfficePDFBinder_Explorer.dll"; DestDir: "{app}"; Flags: ignoreversion; Check: IsArm64
Source: "native\shell_bridge\sparse.build\x64\OfficePDFBinder.ContextMenu.msix"; DestDir: "{app}\shell-integration"; Flags: ignoreversion; Check: not IsArm64
Source: "native\shell_bridge\sparse.build\arm64\OfficePDFBinder.ContextMenu.msix"; DestDir: "{app}\shell-integration"; Flags: ignoreversion; Check: IsArm64
Source: "native\shell_bridge\sparse.build\OfficePDFBinder.ContextMenu.cer"; DestDir: "{app}\shell-integration"; Flags: ignoreversion
Source: "native\shell_bridge\register_sparse_package.ps1"; DestDir: "{app}\shell-integration"; Flags: ignoreversion
Source: "native\shell_bridge\MODERN_MENU.md"; DestDir: "{app}\shell-integration"; Flags: ignoreversion
Source: "LICENSE.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "NOTICE.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "README.html"; DestDir: "{app}"; Flags: ignoreversion
Source: "README.ja.html"; DestDir: "{app}"; Flags: ignoreversion
Source: "source.zip"; DestDir: "{app}"; Flags: ignoreversion
Source: "app.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "docs\images\*"; DestDir: "{app}\docs\images"; Flags: ignoreversion recursesubdirs createallsubdirs

[UninstallDelete]
Type: filesandordirs; Name: "{userappdata}\OfficePDFBinder"
Type: files; Name: "{app}\shell-integration\registered.txt"
Type: files; Name: "{app}\OfficePDFBinder.language"

[Icons]
Name: "{group}\Office PDF Binder"; Filename: "{app}\OfficePDFBinder_Main.exe"
Name: "{autodesktop}\Office PDF Binder"; Filename: "{app}\OfficePDFBinder_Main.exe"; Tasks: desktopicon

[Registry]
; Remove only this application's legacy menu/COM keys, on Windows 10 as well.
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.pdf\shell\OfficePDFBinder"; Flags: deletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.doc\shell\OfficePDFBinder"; Flags: deletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.docx\shell\OfficePDFBinder"; Flags: deletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.docm\shell\OfficePDFBinder"; Flags: deletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.xls\shell\OfficePDFBinder"; Flags: deletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.xlsx\shell\OfficePDFBinder"; Flags: deletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.xlsm\shell\OfficePDFBinder"; Flags: deletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.ppt\shell\OfficePDFBinder"; Flags: deletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.pptx\shell\OfficePDFBinder"; Flags: deletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.pptm\shell\OfficePDFBinder"; Flags: deletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.png\shell\OfficePDFBinder"; Flags: deletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.jpg\shell\OfficePDFBinder"; Flags: deletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.jpeg\shell\OfficePDFBinder"; Flags: deletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.bmp\shell\OfficePDFBinder"; Flags: deletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.webp\shell\OfficePDFBinder"; Flags: deletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.tif\shell\OfficePDFBinder"; Flags: deletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.tiff\shell\OfficePDFBinder"; Flags: deletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.heic\shell\OfficePDFBinder"; Flags: deletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.heif\shell\OfficePDFBinder"; Flags: deletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.hif\shell\OfficePDFBinder"; Flags: deletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.svg\shell\OfficePDFBinder"; Flags: deletekey
Root: HKCU; Subkey: "Software\Classes\CLSID\{{F9869918-6F6D-4FC9-A8DF-70F43498A802}"; Flags: deletekey
Root: HKCU; Subkey: "Software\Classes\AppID\{{F9869918-6F6D-4FC9-A8DF-70F43498A802}"; Flags: deletekey

[Code]
function InitializeUninstall(): Boolean;
var
  ResultCode: Integer;
  Arguments: String;
  CertificateCleanupSucceeded: Boolean;
begin
  Result := True;
  if not FileExists(ExpandConstant('{app}\shell-integration\registered.txt')) then Exit;
  Arguments := '-NoProfile -ExecutionPolicy Bypass -File "' +
    ExpandConstant('{app}\shell-integration\register_sparse_package.ps1') +
    '" -ApplicationDirectory "' + ExpandConstant('{app}') + '" -Unregister';
  Result := Exec(ExpandConstant('{sys}\WindowsPowerShell\v1.0\powershell.exe'),
    Arguments, '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  if Result then Result := ResultCode = 0;
  if not Result then
  begin
    MsgBox(ExpandConstant('{cm:MenuUninstallFailed}'), mbError, MB_OK);
    Exit;
  end;

  { Certificate cleanup is best effort; it must not block application removal. }
  Arguments := '-NoProfile -ExecutionPolicy Bypass -File "' +
    ExpandConstant('{app}\shell-integration\register_sparse_package.ps1') +
    '" -ApplicationDirectory "' + ExpandConstant('{app}') + '" -UninstallMachineCertificate';
  CertificateCleanupSucceeded := ShellExec('runas',
    ExpandConstant('{sys}\WindowsPowerShell\v1.0\powershell.exe'),
    Arguments, '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  if CertificateCleanupSucceeded then
    CertificateCleanupSucceeded := ResultCode = 0;
  if not CertificateCleanupSucceeded then
  begin
    Log(Format('Certificate cleanup did not complete (code %d); continuing uninstall.', [ResultCode]));
    SuppressibleMsgBox(ExpandConstant('{cm:CertificateCleanupFailed}'), mbInformation, MB_OK, IDOK);
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  ResultCode: Integer;
  Arguments: String;
  Started: Boolean;
begin
  if CurStep = ssPostInstall then
  begin
    SaveStringToFile(
      ExpandConstant('{app}\OfficePDFBinder.language'),
      ActiveLanguage,
      False
    );
    if WizardIsTaskSelected('modernmenu') then
    begin
      SaveStringToFile(ExpandConstant('{app}\shell-integration\registered.txt'), 'attempted', False);

      Arguments := '-NoProfile -ExecutionPolicy Bypass -File "' +
        ExpandConstant('{app}\shell-integration\register_sparse_package.ps1') +
        '" -ApplicationDirectory "' + ExpandConstant('{app}') +
        '" -InstallMachineCertificate';

      Started := ShellExec(
        'runas',
        ExpandConstant('{sys}\WindowsPowerShell\v1.0\powershell.exe'),
        Arguments,
        '',
        SW_HIDE,
        ewWaitUntilTerminated,
        ResultCode
      );

      if Started and (ResultCode = 0) then
      begin
        Arguments := '-NoProfile -ExecutionPolicy Bypass -File "' +
          ExpandConstant('{app}\shell-integration\register_sparse_package.ps1') +
          '" -ApplicationDirectory "' + ExpandConstant('{app}') + '"';

        Started := ExecAsOriginalUser(
          ExpandConstant('{sys}\WindowsPowerShell\v1.0\powershell.exe'),
          Arguments,
          '',
          SW_HIDE,
          ewWaitUntilTerminated,
          ResultCode
        );
      end;

      if Started and (ResultCode = 0) then
        SaveStringToFile(ExpandConstant('{app}\shell-integration\registered.txt'), 'registered', False)
      else
        MsgBox(ExpandConstant('{cm:ModernMenuFailed}'), mbError, MB_OK);
    end;
  end;
end;

function GetDriveType(lpRootPathName: String): Integer;
  external 'GetDriveTypeA@kernel32.dll stdcall';

const
  DRIVE_REMOTE = 4;

function IsLegacyInstallPresent(): Boolean;
var
  DisplayVersion: String;
begin
  Result := False;

  if RegQueryStringValue(HKCU,
    'Software\Microsoft\Windows\CurrentVersion\Uninstall\Office PDF Binder_is1',
    'DisplayVersion', DisplayVersion) then
  begin
    Result := True;
    Exit;
  end;

  if RegQueryStringValue(HKLM,
    'Software\Microsoft\Windows\CurrentVersion\Uninstall\Office PDF Binder_is1',
    'DisplayVersion', DisplayVersion) then
  begin
    Result := True;
    Exit;
  end;
end;

function IsMachineInstallPresent(): Boolean;
var
  Key: String;
begin
  Key := 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{85651C7D-2D19-4AD3-A127-173365C70370}_is1';
  Result := RegKeyExists(HKLM32, Key);
  if IsWin64 then Result := Result or RegKeyExists(HKLM64, Key);
end;

function InitializeSetup(): Boolean;
var
  InstallerPath: String;
  DrivePath: String;
  DriveType: Integer;
begin
  if IsMachineInstallPresent() then
  begin
    MsgBox(ExpandConstant('{cm:MachineInstallPresent}'), mbError, MB_OK);
    Result := False;
    Exit;
  end;
  if IsLegacyInstallPresent() then
  begin
    MsgBox(ExpandConstant('{cm:LegacyInstallPresent}') + #13#10 + #13#10 +
           ExpandConstant('{cm:UninstallBeforeUpgrade}') + #13#10 + #13#10 +
           ExpandConstant('{cm:InstallationAborted}'),
           mbError, MB_OK);
    Result := False;
    Exit;
  end;

  InstallerPath := ExpandConstant('{src}');
  
  // ネットワークパスかどうかをチェック
  // UNCパス（\\server\share）の場合
  if (Pos('\\', InstallerPath) = 1) then
  begin
    MsgBox(ExpandConstant('{cm:NetworkInstallDenied}') + #13#10 +
           ExpandConstant('{cm:CopyInstallerLocally}'),
           mbError, MB_OK);
    Result := False;
    Exit;
  end;
  
  // ドライブタイプをチェック（ネットワークドライブの場合）
  if Length(InstallerPath) >= 2 then
  begin
    DrivePath := Copy(InstallerPath, 1, 2) + '\';
    DriveType := GetDriveType(DrivePath);
    if DriveType = DRIVE_REMOTE then
    begin
      MsgBox(ExpandConstant('{cm:NetworkInstallDenied}') + #13#10 +
             ExpandConstant('{cm:CopyInstallerLocally}'),
             mbError, MB_OK);
      Result := False;
      Exit;
    end;
  end;
  
  Result := True;
end;

[Run]
Filename: "{app}\OfficePDFBinder_Main.exe"; Description: "{cm:LaunchProgram,Office PDF Binder}"; Flags: nowait postinstall skipifsilent
