#ifndef MyAppVersion
#define MyAppVersion "1.3.2"
#endif

[Setup]
SourceDir=..
AppName=Office PDF Binder
AppId={{85651C7D-2D19-4AD3-A127-173365C70370}
AppVersion={#MyAppVersion}
AppPublisher=Takeshi Kashiwagi
AppCopyright=Takeshi Kashiwagi

PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog

DefaultDirName={autopf}\Office PDF Binder
DefaultGroupName=Office PDF Binder

OutputDir=Output
OutputBaseFilename=OfficePDFBinder_Setup_{#MyAppVersion}

Compression=lzma2
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64compatible

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

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
Source: "OfficePDFBinder_Main.dist\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "LICENSE.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "NOTICE.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "README.html"; DestDir: "{app}"; Flags: ignoreversion
Source: "README.ja.html"; DestDir: "{app}"; Flags: ignoreversion
Source: "source.zip"; DestDir: "{app}"; Flags: ignoreversion
Source: "app.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "docs\images\*"; DestDir: "{app}\docs\images"; Flags: ignoreversion recursesubdirs createallsubdirs

[UninstallDelete]
Type: filesandordirs; Name: "{userappdata}\OfficePDFBinder"

[Icons]
Name: "{group}\Office PDF Binder"; Filename: "{app}\OfficePDFBinder_Main.exe"
Name: "{autodesktop}\Office PDF Binder"; Filename: "{app}\OfficePDFBinder_Main.exe"; Tasks: desktopicon

[Registry]
; 右クリックメニュー（複数ファイル対応）
; 注意:
;  - Software\Classes\*（すべてのファイルタイプ）では既定アプリへの影響が読みにくいため使用しない
;  - 代わりに SystemFileAssociations\.拡張子\shell を使い、既定アプリは一切変更しない
;
; PDFファイル
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.pdf\shell\OfficePDFBinder"; ValueType: string; ValueName: ""; ValueData: "{cm:OpenWith}"; Flags: uninsdeletekey
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.pdf\shell\OfficePDFBinder\command"; ValueType: string; ValueName: ""; ValueData: """{app}\OfficePDFBinder_Main.exe"" ""%1"""; Flags: uninsdeletekey
; Wordファイル
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.docx\shell\OfficePDFBinder"; ValueType: string; ValueName: ""; ValueData: "{cm:OpenWith}"; Flags: uninsdeletekey
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.docx\shell\OfficePDFBinder\command"; ValueType: string; ValueName: ""; ValueData: """{app}\OfficePDFBinder_Main.exe"" ""%1"""; Flags: uninsdeletekey
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.doc\shell\OfficePDFBinder"; ValueType: string; ValueName: ""; ValueData: "{cm:OpenWith}"; Flags: uninsdeletekey
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.doc\shell\OfficePDFBinder\command"; ValueType: string; ValueName: ""; ValueData: """{app}\OfficePDFBinder_Main.exe"" ""%1"""; Flags: uninsdeletekey
; Excelファイル
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.xlsx\shell\OfficePDFBinder"; ValueType: string; ValueName: ""; ValueData: "{cm:OpenWith}"; Flags: uninsdeletekey
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.xlsx\shell\OfficePDFBinder\command"; ValueType: string; ValueName: ""; ValueData: """{app}\OfficePDFBinder_Main.exe"" ""%1"""; Flags: uninsdeletekey
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.xls\shell\OfficePDFBinder"; ValueType: string; ValueName: ""; ValueData: "{cm:OpenWith}"; Flags: uninsdeletekey
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.xls\shell\OfficePDFBinder\command"; ValueType: string; ValueName: ""; ValueData: """{app}\OfficePDFBinder_Main.exe"" ""%1"""; Flags: uninsdeletekey
; PowerPointファイル
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.pptx\shell\OfficePDFBinder"; ValueType: string; ValueName: ""; ValueData: "{cm:OpenWith}"; Flags: uninsdeletekey
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.pptx\shell\OfficePDFBinder\command"; ValueType: string; ValueName: ""; ValueData: """{app}\OfficePDFBinder_Main.exe"" ""%1"""; Flags: uninsdeletekey
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.ppt\shell\OfficePDFBinder"; ValueType: string; ValueName: ""; ValueData: "{cm:OpenWith}"; Flags: uninsdeletekey
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.ppt\shell\OfficePDFBinder\command"; ValueType: string; ValueName: ""; ValueData: """{app}\OfficePDFBinder_Main.exe"" ""%1"""; Flags: uninsdeletekey

[Code]
procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    SaveStringToFile(
      ExpandConstant('{app}\OfficePDFBinder.language'),
      ActiveLanguage,
      False
    );
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

function InitializeSetup(): Boolean;
var
  InstallerPath: String;
  DrivePath: String;
  DriveType: Integer;
begin
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
