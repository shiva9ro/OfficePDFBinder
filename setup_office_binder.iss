#ifndef MyAppVersion
#define MyAppVersion "1.0.0"
#endif

[Setup]
AppName=Office PDF Binder
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
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
Source: "OfficePDFBinder_Main.dist\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "LICENSE.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "NOTICE.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "README.html"; DestDir: "{app}"; Flags: ignoreversion
Source: "source.zip"; DestDir: "{app}"; Flags: ignoreversion

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
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.pdf\shell\OfficePDFBinder"; ValueType: string; ValueName: ""; ValueData: "Office PDF Binder で開く"; Flags: uninsdeletekey
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.pdf\shell\OfficePDFBinder\command"; ValueType: string; ValueName: ""; ValueData: """{app}\OfficePDFBinder_Main.exe"" ""%1"""; Flags: uninsdeletekey
; Wordファイル
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.docx\shell\OfficePDFBinder"; ValueType: string; ValueName: ""; ValueData: "Office PDF Binder で開く"; Flags: uninsdeletekey
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.docx\shell\OfficePDFBinder\command"; ValueType: string; ValueName: ""; ValueData: """{app}\OfficePDFBinder_Main.exe"" ""%1"""; Flags: uninsdeletekey
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.doc\shell\OfficePDFBinder"; ValueType: string; ValueName: ""; ValueData: "Office PDF Binder で開く"; Flags: uninsdeletekey
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.doc\shell\OfficePDFBinder\command"; ValueType: string; ValueName: ""; ValueData: """{app}\OfficePDFBinder_Main.exe"" ""%1"""; Flags: uninsdeletekey
; Excelファイル
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.xlsx\shell\OfficePDFBinder"; ValueType: string; ValueName: ""; ValueData: "Office PDF Binder で開く"; Flags: uninsdeletekey
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.xlsx\shell\OfficePDFBinder\command"; ValueType: string; ValueName: ""; ValueData: """{app}\OfficePDFBinder_Main.exe"" ""%1"""; Flags: uninsdeletekey
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.xls\shell\OfficePDFBinder"; ValueType: string; ValueName: ""; ValueData: "Office PDF Binder で開く"; Flags: uninsdeletekey
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.xls\shell\OfficePDFBinder\command"; ValueType: string; ValueName: ""; ValueData: """{app}\OfficePDFBinder_Main.exe"" ""%1"""; Flags: uninsdeletekey
; PowerPointファイル
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.pptx\shell\OfficePDFBinder"; ValueType: string; ValueName: ""; ValueData: "Office PDF Binder で開く"; Flags: uninsdeletekey
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.pptx\shell\OfficePDFBinder\command"; ValueType: string; ValueName: ""; ValueData: """{app}\OfficePDFBinder_Main.exe"" ""%1"""; Flags: uninsdeletekey
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.ppt\shell\OfficePDFBinder"; ValueType: string; ValueName: ""; ValueData: "Office PDF Binder で開く"; Flags: uninsdeletekey
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.ppt\shell\OfficePDFBinder\command"; ValueType: string; ValueName: ""; ValueData: """{app}\OfficePDFBinder_Main.exe"" ""%1"""; Flags: uninsdeletekey

[Code]
function GetDriveType(lpRootPathName: String): Integer;
  external 'GetDriveTypeA@kernel32.dll stdcall';

const
  DRIVE_REMOTE = 4;

function InitializeSetup(): Boolean;
var
  InstallerPath: String;
  DrivePath: String;
  DriveType: Integer;
begin
  InstallerPath := ExpandConstant('{src}');
  
  // ネットワークパスかどうかをチェック
  // UNCパス（\\server\share）の場合
  if (Pos('\\', InstallerPath) = 1) then
  begin
    MsgBox('ネットワーク経由でのインストールは許可されていません。' + #13#10 +
           'インストーラーをローカルドライブにコピーしてから実行してください。',
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
      MsgBox('ネットワーク経由でのインストールは許可されていません。' + #13#10 +
             'インストーラーをローカルドライブにコピーしてから実行してください。',
             mbError, MB_OK);
      Result := False;
      Exit;
    end;
  end;
  
  Result := True;
end;

[Run]
Filename: "{app}\OfficePDFBinder_Main.exe"; Description: "{cm:LaunchProgram,Office PDF Binder}"; Flags: nowait postinstall skipifsilent
