"""Execute the real uninstall decision code with mocked external operations.

The generated test EXE never installs anything, invokes PowerShell or requests
elevation. InitializeSetup runs the cases and returns False before setup starts.
"""
from pathlib import Path
import re
import subprocess

import pytest

ROOT = Path(__file__).resolve().parents[1]


def test_certificate_failure_does_not_block_uninstall(tmp_path):
    compiler = Path(r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe")
    if not compiler.exists():
        pytest.skip("Inno Setup 6 is required for executable Pascal policy tests")
    source = (ROOT / "packaging/setup_office_binder.iss").read_text(encoding="utf-8-sig")
    policy = source.split("function InitializeUninstall(): Boolean;", 1)[1]
    policy = "function CheckUninstall(): Boolean;" + policy.split("procedure CurStepChanged", 1)[0]
    for name in ("FileExists", "ExpandConstant", "Exec", "ShellExec", "MsgBox", "SuppressibleMsgBox"):
        policy = re.sub(rf"\b{name}\b", "Mock" + name, policy)
    script = r"""
[Setup]
AppName=Uninstall policy test
AppVersion=1
DefaultDirName={tmp}\UnusedPolicyTest
PrivilegesRequired=lowest
Uninstallable=no
CreateAppDir=no
OutputBaseFilename=policy-test
[Code]
var
  Marker, PackageStarted, CertificateStarted: Boolean;
  PackageCode, CertificateCode, PackageCalls, CertificateCalls, Warnings, Errors: Integer;
function MockFileExists(Path: String): Boolean;
begin Result := Marker; end;
function MockExpandConstant(Value: String): String;
begin Result := Value; end;
function MockExec(FileName, Params, Dir: String; Show: Integer; Wait: TExecWait; var Code: Integer): Boolean;
begin
  PackageCalls := PackageCalls + 1;
  Code := PackageCode;
  Result := PackageStarted;
end;
function MockShellExec(Verb, FileName, Params, Dir: String; Show: Integer; Wait: TExecWait; var Code: Integer): Boolean;
begin
  CertificateCalls := CertificateCalls + 1;
  Code := CertificateCode;
  Result := CertificateStarted;
end;
function MockMsgBox(Text: String; Typ: TMsgBoxType; Buttons: Integer): Integer;
begin Errors := Errors + 1; Result := IDOK; end;
function MockSuppressibleMsgBox(Text: String; Typ: TMsgBoxType; Buttons, Default: Integer): Integer;
begin Warnings := Warnings + 1; Result := IDOK; end;
""" + policy + r"""
procedure RunCase(Name: String; HasMarker, PStarted: Boolean; PCode: Integer;
  CStarted: Boolean; CCode: Integer; Expected: Boolean;
  ExpectedPackageCalls, ExpectedCertificateCalls, ExpectedWarnings, ExpectedErrors: Integer);
begin
  Marker := HasMarker; PackageStarted := PStarted; PackageCode := PCode;
  CertificateStarted := CStarted; CertificateCode := CCode;
  PackageCalls := 0; CertificateCalls := 0; Warnings := 0; Errors := 0;
  if CheckUninstall() <> Expected then RaiseException(Name + ': wrong uninstall decision');
  if PackageCalls <> ExpectedPackageCalls then RaiseException(Name + ': wrong package calls');
  if CertificateCalls <> ExpectedCertificateCalls then RaiseException(Name + ': wrong certificate calls');
  if Warnings <> ExpectedWarnings then RaiseException(Name + ': wrong warnings');
  if Errors <> ExpectedErrors then RaiseException(Name + ': wrong errors');
end;
function InitializeSetup(): Boolean;
var Report: String;
begin
  Result := False;
  try
    RunCase('no registration', False, True, 0, True, 0, True, 0, 0, 0, 0);
    RunCase('normal cleanup', True, True, 0, True, 0, True, 1, 1, 0, 0);
    RunCase('package launch failed', True, False, 2, True, 0, False, 1, 0, 0, 1);
    RunCase('package removal failed', True, True, 1, True, 0, False, 1, 0, 0, 1);
    RunCase('certificate UAC denied', True, True, 0, False, 1223, True, 1, 1, 1, 0);
    RunCase('certificate launch failed', True, True, 0, False, 2, True, 1, 1, 1, 0);
    RunCase('certificate query or deletion failed', True, True, 0, True, 1, True, 1, 1, 1, 0);
    Report := 'PASS: 7 uninstall decision cases';
  except
    Report := 'FAIL: ' + GetExceptionMessage;
  end;
  SaveStringToFile(ExpandConstant('{param:Report}'), Report, False);
end;
"""
    test_source = tmp_path / "policy.iss"
    test_source.write_text(script, encoding="utf-8-sig")
    compiled = subprocess.run(
        [str(compiler), "/Q", f"/O{tmp_path}", str(test_source)],
        capture_output=True, text=True, timeout=60,
    )
    assert compiled.returncode == 0, compiled.stdout + compiled.stderr
    report = tmp_path / "result.txt"
    subprocess.run(
        [str(tmp_path / "policy-test.exe"), "/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART", f"/Report={report}"],
        capture_output=True, timeout=30,
    )
    assert report.read_text(encoding="utf-8-sig") == "PASS: 7 uninstall decision cases"
