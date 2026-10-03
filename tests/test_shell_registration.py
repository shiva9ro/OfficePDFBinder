"""Exercise package and certificate scopes without changing Windows registration."""
import json
from pathlib import Path
import shutil
import subprocess

import pytest

ROOT = Path(__file__).resolve().parents[1]


@pytest.mark.parametrize("scenario", ["user-install", "pretrusted-install", "user-uninstall",
    "pretrusted-uninstall", "other-uninstall"])
def test_registration_scope_and_ownership(tmp_path, scenario):
    support = tmp_path / "shell-integration"
    support.mkdir()
    cert = ROOT / "native/shell_bridge/sparse.build/OfficePDFBinder.ContextMenu.cer"
    if not cert.exists():
        pytest.skip("Build sparse packages first")
    shutil.copy2(cert, support)
    result = subprocess.run(["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
        "-File", str(ROOT / "tests/shell_registration_harness.ps1"),
        "-Script", str(ROOT / "native/shell_bridge/register_sparse_package.ps1"),
        "-AppDir", str(tmp_path), "-Scenario", scenario], capture_output=True, text=True)
    assert result.returncode == 0, result.stderr
    calls = json.loads(result.stdout)
    assert not any("LocalMachine" in c or "HKLM" in c for c in calls)
    if scenario.endswith("-install"):
        assert "add:OfficePDFBinder.ContextMenu.msix" in calls
        # Trust is installed separately by the elevated certificate operation.
        assert not any(c.startswith("trust:") for c in calls)
    elif scenario == "other-uninstall":
        assert calls == []
    else:
        assert "unregister" in calls
        removed_trust = [c for c in calls if c.startswith("remove:Cert:")]
        assert bool(removed_trust) == (scenario == "user-uninstall")


@pytest.mark.parametrize("scenario", [
    "machine-install", "machine-pretrusted-install", "machine-uninstall",
    "machine-shared-uninstall", "machine-absent-uninstall", "machine-query-error",
    "machine-delete-error",
])
def test_machine_certificate_lifecycle(tmp_path, scenario):
    result = subprocess.run([
        "powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
        str(ROOT / "tests/machine_certificate_harness.ps1"),
        "-Script", str(ROOT / "native/shell_bridge/register_sparse_package.ps1"),
        "-AppDir", str(tmp_path), "-Scenario", scenario,
    ], capture_output=True, text=True, timeout=30)
    assert result.returncode == 0, result.stderr
    outcome = json.loads(result.stdout)
    calls = outcome["calls"]
    assert not any("CurrentUser" in c or c == "unregister" for c in calls)
    if scenario == "machine-query-error":
        assert "remove:Cert:\\LocalMachine\\TrustedPeople\\TEST" not in calls, (
            "A failed all-users query must not authorize certificate deletion"
        )
        assert outcome["error"], "The query error must reach the caller"
    elif scenario == "machine-delete-error":
        assert "query:all-users" in calls
        assert "remove:Cert:\\LocalMachine\\TrustedPeople\\TEST" in calls
        assert outcome["error"], "The deletion error must reach the installer"
    else:
        assert not outcome["error"]
        assert ("trust:Cert:\\LocalMachine\\TrustedPeople" in calls) == (
            scenario == "machine-install"
        )
        assert ("remove:Cert:\\LocalMachine\\TrustedPeople\\TEST" in calls) == (
            scenario == "machine-uninstall"
        )
        if "uninstall" in scenario:
            assert "query:all-users" in calls
