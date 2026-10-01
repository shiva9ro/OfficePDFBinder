"""Exercise user-only installation without touching real packages, trust or registry."""
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
        assert ("trust:Cert:\\CurrentUser\\TrustedPeople" in calls) == (scenario == "user-install")
    elif scenario == "other-uninstall":
        assert calls == []
    else:
        assert "unregister" in calls
        removed_trust = [c for c in calls if c.startswith("remove:Cert:")]
        assert bool(removed_trust) == (scenario == "user-uninstall")
