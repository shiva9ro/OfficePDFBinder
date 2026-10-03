"""Require external-location capabilities on both supported architectures."""
from pathlib import Path
import re
import xml.etree.ElementTree as ET
import pytest
from version import APP_VERSION


@pytest.mark.parametrize("architecture", ["x64", "arm64"])
def test_external_location_manifest_declares_required_capabilities(architecture):
    script = (Path(__file__).resolve().parents[1] / "scripts/build_sparse_package.ps1").read_text(
        encoding="utf-8-sig"
    )
    manifest = re.search(r'<\?xml.*?</Package>', script, re.DOTALL).group()
    manifest = manifest.replace('$Subject', 'CN=OfficePDFBinder.ContextMenu')
    manifest = manifest.replace('$Version', f'{APP_VERSION}.1').replace('$Arch', architecture).replace('$Verbs', '')
    root = ET.fromstring(manifest)
    ns = {
        "f": "http://schemas.microsoft.com/appx/manifest/foundation/windows10",
        "uap10": "http://schemas.microsoft.com/appx/manifest/uap/windows10/10",
        "rescap": "http://schemas.microsoft.com/appx/manifest/foundation/windows10/restrictedcapabilities",
    }
    assert root.find('f:Properties/uap10:AllowExternalContent', ns).text == 'true'
    capabilities = {node.attrib['Name'] for node in root.findall('f:Capabilities/rescap:Capability', ns)}
    assert {'runFullTrust', 'unvirtualizedResources'} <= capabilities
    application = root.find('f:Applications/f:Application', ns)
    assert application.attrib[f'{{{ns["uap10"]}}}TrustLevel'] == 'mediumIL'
    assert application.attrib[f'{{{ns["uap10"]}}}RuntimeBehavior'] == 'win32App'
