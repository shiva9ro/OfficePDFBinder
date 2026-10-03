"""Validate the external-location manifest's basic Win32 declarations.

Actual registration is a separate integration check. On the tested Windows ARM64
PC, the shipping 1.5.1.1 package registers without unvirtualizedResources; do not
turn that documentation recommendation into an unverified rejection criterion.
"""
from pathlib import Path
import re
import xml.etree.ElementTree as ET


def test_external_location_manifest_declares_required_capabilities():
    script = (Path(__file__).resolve().parents[1] / "scripts/build_sparse_package.ps1").read_text(
        encoding="utf-8-sig"
    )
    manifest = re.search(r'<\?xml.*?</Package>', script, re.DOTALL).group()
    manifest = manifest.replace('$Subject', 'CN=OfficePDFBinder.ContextMenu')
    manifest = manifest.replace('$Version', '1.5.0.1').replace('$Arch', 'x64').replace('$Verbs', '')
    root = ET.fromstring(manifest)
    ns = {
        "f": "http://schemas.microsoft.com/appx/manifest/foundation/windows10",
        "uap10": "http://schemas.microsoft.com/appx/manifest/uap/windows10/10",
        "rescap": "http://schemas.microsoft.com/appx/manifest/foundation/windows10/restrictedcapabilities",
    }
    assert root.find('f:Properties/uap10:AllowExternalContent', ns).text == 'true'
    capabilities = {node.attrib['Name'] for node in root.findall('f:Capabilities/rescap:Capability', ns)}
    assert 'runFullTrust' in capabilities
    application = root.find('f:Applications/f:Application', ns)
    assert application.attrib[f'{{{ns["uap10"]}}}TrustLevel'] == 'mediumIL'
    assert application.attrib[f'{{{ns["uap10"]}}}RuntimeBehavior'] == 'win32App'
