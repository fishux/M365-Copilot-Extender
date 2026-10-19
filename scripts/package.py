"""Produce a Windows-friendly ZIP with the unpacked extension in its root."""
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import json

root = Path(__file__).resolve().parent.parent
extension = root / 'extension'
version = json.loads((extension / 'manifest.json').read_text())['version']
output = root / 'outputs' / f'M365-Copilot-Extender-v{version}.zip'
output.parent.mkdir(exist_ok=True)
with ZipFile(output, 'w', ZIP_DEFLATED) as archive:
    for file in sorted(extension.rglob('*')):
        if file.is_file():
            archive.write(file, file.relative_to(extension))
print(output)
