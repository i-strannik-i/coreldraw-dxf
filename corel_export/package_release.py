"""Assemble the standalone designer distribution after PyInstaller/GMS builds."""
from pathlib import Path
import shutil
import zipfile
import importlib.metadata

root = Path(__file__).resolve().parent
package = root / 'package' / 'CorelDXF_2024'
package.mkdir(parents=True, exist_ok=True)
shutil.copytree(root / 'dist' / 'CorelDXF', package / 'runtime', dirs_exist_ok=True)
shutil.copytree(root / 'icons', package / 'icons', dirs_exist_ok=True)
for name in ('SkladCorelDXF.gms', 'README_RU.txt', 'ExportDXF.bas', 'HotkeyBridge.bas', 'ContourCheck.bas'):
    shutil.copy2(root / name, package / name)
(package / 'Install.vbs').unlink(missing_ok=True)
licenses = package / 'licenses'
licenses.mkdir(exist_ok=True)
shutil.copy2(root.parent / 'LICENSE', package / 'LICENSE')
shutil.copy2(root.parent / 'LICENSE', package / 'runtime' / 'LICENSE')
shutil.copy2(root.parent / 'LICENSE', licenses / 'CorelDXF-MIT.txt')
for name in ('numpy', 'ezdxf', 'pywin32', 'typing_extensions', 'fonttools'):
    distribution = importlib.metadata.distribution(name)
    for file in distribution.files or []:
        if 'license' in str(file).lower() and Path(file).suffix.lower() in ('.txt', '.md', ''):
            source = Path(distribution.locate_file(file))
            if source.is_file():
                shutil.copy2(source, licenses / (name + '_' + source.name))
archive = root / 'CorelDXF_2024.zip'
with zipfile.ZipFile(archive, 'w', zipfile.ZIP_DEFLATED) as target:
    for path in package.rglob('*'):
        if path.is_file():
            target.write(path, path.relative_to(package.parent))
print(archive)
