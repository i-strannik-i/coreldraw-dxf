"""Build the one-file designer installer after building the converter and GMS."""
from pathlib import Path
import subprocess
import shutil
import sys
from version import INSTALLER_NAME

root = Path(__file__).resolve().parent
# Keep the macro's executable paths stable, with private dependency directories.
# Only the installer is onefile; installed tools never unpack at startup.
for name, script, work in (
    ('VectorTools', 'vector_tools.py', 'vectors'),
    ('UpdateCorelDXF', 'update_manager.py', 'updater'),
    ('HotkeySettings', 'hotkey_settings.py', 'hotkeys'),
):
    subprocess.run([
        sys.executable, '-m', 'PyInstaller', '--noconfirm', '--onedir', '--windowed',
        '--name', name, '--contents-directory', '_' + name,
        '--hidden-import', 'win32timezone', '--paths', str(root.parent / 'engine'),
        '--distpath', str(root / 'build/installed-tools'),
        '--workpath', str(root / 'build' / work), '--specpath', str(root),
        str(root / script),
    ], check=True)
    shutil.copytree(root / 'build/installed-tools' / name, root / 'dist/CorelDXF', dirs_exist_ok=True)
for script in ('build_icon.py', 'package_release.py'):
    subprocess.run([sys.executable, str(root / script)], check=True)
subprocess.run([
    sys.executable, '-m', 'PyInstaller', '--noconfirm', '--onefile', '--windowed',
    '--name', INSTALLER_NAME, '--icon', str(root / 'icons/CorelDXF.ico'),
    '--hidden-import', 'win32timezone',
    '--add-data', str(root / 'CorelDXF_2024.zip') + ';.',
    '--distpath', str(root), '--workpath', str(root / 'build/setup'),
    '--specpath', str(root), str(root / 'setup_corel.py'),
], check=True)
