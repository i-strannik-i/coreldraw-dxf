"""Build the one-file designer installer after building the converter and GMS."""
from pathlib import Path
import subprocess
import sys
from version import INSTALLER_NAME

root = Path(__file__).resolve().parent
subprocess.run([
    sys.executable, '-m', 'PyInstaller', '--noconfirm', '--onefile', '--windowed',
    '--name', 'VectorTools', '--hidden-import', 'win32timezone',
    '--paths', str(root.parent / 'engine'), '--distpath', str(root / 'dist/CorelDXF'),
    '--workpath', str(root / 'build/vectors'), '--specpath', str(root),
    str(root / 'vector_tools.py'),
], check=True)
subprocess.run([
    sys.executable, '-m', 'PyInstaller', '--noconfirm', '--onefile', '--windowed',
    '--name', 'UpdateCorelDXF', '--distpath', str(root / 'dist/CorelDXF'),
    '--workpath', str(root / 'build/updater'), '--specpath', str(root),
    str(root / 'update_manager.py'),
], check=True)
subprocess.run([
    sys.executable, '-m', 'PyInstaller', '--noconfirm', '--onefile', '--windowed',
    '--name', 'HotkeySettings', '--hidden-import', 'win32timezone',
    '--distpath', str(root / 'dist/CorelDXF'), '--workpath', str(root / 'build/hotkeys'),
    '--specpath', str(root), str(root / 'hotkey_settings.py'),
], check=True)
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
