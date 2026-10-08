"""Build the converter and designer installer from this standalone project."""
from pathlib import Path
import subprocess
import sys

root = Path(__file__).resolve().parent
macro = root / 'corel_export'
subprocess.run([
    sys.executable, '-m', 'PyInstaller', '--noconfirm', '--clean',
    '--onedir', '--windowed', '--name', 'CorelDXF',
    '--hidden-import', 'win32timezone', '--paths', str(root / 'engine'),
    '--distpath', str(macro / 'dist'), '--workpath', str(macro / 'build/converter'),
    '--specpath', str(macro), str(macro / 'export_dxf.py'),
], check=True, cwd=root)
subprocess.run([sys.executable, str(macro / 'build_setup.py')], check=True, cwd=root)
