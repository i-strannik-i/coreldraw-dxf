"""Prepare a versioned release manifest and atomically mirror to shared storage."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import shutil
from corel_export.version import VERSION, INSTALLER_NAME, CHANNEL


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--mirror', type=Path)
    args = parser.parse_args()
    root = Path(__file__).parent / 'corel_export'
    installer = root / (INSTALLER_NAME + '.exe')
    data = installer.read_bytes()
    info = {'product': 'CorelDXF', 'version': VERSION, 'channel': CHANNEL, 'corel_major': 25,
            'file': installer.name, 'size': len(data), 'sha256': hashlib.sha256(data).hexdigest()}
    manifest = root / 'CorelDXF-latest.json'
    manifest.write_text(json.dumps(info, ensure_ascii=False, indent=2), encoding='utf-8')
    if args.mirror:
        args.mirror.mkdir(parents=True, exist_ok=True)
        destination = args.mirror / installer.name
        if destination.exists() and hashlib.sha256(destination.read_bytes()).hexdigest() != info['sha256']:
            raise RuntimeError('Published version differs; increase VERSION instead of overwriting it')
        temp = destination.with_suffix('.upload')
        shutil.copy2(installer, temp)
        if hashlib.sha256(temp.read_bytes()).hexdigest() != info['sha256']:
            raise RuntimeError('Mirror verification failed')
        os.replace(temp, destination)
        pending = args.mirror / (manifest.name + '.upload')
        shutil.copy2(manifest, pending)
        os.replace(pending, args.mirror / manifest.name)
    print(manifest)


if __name__ == '__main__':
    main()
