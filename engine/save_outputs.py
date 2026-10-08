"""Publish checked import files beside the source, retaining replaced versions."""
from datetime import datetime
import hashlib
import json
import os
from pathlib import Path
import shutil
import tempfile
import uuid


def publish_file(source, target):
    source, target = Path(source), Path(target)
    if source.resolve() == target.resolve():
        raise ValueError('Output cannot replace itself')
    if target.is_symlink():
        raise ValueError('Refusing to overwrite a symbolic link: ' + str(target))
    if not source.is_file() or source.stat().st_size == 0:
        raise ValueError('Saved output is missing or empty: ' + str(source))
    fd, staged = tempfile.mkstemp(prefix='.cdr-import-', suffix='.tmp', dir=target.parent)
    os.close(fd)
    backup = None
    try:
        shutil.copyfile(source, staged)
        if target.exists():
            backup_dir = target.parent / '_cdr_import_backups'
            backup_dir.mkdir(exist_ok=True)
            stamp = datetime.now().strftime('%Y%m%d_%H%M%S_%f')
            backup = backup_dir / (target.stem + '_' + stamp + '_' + uuid.uuid4().hex[:8] + target.suffix)
            with target.open('rb') as old, backup.open('xb') as saved:
                shutil.copyfileobj(old, saved)
            if hashlib.sha256(backup.read_bytes()).digest() != hashlib.sha256(target.read_bytes()).digest():
                raise RuntimeError('Existing output changed during backup; not overwritten')
        os.replace(staged, target)
    finally:
        Path(staged).unlink(missing_ok=True)
    return str(backup) if backup else None


def check_source(report):
    source = Path(report['source']).resolve(strict=True)
    if source.suffix.lower() != '.cdr':
        raise ValueError('Expected CDR source')
    if hashlib.sha256(source.read_bytes()).hexdigest() != report['sha256']:
        raise RuntimeError('Source CDR changed; repeat the import before saving')
    return source


def prepare_outputs(report_path):
    report_path = Path(report_path)
    report = json.loads(report_path.read_text(encoding='utf-8'))
    source = check_source(report)
    target = source.with_suffix('.dxf')
    backup = publish_file(report['dxf'], target)
    report.update(saved_dxf=str(target), dxf_backup=backup,
                  aspire_target=str(source.with_suffix('.crv3d')),
                  aspire_staging=str(report_path.parent / 'project.crv3d'))
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    return report


def publish_job(report_path):
    report_path = Path(report_path)
    report = json.loads(report_path.read_text(encoding='utf-8'))
    source = check_source(report)
    staging = report_path.parent / 'project.crv3d'
    target = source.with_suffix('.crv3d')
    backup = publish_file(staging, target)
    report.update(saved_aspire=str(target), aspire_backup=backup)
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    # Lua only reopens the job after this success marker has been written.
    (report_path.parent / 'published-job.txt').write_text(str(target), encoding='utf-8')
    return target
