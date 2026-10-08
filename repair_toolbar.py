"""Repair only the DXF toolbar in closed Corel workspaces, preserving backups."""
import datetime
import os
from pathlib import Path
import shutil

from corel_export.setup_corel import corel_running
from corel_export.workspace_install import merge_workspace


def main():
    if corel_running():
        raise RuntimeError('Close CorelDRAW before repairing the saved workspace.')
    draw = Path(os.environ['APPDATA']) / 'Corel/CorelDRAW Graphics Suite 2024/Draw'
    icons = Path(__file__).parent / 'corel_export/icons'
    backup = draw / 'SkladDXF_backups' / ('toolbar_' + datetime.datetime.now().strftime('%Y%m%d_%H%M%S_%f'))
    pending = []
    for workspace in (draw / 'Workspace').glob('*.cdws'):
        try:
            updated = merge_workspace(workspace.read_bytes(), icons)
        except ValueError:
            if workspace.name == '_default.cdws':
                raise
            continue
        pending.append((workspace, updated))
    if not pending:
        raise RuntimeError('No compatible workspace found')
    backup.mkdir(parents=True)
    for workspace, _ in pending:
        shutil.copy2(workspace, backup / workspace.name)
    if corel_running():
        raise RuntimeError('CorelDRAW opened during repair; no workspaces changed')
    committed = []
    try:
        for workspace, updated in pending:
            temp = workspace.with_suffix('.cdws.pending')
            temp.write_bytes(updated)
            os.replace(temp, workspace)
            committed.append(workspace)
    except Exception:
        for workspace in committed:
            shutil.copy2(backup / workspace.name, workspace)
        raise
    shutil.copytree(icons, draw / 'GMS/SkladCorelDXF_runtime/icons', dirs_exist_ok=True)
    print('Backup:', backup)


if __name__ == '__main__':
    main()
