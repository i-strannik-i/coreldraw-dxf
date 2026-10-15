"""Select the Corel version belonging to this installed runtime."""
import json
from pathlib import Path
import sys

SUPPORTED_MAJORS = (25, 26, 27)


def version_label(major):
    return f'CorelDRAW {major + 1999} (v{major})'


def configured_major():
    folder = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).resolve().parents[1] / 'corel_export'
    path = folder / 'corel-target.json'
    if path.exists():
        major = json.loads(path.read_text(encoding='utf-8'))['major']
        if type(major) is not int or major not in SUPPORTED_MAJORS:
            raise ValueError('Некорректная версия Corel в настройках установки.')
        return major
    return None


def connect(major=None):
    import win32com.client
    major = major if major is not None else configured_major()
    if major is not None:
        app = win32com.client.gencache.EnsureDispatch(f'CorelDRAW.Application.{major}')
    else:
        # Source/development launch: use a single running supported version.
        # Never silently launch another installed Corel when selection is ambiguous.
        running = []
        for candidate in SUPPORTED_MAJORS:
            try:
                running.append(win32com.client.GetActiveObject(f'CorelDRAW.Application.{candidate}'))
            except Exception:
                continue
        if len(running) != 1:
            raise RuntimeError('Откройте одну версию CorelDRAW 2024–2026 или запустите инструмент из установленной панели DXF.')
        app = win32com.client.gencache.EnsureDispatch(running[0])
    actual = int(app.VersionMajor)
    if actual not in SUPPORTED_MAJORS or (major is not None and actual != major):
        raise RuntimeError('COM подключился к другой версии CorelDRAW. Повторите установку для нужной версии.')
    return app
