"""Per-user installer. Corel must be closed before committing workspace changes."""
from pathlib import Path
import csv
import datetime
import json
import os
import queue
import shutil
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
from tkinter import ttk
import traceback
import winreg
import zipfile

try:
    from .version import VERSION
    from .workspace_install import merge_workspace
    from .setup_checks import discover, validate_binary, repair_path, run_probe, install_vba
except ImportError:
    from version import VERSION
    from workspace_install import merge_workspace
    from setup_checks import discover, validate_binary, repair_path, run_probe, install_vba


def corel_executable():
    versions = discover()
    if 25 not in versions:
        detected = ', '.join('v' + str(v) for v in versions) or 'не найден'
        raise RuntimeError('Обнаружен CorelDRAW: ' + detected + '. Этот установщик поддерживает CorelDRAW 2024 (v25) x64. Другие версии не изменены.')
    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r'CorelDRAW.Application.25\CLSID') as key:
        clsid = winreg.QueryValue(key, None)
    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, 'CLSID\\' + clsid + r'\LocalServer32') as key:
        command = winreg.QueryValue(key, None)
    path = Path(command.split('"')[1] if command.startswith('"') else command.split(' /')[0])
    if not path.is_file():
        raise RuntimeError('Не найден CorelDRAW 2024. Установите его с компонентом VBA.')
    validate_binary(path)
    return path


def corel_running():
    result = subprocess.run(['tasklist', '/FI', 'IMAGENAME eq CorelDRW.exe', '/FO', 'CSV', '/NH'],
                            capture_output=True, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
    return any(row and row[0].lower() == 'coreldrw.exe'
               for row in csv.reader(result.stdout.decode('utf-8', errors='replace').splitlines()))


def install(package, draw, status=lambda message: None):
    if corel_running():
        raise RuntimeError('Закройте CorelDRAW, сохранив свои документы, и повторите установку.')
    package, draw = Path(package), Path(draw)
    workspaces = sorted((draw / 'Workspace').glob('*.cdws'))
    if not workspaces or not (draw / 'Workspace' / '_default.cdws').exists():
        raise RuntimeError('Сначала один раз откройте и закройте CorelDRAW 2024, затем повторите установку.')
    stamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S_%f')
    backup = draw / 'SkladDXF_backups' / stamp
    backup.mkdir(parents=True)
    gms = draw / 'GMS'
    gms.mkdir(exist_ok=True)
    staged = Path(tempfile.mkdtemp(prefix='SkladDXF-install-', dir=draw))
    changes = []
    committed = []
    try:
        status('Подготовка конвертера и резервных копий...')
        runtime = staged / 'SkladCorelDXF_runtime'
        shutil.copytree(package / 'runtime', runtime)
        shutil.copytree(package / 'icons', runtime / 'icons', dirs_exist_ok=True)
        macro = staged / 'SkladCorelDXF.gms'
        shutil.copy2(package / macro.name, macro)
        changes.extend([(runtime, gms / runtime.name), (macro, gms / macro.name)])
        status('Настройка кнопок DXF и справки...')
        for workspace in workspaces:
            try:
                updated = merge_workspace(workspace.read_bytes(), package / 'icons')
            except ValueError:
                if workspace.name == '_default.cdws':
                    raise
                continue
            pending = staged / workspace.name
            pending.write_bytes(updated)
            changes.append((pending, workspace))
        if corel_running():
            raise RuntimeError('CorelDRAW был открыт во время установки. Закройте его и повторите.')
        for number, (pending, destination) in enumerate(changes):
            saved = backup / (str(number) + '_' + destination.name)
            if destination.exists():
                os.replace(destination, saved)
            else:
                saved = None
            committed.append((destination, saved))
            os.replace(pending, destination)
        (backup / 'manifest.json').write_text(json.dumps(
            [{'destination': str(d), 'backup': str(b) if b else None} for d, b in committed],
            ensure_ascii=False, indent=2), encoding='utf-8')
        status('Установка завершена. Отдельная панель Sklad DXF появится в CorelDRAW.')
        return backup
    except Exception:
        for destination, saved in reversed(committed):
            if destination.is_dir():
                shutil.rmtree(destination)
            elif destination.exists():
                destination.unlink()
            if saved and saved.exists():
                os.replace(saved, destination)
        raise
    finally:
        shutil.rmtree(staged)


def main():
    root = tk.Tk()
    root.title('Установка CorelDRAW → DXF ' + VERSION)
    root.geometry('620x470')
    root.resizable(False, False)
    root.configure(bg='#f3f5f3')
    frame = ttk.Frame(root, padding=26)
    frame.pack(fill='both', expand=True)
    ttk.Label(frame, text='CorelDRAW → DXF', font=('Segoe UI', 21, 'bold')).pack(anchor='w')
    ttk.Label(frame, text='Версия ' + VERSION + ' · CorelDRAW 2024', font=('Segoe UI', 11)).pack(anchor='w', pady=(5, 22))
    ttk.Label(frame, text='• Экспорт всего листа или выделенных объектов\n• Перемещаемая панель Sklad DXF и настройка хоткея\n• Без Python, Aspire и ручного добавления макросов',
              font=('Segoe UI', 10), justify='left').pack(anchor='w')
    label = ttk.Label(frame, text='Сохраните документы и закройте CorelDRAW. Установщик сам проверит версию, VBA и запуск макроса. Во время проверки Corel может открыться.',
                      wraplength=560, font=('Segoe UI', 10))
    label.pack(anchor='w', pady=(20, 10))
    progress = ttk.Progressbar(frame, mode='indeterminate')
    progress.pack(fill='x')
    row = ttk.Frame(frame)
    row.pack(fill='x', pady=(16, 0))
    events = queue.Queue()
    busy = False
    executable = None

    def worker():
        try:
            base = Path(getattr(sys, '_MEIPASS', Path(__file__).parent))
            payload = base / 'CorelDXF_2024.zip'
            with tempfile.TemporaryDirectory(prefix='CorelDXF-payload-') as temp:
                events.put(('status', 'Проверка 1/3: запуск CorelDRAW и загрузка VBA. До 2 минут; окно установщика остаётся доступным.'))
                result = run_probe(temp)
                if not result['ok']:
                    events.put(('prerequisite', result))
                    return
                with zipfile.ZipFile(payload) as archive:
                    for name in archive.namelist():
                        if not (Path(temp) / name).resolve().is_relative_to(Path(temp).resolve()):
                            raise ValueError('Invalid installer payload')
                    archive.extractall(temp)
                draw = Path(os.environ['APPDATA']) / 'Corel/CorelDRAW Graphics Suite 2024/Draw'
                backup = install(Path(temp) / 'CorelDXF_2024', draw, lambda s: events.put(('status', s)))
                events.put(('status', 'Проверка 3/3: запуск установленного макроса в CorelDRAW...'))
                result = run_probe(temp, installed=True)
                if not result['ok']:
                    result['message'] = 'Файлы установлены, но запуск не подтверждён. ' + result['message'] + '\nРезервная копия: ' + str(backup)
                    events.put(('prerequisite', result))
                    return
            events.put(('done', str(backup)))
        except Exception as error:
            log = Path(tempfile.gettempdir()) / 'CorelDXF-setup-error.log'
            log.write_text(traceback.format_exc(), encoding='utf-8')
            events.put(('error', str(error) + '\nЖурнал: ' + str(log)))

    def start():
        nonlocal busy, executable
        try:
            executable = corel_executable()
            if corel_running():
                label.configure(text='CorelDRAW ещё открыт. Сохраните документы и закройте его. Затем нажмите «Установить».')
                return
        except Exception as error:
            label.configure(text=str(error))
            return
        busy = True
        repair_button.pack_forget()
        maintenance_button.pack_forget()
        button.configure(state='disabled')
        progress.start(12)
        threading.Thread(target=worker, daemon=True).start()

    def launch():
        subprocess.Popen([str(executable)])
        root.destroy()

    def repair():
        nonlocal busy
        if corel_running():
            label.configure(text='Перед изменением компонентов сохраните документы и закройте CorelDRAW.')
            return
        busy = True
        button.configure(state='disabled')
        repair_button.configure(state='disabled')
        progress.start(12)

        def add_components():
            try:
                logs = Path(os.environ['LOCALAPPDATA']) / 'SkladCorelDXF/setup-logs'
                logs.mkdir(parents=True, exist_ok=True)
                install_vba(executable, logs, lambda s: events.put(('status', s)))
                events.put(('repaired', ''))
            except Exception as error:
                events.put(('repair_failed', str(error)))
        threading.Thread(target=add_components, daemon=True).start()

    def maintenance():
        if corel_running():
            label.configure(text='Сначала сохраните документы и закройте CorelDRAW.')
            return
        path = repair_path(executable)
        if path is None:
            label.configure(text='Локальный мастер Corel отсутствует. Нужен официальный установочный комплект CorelDRAW 2024 от вашей организации. VBA отдельно из неизвестных источников не скачивается.')
            return
        # Use Corel maintenance UI, never an uninstall command or global security change.
        os.startfile(str(path), arguments='/arp')
        label.configure(text='Открыт мастер Corel: «Изменить» → «Утилиты» → Visual Basic for Applications. После завершения нажмите «Проверить и установить». Windows может запросить права администратора.')

    def poll():
        nonlocal busy
        while not events.empty():
            kind, value = events.get_nowait()
            if kind == 'status':
                label.configure(text=value)
            else:
                busy = False
                progress.stop()
                button.configure(state='normal')
                repair_button.configure(state='normal')
                if kind == 'done':
                    label.configure(text='Проверено: CorelDRAW 2024 x64, VBA и компоненты плагина. Панель: экспорт, проверка векторов, соединение, памятка слоёв. Установка завершена.')
                    button.configure(text='Открыть CorelDRAW', command=launch)
                elif kind == 'prerequisite':
                    message = value['message']
                    if value['kind'] == 'vba':
                        message = 'VBA не удалось запустить. Добавьте или восстановите компонент кнопкой ниже. Если VBA запрещён политикой организации, нужен администратор.\n' + message
                        repair_button.pack(side='left')
                    elif value['kind'] == 'macro':
                        message += '\nПроверьте разрешение макросов у администратора. Защита автоматически не отключается.'
                    label.configure(text=message[:950])
                    button.configure(text='Проверить и установить', command=start)
                elif kind == 'repaired':
                    start()
                elif kind == 'repair_failed':
                    label.configure(text=value)
                    maintenance_button.pack(side='left', padx=6)
                else:
                    label.configure(text=value)
        root.after(150, poll)

    button = ttk.Button(row, text='Установить', command=start)
    button.pack(side='right')
    repair_button = ttk.Button(row, text='Установить VBA', command=repair)
    maintenance_button = ttk.Button(frame, text='Открыть мастер Corel', command=maintenance)
    root.protocol('WM_DELETE_WINDOW', lambda: None if busy else root.destroy())
    poll()
    root.mainloop()


if __name__ == '__main__':
    if '--probe' in sys.argv:
        try:
            from .setup_checks import probe
        except ImportError:
            from setup_checks import probe
        probe(sys.argv[sys.argv.index('--probe') + 1], '--installed' in sys.argv)
    else:
        main()
