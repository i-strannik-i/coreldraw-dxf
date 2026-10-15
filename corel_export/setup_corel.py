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


def corel_executable(major=25):
    versions = discover()
    from corel_connection import SUPPORTED_MAJORS, version_label
    if major not in SUPPORTED_MAJORS or major not in versions:
        detected = ', '.join('v' + str(v) for v in versions) or 'не найден'
        raise RuntimeError('Обнаружен CorelDRAW: ' + detected + '. Выберите установленную версию CorelDRAW 2024–2026 x64.')
    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, f'CorelDRAW.Application.{major}\\CLSID') as key:
        clsid = winreg.QueryValue(key, None)
    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, 'CLSID\\' + clsid + r'\LocalServer32') as key:
        command = winreg.QueryValue(key, None)
    path = Path(command.split('"')[1] if command.startswith('"') else command.split(' /')[0])
    if not path.is_file():
        raise RuntimeError('Не найден ' + version_label(major) + '. Установите его с компонентом VBA.')
    validate_binary(path)
    return path


def corel_running():
    result = subprocess.run(['tasklist', '/FI', 'IMAGENAME eq CorelDRW.exe', '/FO', 'CSV', '/NH'],
                            capture_output=True, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
    return any(row and row[0].lower() == 'coreldrw.exe'
               for row in csv.reader(result.stdout.decode('utf-8', errors='replace').splitlines()))


def wait_for_corel_exit(cancelled, status, *, interval=.5):
    """Wait for shutdown and workspace flushing, without closing Corel ourselves."""
    quiet = 0
    was_running = False
    while not cancelled.is_set():
        if corel_running():
            if not was_running:
                status('CorelDRAW открыт или ещё завершает работу. Сохраните документы и закройте его. Установка продолжится автоматически; перезапуск установщика не нужен.')
            was_running = True
            quiet = 0
        else:
            if was_running:
                status('Окно CorelDRAW закрыто. Ожидание завершения процесса и сохранения рабочей среды…')
            was_running = False
            quiet += 1
            if quiet >= 3:
                return not cancelled.is_set()
        if cancelled.wait(interval):
            break
    return False


def install(package, draw, status=lambda message: None, *, major=25, gms_path=None, workspace_path=None):
    if corel_running():
        raise RuntimeError('Закройте CorelDRAW, сохранив свои документы, и повторите установку.')
    package, draw = Path(package), Path(draw)
    from corel_connection import SUPPORTED_MAJORS
    if major not in SUPPORTED_MAJORS:
        raise RuntimeError('Версия CorelDRAW не поддерживается этим установщиком.')
    workspace_folder = Path(workspace_path) if workspace_path else draw / 'Workspace'
    workspaces = sorted(workspace_folder.glob('*.cdws'))
    if not workspaces or not (workspace_folder / '_default.cdws').exists():
        raise RuntimeError('Сначала один раз откройте и закройте выбранный CorelDRAW, затем повторите установку.')
    stamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S_%f')
    backup = draw / 'SkladDXF_backups' / stamp
    backup.mkdir(parents=True)
    gms = Path(gms_path) if gms_path else draw / 'GMS'
    gms.mkdir(parents=True, exist_ok=True)
    staged = Path(tempfile.mkdtemp(prefix='SkladDXF-install-', dir=draw))
    changes = []
    committed = []
    try:
        status('Подготовка конвертера и резервных копий...')
        runtime = staged / 'SkladCorelDXF_runtime'
        shutil.copytree(package / 'runtime', runtime)
        (runtime / 'corel-target.json').write_text(json.dumps({'major': major}), encoding='utf-8')
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
    root.geometry('640x540')
    root.resizable(False, False)
    root.configure(bg='#f3f5f3')
    frame = ttk.Frame(root, padding=26)
    frame.pack(fill='both', expand=True)
    ttk.Label(frame, text='CorelDRAW → DXF', font=('Segoe UI', 21, 'bold')).pack(anchor='w')
    ttk.Label(frame, text='Версия ' + VERSION + ' · CorelDRAW 2024–2026 x64', font=('Segoe UI', 11)).pack(anchor='w', pady=(5, 12))
    from corel_connection import SUPPORTED_MAJORS, version_label
    versions = [major for major in discover() if major in SUPPORTED_MAJORS]
    targets = {version_label(major): major for major in versions}
    ttk.Label(frame, text='Установить в CorelDRAW:').pack(anchor='w')
    target_choice = tk.StringVar(value=version_label(max(versions)) if versions else '')
    selector = ttk.Combobox(frame, textvariable=target_choice, values=list(targets), state='readonly')
    selector.pack(fill='x', pady=(4, 12))
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
    selected_major = None
    waiting_for_corel = False
    cancel_wait = threading.Event()

    def worker():
        try:
            if not wait_for_corel_exit(cancel_wait, lambda s: events.put(('status', s))):
                events.put(('cancelled', 'Ожидание отменено. Нажмите «Установить», чтобы повторить.'))
                return
            events.put(('checking', None))
            base = Path(getattr(sys, '_MEIPASS', Path(__file__).parent))
            payload = base / 'CorelDXF_2024.zip'
            with tempfile.TemporaryDirectory(prefix='CorelDXF-payload-') as temp:
                events.put(('status', 'Проверка 1/3: запуск CorelDRAW и загрузка VBA. До 2 минут; окно установщика остаётся доступным.'))
                if cancel_wait.is_set():
                    events.put(('cancelled', 'Ожидание отменено. Нажмите «Установить», чтобы повторить.'))
                    return
                result = run_probe(temp, major=selected_major)
                if not result['ok']:
                    events.put(('prerequisite', result))
                    return
                with zipfile.ZipFile(payload) as archive:
                    for name in archive.namelist():
                        if not (Path(temp) / name).resolve().is_relative_to(Path(temp).resolve()):
                            raise ValueError('Invalid installer payload')
                    archive.extractall(temp)
                workspace_path = Path(result['workspace_path'])
                gms_path = Path(result['gms_path'])
                if not workspace_path.is_absolute() or not gms_path.is_absolute():
                    raise RuntimeError('Corel вернул некорректные пользовательские папки.')
                draw = workspace_path.parent
                if cancel_wait.is_set():
                    events.put(('cancelled', 'Ожидание отменено. Установка не выполнялась.'))
                    return
                backup = install(Path(temp) / 'CorelDXF_2024', draw, lambda s: events.put(('status', s)),
                                 major=selected_major, gms_path=gms_path, workspace_path=workspace_path)
                events.put(('status', 'Проверка 3/3: запуск установленного макроса в CorelDRAW...'))
                result = run_probe(temp, installed=True, major=selected_major)
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
        nonlocal busy, executable, selected_major, waiting_for_corel
        if busy:
            return
        try:
            selected_major = targets.get(target_choice.get())
            if selected_major is None:
                raise RuntimeError('Не найден CorelDRAW 2024–2026 x64. Установите Corel с VBA и откройте установщик заново.')
            executable = corel_executable(selected_major)
        except Exception as error:
            label.configure(text=str(error))
            return
        busy = True
        waiting_for_corel = True
        cancel_wait.clear()
        selector.configure(state='disabled')
        repair_button.pack_forget()
        maintenance_button.pack_forget()
        button.configure(state='normal', text='Отменить ожидание', command=cancel_wait.set)
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
        selector.configure(state='disabled')
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
            label.configure(text='Локальный мастер Corel отсутствует. Нужен официальный комплект выбранной версии CorelDRAW от вашей организации.')
            return
        # Use Corel maintenance UI, never an uninstall command or global security change.
        os.startfile(str(path), arguments='/arp')
        label.configure(text='Открыт мастер Corel: «Изменить» → «Утилиты» → Visual Basic for Applications. После завершения нажмите «Проверить и установить». Windows может запросить права администратора.')

    def poll():
        nonlocal busy, waiting_for_corel
        while not events.empty():
            kind, value = events.get_nowait()
            if kind == 'status':
                label.configure(text=value)
            elif kind == 'checking':
                waiting_for_corel = False
                button.configure(state='disabled', text='Установка…', command=start)
            else:
                busy = False
                waiting_for_corel = False
                selector.configure(state='readonly')
                progress.stop()
                button.configure(state='normal', text='Проверить и установить', command=start)
                repair_button.configure(state='normal')
                if kind == 'done':
                    label.configure(text='Проверено: ' + version_label(selected_major) + ' x64, VBA и компоненты плагина. Установка завершена. Для другой версии выберите её в списке и нажмите «Установить».')
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
    def change_target(event):
        repair_button.pack_forget()
        maintenance_button.pack_forget()
        button.configure(text='Установить', command=start)
        label.configure(text='Выбрано: ' + target_choice.get() + '. Сохраните документы и закройте CorelDRAW перед установкой.')
    selector.bind('<<ComboboxSelected>>', change_target)
    def close():
        if waiting_for_corel:
            cancel_wait.set()
            root.destroy()
        elif not busy:
            root.destroy()
    root.protocol('WM_DELETE_WINDOW', close)
    poll()
    root.mainloop()


if __name__ == '__main__':
    if '--probe' in sys.argv:
        try:
            from .setup_checks import probe
        except ImportError:
            from setup_checks import probe
        major = int(sys.argv[sys.argv.index('--corel-major') + 1]) if '--corel-major' in sys.argv else 25
        probe(sys.argv[sys.argv.index('--probe') + 1], '--installed' in sys.argv, major=major)
    else:
        main()
