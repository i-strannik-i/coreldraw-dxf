"""Responsive progress window; Corel COM runs in a separate worker process."""
import json
import os
from pathlib import Path
import subprocess
import sys
import time


class ConversionCancelled(Exception):
    pass


class ProgressWriter:
    def __init__(self, path, cancel=None):
        self.path = Path(path)
        self.cancel = Path(cancel) if cancel else None
        self.last = 0
        self.stage = None

    def __call__(self, stage, **details):
        if self.cancel and self.cancel.exists():
            raise ConversionCancelled('Импорт отменён пользователем')
        now = time.monotonic()
        if stage == self.stage and now - self.last < .3 and not details.get('finished'):
            return
        payload = {'stage': stage, 'updated': time.time(), **details}
        temp = self.path.with_suffix('.tmp')
        temp.write_text(json.dumps(payload, ensure_ascii=False), encoding='utf-8')
        try:
            os.replace(temp, self.path)
        except PermissionError:
            # A Windows reader may briefly hold the status file open.
            return
        self.stage, self.last = stage, now


def run_with_progress(root, source, output, tolerance, request, optimize_tolerance=0):
    import tkinter as tk
    from tkinter import ttk
    progress = request.with_suffix('.progress.json')
    cancel = request.with_suffix('.cancel')
    error = progress.with_suffix('.error')
    for path in (progress, cancel, error):
        path.unlink(missing_ok=True)
    window = tk.Toplevel(root)
    window.title('Импорт CDR в Aspire')
    window.geometry('620x320')
    window.resizable(False, False)
    window.attributes('-topmost', True)
    frame = ttk.Frame(window, padding=18)
    frame.pack(fill='both', expand=True)
    ttk.Label(frame, text=Path(source).name, font=('Segoe UI', 11, 'bold'), wraplength=540).pack(anchor='w')
    stage = tk.StringVar(value='Запуск конвертера…')
    detail = tk.StringVar(value='')
    elapsed = tk.StringVar(value='')
    ttk.Label(frame, textvariable=stage, wraplength=540).pack(anchor='w', pady=(14, 4))
    ttk.Label(frame, textvariable=detail).pack(anchor='w')
    bar = ttk.Progressbar(frame, mode='indeterminate')
    bar.pack(fill='x', pady=10)
    bar.start(20)
    ttk.Label(frame, textvariable=elapsed, wraplength=540).pack(anchor='w')
    state = {'cancel': False, 'result': None, 'error': None}

    def cancel_import():
        state['cancel'] = True
        cancel.touch()
        button.configure(state='disabled')
        stage.set('Отмена: ожидаю освобождения CorelDRAW…')

    button = ttk.Button(frame, text='Отменить импорт', command=cancel_import)
    button.pack(anchor='e', pady=(8, 0))
    window.protocol('WM_DELETE_WINDOW', cancel_import)
    started = time.monotonic()
    try:
        worker = subprocess.Popen([
            sys.executable, str(Path(__file__).with_name('convert_cdr.py')),
            '--source', str(source), '--output', str(output), '--arc-tolerance', str(tolerance),
            '--optimize-tolerance', str(optimize_tolerance),
            '--progress', str(progress), '--cancel', str(cancel)],
            stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
    except Exception:
        window.destroy()
        raise

    def poll():
        age = time.monotonic() - started
        status = {}
        try:
            status = json.loads(progress.read_text(encoding='utf-8'))
        except (OSError, ValueError):
            pass
        if status and not state['cancel']:
            stage.set(status['stage'])
            detail.set('  |  '.join(str(v) for v in [status.get('layer', ''),
                f"Сегментов: {status['segments']}" if 'segments' in status else '',
                f"Контуров: {status['contours']}" if 'contours' in status else '',
                f"Участок: {status['optimized']}/{status['total']}" if 'optimized' in status else ''] if v))
        idle = time.time() - status.get('updated', time.time() - age)
        suffix = '  Пока нет новых данных от конвертера. Возможна загрузка CorelDRAW или ожидание его диалога; это не подтверждение зависания.' if idle > 30 else ''
        elapsed.set(f'Прошло {int(age)//60:02d}:{int(age)%60:02d}.' + suffix)
        if worker.poll() is None:
            window.after(150, poll)
            return
        if not state['cancel']:
            try:
                if worker.returncode:
                    raise RuntimeError(error.read_text(encoding='utf-8') if error.exists() else 'Конвертер завершился с ошибкой')
                state['result'] = json.loads((output/'report.json').read_text(encoding='utf-8'))
            except Exception as exc:
                state['error'] = exc
        bar.stop()
        if state['result']:
            report = state['result']
            stage.set('Готово. Передача контуров в Aspire…')
            detail.set(f"Сегментов: {report.get('segments_before', '?')} → {report.get('segments_after', '?')}")
            button.configure(state='disabled')
            window.after(1600, window.destroy)
            return
        window.destroy()

    window.after(100, poll)
    root.wait_window(window)
    if state['error']:
        raise state['error']
    return state['result']
