"""Standalone runtime for the CorelDRAW DXF macro (no Aspire dependency)."""
import json
import os
from pathlib import Path
import queue
import sys
import threading
import time
import traceback

if not getattr(sys, 'frozen', False):
    sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'engine'))


def read_request(path):
    lines = Path(path).read_text(encoding='utf-16').splitlines()
    if len(lines) not in (3, 4) or lines[2] not in ('selection', 'page'):
        raise ValueError('Invalid export request')
    source = Path(lines[0]).resolve(strict=True)
    if source.suffix.lower() != '.cdr':
        raise ValueError('Expected a CDR snapshot')
    target = Path(lines[1]) if lines[1] else None
    if target and target.suffix.lower() != '.dxf':
        raise ValueError('Expected a DXF target')
    page = int(lines[3]) if len(lines) == 4 else 1
    if page < 1 or (lines[2] == 'selection' and page != 1):
        raise ValueError('Invalid source page')
    return source, target, lines[2], page


def export_request(request, target, progress):
    from convert_cdr import convert
    from save_outputs import publish_file
    source, _, scope, page = read_request(request)
    if Path(target).suffix.lower() != '.dxf':
        raise ValueError('Output extension must be .dxf')
    report = convert(source, Path(request).parent / 'output', .1, progress, .1, page_index=page, connected=True)
    validation = report.get('validation', {})
    if validation.get('open_required') or validation.get('zero_length'):
        issues = validation.get('open_required', [])
        layers = ', '.join(sorted({item['layer'] for item in issues}))
        raise ValueError(f'Проверка контуров: незамкнутых для резки/выборки — {len(issues)} ({layers}); '
                         f'нулевых участков — {len(validation.get("zero_length", []))}. '
                         'DXF не заменён. Нажмите «Проверить» на панели плагина. Допуск 0,1 мм не замыкает исходные контуры автоматически.')
    progress('Сохранение DXF')
    report['backup'] = publish_file(report['dxf'], target)
    report.update(saved_dxf=str(target), scope=scope)
    (Path(request).parent / 'result.json').write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    return report


def main():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    import win32event
    import win32api
    import winerror
    request = Path(sys.argv[1]).resolve(strict=True)
    root = tk.Tk()
    root.withdraw()
    mutex = win32event.CreateMutex(None, False, 'Local\\SkladCorelDXFExport')
    duplicate = win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS
    try:
        if duplicate:
            messagebox.showerror('Экспорт DXF', 'Другой экспорт ещё выполняется.', parent=root)
            return
        source, target, scope, _ = read_request(request)
        if target is None:
            choice = filedialog.asksaveasfilename(parent=root, title='Сохранить DXF',
                defaultextension='.dxf', filetypes=[('DXF', '*.dxf')])
            if not choice:
                return
            target = Path(choice)
        root.title('CorelDRAW → DXF')
        root.geometry('740x460')
        root.minsize(640,400)
        root.resizable(True, True)
        footer = ttk.Frame(root, padding=(18,8,18,16))
        footer.pack(side='bottom',fill='x')
        frame = ttk.Frame(root, padding=18)
        frame.pack(fill='both', expand=True)
        ttk.Label(frame, text=target.name, font=('Segoe UI', 11, 'bold'), wraplength=565).pack(anchor='w')
        ttk.Label(frame, text=('Выделенные объекты' if scope == 'selection' else 'Весь текущий лист') +
                  '  |  мм  |  допуск 0,1 мм').pack(anchor='w', pady=(5, 12))
        stage = tk.StringVar(value='Подготовка…')
        detail = tk.StringVar()
        clock = tk.StringVar()
        ttk.Label(frame, textvariable=stage, wraplength=565).pack(anchor='w')
        bar = ttk.Progressbar(frame, mode='indeterminate')
        bar.pack(fill='x', pady=12)
        bar.start(20)
        detail_area=ttk.Frame(frame);detail_area.pack(fill='both',expand=True)
        detail_text=tk.Text(detail_area,height=7,wrap='word',font=('Segoe UI',10),
                            relief='solid',borderwidth=1,padx=10,pady=8,state='disabled')
        scroll=ttk.Scrollbar(detail_area,orient='vertical',command=detail_text.yview)
        scroll.pack(side='right',fill='y');detail_text.pack(side='left',fill='both',expand=True)
        detail_text.configure(yscrollcommand=scroll.set)
        def render_detail(*args):
            detail_text.configure(state='normal');detail_text.delete('1.0','end')
            detail_text.insert('1.0',detail.get());detail_text.configure(state='disabled')
        detail.trace_add('write',render_detail)
        ttk.Label(frame, textvariable=clock).pack(anchor='w', pady=5)
        events = queue.Queue()
        cancel = threading.Event()
        started = time.monotonic()
        state = {'finished': False}

        def stop():
            if state['finished']:
                root.destroy()
                return
            cancel.set()
            button.configure(state='disabled')
            stage.set('Отмена: ожидаю завершения текущей операции CorelDRAW…')

        def open_log():
            path=request.parent/'error.txt'
            try:os.startfile(str(path if path.is_file() else request.parent))
            except OSError as error:messagebox.showerror('Журнал экспорта',str(error),parent=root)
        log_button=ttk.Button(footer,text='Открыть журнал',command=open_log,state='disabled')
        log_button.pack(side='left')
        button = ttk.Button(footer, text='Отменить', command=stop,width=16)
        button.pack(side='right')
        root.protocol('WM_DELETE_WINDOW', stop)

        def worker():
            import pythoncom
            from import_progress import ConversionCancelled
            pythoncom.CoInitialize()
            last = [0.0, None]
            def progress(label, **info):
                if cancel.is_set():
                    raise ConversionCancelled('Экспорт отменён')
                now = time.monotonic()
                if label != last[1] or now - last[0] > .2:
                    events.put(('progress', (label, info)))
                    last[:] = [now, label]
            try:
                report = export_request(request, target, progress)
                events.put(('done', report))
            except ConversionCancelled:
                events.put(('cancelled', None))
            except Exception as exc:
                (request.parent / 'error.txt').write_text(traceback.format_exc(), encoding='utf-8')
                events.put(('error', str(exc)))
            finally:
                pythoncom.CoUninitialize()

        def poll():
            while not events.empty():
                kind, data = events.get_nowait()
                if kind == 'progress' and not cancel.is_set():
                    label, info = data
                    stage.set(label)
                    detail.set(' | '.join(str(info[k]) for k in ('layer', 'segments', 'contours') if k in info))
                elif kind in ('done', 'cancelled', 'error'):
                    state['finished'] = True
                    bar.stop()
                    button.configure(text='Закрыть', state='normal')
                    log_button.configure(state='normal')
                    if kind == 'done':
                        stage.set('DXF сохранён')
                        detail.set(str(target))
                        root.after(2300, root.destroy)
                    elif kind == 'cancelled':
                        stage.set('Экспорт отменён. DXF не заменён.')
                        root.after(1500, root.destroy)
                    else:
                        stage.set('Не удалось экспортировать. Исходный CDR не изменён.')
                        detail.set(data + '\n\nЖурнал: ' + str(request.parent / 'error.txt'))
                    return
            seconds = int(time.monotonic() - started)
            clock.set(f'Прошло {seconds // 60:02d}:{seconds % 60:02d}')
            root.after(150, poll)

        root.deiconify()
        threading.Thread(target=worker, daemon=False).start()
        poll()
        root.mainloop()
    finally:
        win32api.CloseHandle(mutex)
        # Snapshots can be large; retain diagnostics, not another copy of the drawing.
        if not duplicate:
            try:
                if source.name == 'snapshot.cdr' and source.parent == request.parent:
                    source.unlink(missing_ok=True)
            except (UnboundLocalError, OSError):
                pass


if __name__ == '__main__':
    main()
