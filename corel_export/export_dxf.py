"""Standalone runtime for the CorelDRAW DXF macro (no Aspire dependency)."""
import json
import os
from pathlib import Path
import queue
import sys
import threading
import time
import traceback
from collections import Counter

if not getattr(sys, 'frozen', False):
    sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'engine'))


def read_request(path):
    lines = Path(path).read_text(encoding='utf-16').splitlines()
    if len(lines) not in (3, 4) or lines[2] not in ('selection', 'page'):
        raise ValueError('Некорректный запрос экспорта. Запустите экспорт заново из панели плагина.')
    source = Path(lines[0]).resolve(strict=True)
    if source.suffix.lower() != '.cdr':
        raise ValueError('Ожидалась временная копия в формате CDR.')
    target = Path(lines[1]) if lines[1] else None
    if target and target.suffix.lower() != '.dxf':
        raise ValueError('Выходной файл должен иметь расширение .dxf.')
    page = int(lines[3]) if len(lines) == 4 else 1
    if page < 1 or (lines[2] == 'selection' and page != 1):
        raise ValueError('Некорректная исходная страница.')
    return source, target, lines[2], page


class GeometryBlocked(ValueError):
    """The temporary conversion must not replace the user's DXF."""


def validation_message(validation):
    lines = []
    for key, title in (('open_required', 'Незамкнутые контуры резки/выборки'),
                       ('zero_length', 'Нулевые участки')):
        counts = Counter(item['layer'] for item in validation.get(key, []))
        if counts:
            lines.append(f'{title}: {sum(counts.values())}. По слоям: ' +
                         ', '.join(f'{layer}: {count}' for layer, count in sorted(counts.items())))
    return '\n'.join(lines)


def progress_text(label, info):
    fields = [('layer', 'Слой'), ('segments', 'Сегментов'), ('contours', 'Контуров'),
              ('optimized', 'Обработано'), ('total', 'Всего')]
    details = ' | '.join(f'{title}: {info[key]}' for key, title in fields if key in info)
    return label + (': ' + details if details else '')


def export_request(request, target, progress, confirm_warnings=None):
    from convert_cdr import convert
    from save_outputs import publish_file
    source, _, scope, page = read_request(request)
    if Path(target).suffix.lower() != '.dxf':
        raise ValueError('Выходной файл должен иметь расширение .dxf.')
    report = convert(source, Path(request).parent / 'output', .1, progress, .1, page_index=page, connected=True, check_vectors=True)
    validation = report.get('validation', {})
    open_warning = validation_message({'open_required': validation.get('open_required', [])})
    if open_warning:
        report.setdefault('warnings', []).append(
            open_warning + '. Контуры сохранены открытыми, без автоматического соединения. '
            'Если это не задумано, используйте «Проверить» на панели плагина.')
    problems = validation_message({'zero_length': validation.get('zero_length', [])})
    if problems:
        raise GeometryBlocked(problems + '\nЭкспорт заблокирован. DXF не сохранён и существующий файл не заменён. '
                              'Нажмите «Проверить» на панели плагина: ошибки будут показаны в общем списке и отмечены на чертеже. '
                              'Исправьте их и повторите экспорт. Допуск 0,1 мм не замыкает контуры автоматически.')
    check = report.get('vector_check')
    if not check or not check.get('complete'):
        raise GeometryBlocked('Проверка векторов не завершена. DXF не сохранён. Используйте «Проверить» на панели плагина.')
    if any(issue['kind'] == 'zero' for issue in check['issues']):
        raise GeometryBlocked('Обнаружены нулевые участки. DXF не сохранён. Используйте «Проверить» на панели плагина.')
    warnings = [issue for issue in check['issues'] if issue['kind'] in ('intersection', 'overlap', 'duplicate')]
    if warnings:
        counts = Counter((issue['layer'], issue['kind']) for issue in warnings)
        names = {'intersection': 'пересечения/касания', 'overlap': 'наложения', 'duplicate': 'группы дублей'}
        summary = '\n'.join(f'{layer}: {names[kind]} — {count}' for (layer, kind), count in sorted(counts.items()))
        summary += '\nЭто места для визуальной проверки, а не подтверждённые дефекты. Допуск проверки: 0,1 мм.'
        if confirm_warnings is None or not confirm_warnings(summary):
            from import_progress import ConversionCancelled
            raise ConversionCancelled('Экспорт с предупреждениями не подтверждён. DXF не заменён.')
        report.setdefault('warnings', []).append(summary)
        report['warnings_accepted'] = True
    report['check_status'] = 'Экспортировано с предупреждениями' if report.get('warnings') else 'Проверка пройдена'
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
        ttk.Label(frame, textvariable=clock).pack(anchor='w', pady=5)
        events = queue.Queue()
        cancel = threading.Event()
        decision_ready = threading.Event()
        decision = {'accepted': False}
        started = time.monotonic()
        state = {'finished': False}
        log_path = request.parent / 'operations.log'
        log_path.write_text('', encoding='utf-8-sig')

        def append_log(message):
            seconds = int(time.monotonic() - started)
            line = f'[{seconds // 60:02d}:{seconds % 60:02d}] {message}\n'
            with log_path.open('a', encoding='utf-8') as stream:
                stream.write(line)
            follow = detail_text.yview()[1] >= .99
            detail_text.configure(state='normal')
            detail_text.insert('end', line)
            detail_text.configure(state='disabled')
            if follow:
                detail_text.see('end')

        append_log(f'Экспорт: {target}')
        append_log('Оптимизация: прямые сохраняются прямыми; кривые упрощаются и аппроксимируются '
                   'линиями/дугами. Общий допуск 0,1 мм. Автоматического замыкания нет.')
        append_log('Подробная проверка и метки ошибок доступны через «Проверить» на панели плагина.')

        def stop():
            if state['finished']:
                root.destroy()
                return
            cancel.set()
            button.configure(state='disabled')
            stage.set('Отмена: ожидаю завершения текущей операции CorelDRAW…')

        def open_log():
            try:os.startfile(str(log_path))
            except OSError as error:messagebox.showerror('Журнал экспорта',str(error),parent=root)
        log_button=ttk.Button(footer,text='Открыть журнал',command=open_log)
        log_button.pack(side='left')
        button = ttk.Button(footer, text='Отменить', command=stop,width=16)
        button.pack(side='right')
        choice_frame = ttk.Frame(root, padding=(18, 4))
        def choose(accepted):
            decision['accepted'] = accepted
            if not accepted:
                cancel.set()
            decision_ready.set()
            choice_frame.pack_forget()
            button.configure(state='normal')
            bar.start(20)
        return_button = ttk.Button(choice_frame, text='Вернуться к проверке', command=lambda: choose(False))
        return_button.pack(side='left')
        ttk.Button(choice_frame, text='Экспортировать с предупреждениями', command=lambda: choose(True)).pack(side='right')
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
                def confirm(summary):
                    events.put(('confirm', summary))
                    while not decision_ready.wait(.1):
                        if cancel.is_set():
                            return False
                    return decision['accepted'] and not cancel.is_set()
                report = export_request(request, target, progress, confirm)
                events.put(('done', report))
            except ConversionCancelled:
                events.put(('cancelled', None))
            except Exception as exc:
                (request.parent / 'error.txt').write_text(traceback.format_exc(), encoding='utf-8')
                events.put(('error', str(exc)))
            finally:
                pythoncom.CoUninitialize()

        last_logged = [None, 0.0]

        def poll():
            while not events.empty():
                kind, data = events.get_nowait()
                if kind == 'confirm':
                    append_log('Требуется подтверждение:\n' + data)
                    stage.set('Найдены предупреждения. DXF ещё не сохранён.')
                    bar.stop()
                    choice_frame.pack(side='bottom', fill='x', before=frame)
                    return_button.focus_set()
                elif kind == 'progress' and not cancel.is_set():
                    label, info = data
                    stage.set(progress_text(label, info))
                    key = (label, info.get('layer'))
                    now = time.monotonic()
                    if key != last_logged[0] or now - last_logged[1] >= 5:
                        append_log(progress_text(label, info))
                        last_logged[:] = [key, now]
                elif kind in ('done', 'cancelled', 'error'):
                    state['finished'] = True
                    bar.stop()
                    button.configure(text='Закрыть', state='normal')
                    log_button.configure(state='normal')
                    if kind == 'done':
                        stage.set(data.get('check_status', 'DXF сохранён'))
                        append_log(data.get('check_status', 'DXF сохранён'))
                        append_log(f'Сегментов до оптимизации: {data.get("segments_before", "—")}; '
                                   f'после: {data.get("segments_after", "—")}.')
                        for warning in data.get('warnings', []):
                            append_log(f'Предупреждение: {warning}')
                        append_log(f'DXF сохранён: {target}')
                    elif kind == 'cancelled':
                        stage.set('Экспорт отменён. DXF не заменён.')
                        append_log('Экспорт отменён. DXF не заменён.')
                        append_log('Для просмотра и исправления найденных мест нажмите «Проверить» на панели плагина.')
                    else:
                        stage.set('Не удалось экспортировать. Исходный CDR не изменён.')
                        append_log(data)
                        append_log('Технические подробности: ' + str(request.parent / 'error.txt'))
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
