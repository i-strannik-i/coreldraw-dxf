"""Inline export UI. Geometry validation and atomic publication stay in export_dxf."""
import os
from pathlib import Path
import queue
import threading
import time
import traceback
import tkinter as tk
from tkinter import ttk, filedialog

try:
    from .export_dxf import export_request, read_request, progress_text
except ImportError:
    from export_dxf import export_request, read_request, progress_text


class ExportPanel:
    def __init__(self, parent, prepare, on_busy, on_check):
        self.parent=parent
        self.prepare=prepare
        self.on_busy=on_busy
        self.on_check=on_check
        self.busy=False
        self.events=queue.Queue()
        self.cancelled=threading.Event()
        self.decision_ready=threading.Event()
        self.accepted=False
        self.mutex=None
        self.request=None
        self.source=None
        self.log_path=None
        self.return_to_check=False
        self.started=0
        self.stage=tk.StringVar(value='Готово к экспорту.')
        self.elapsed=tk.StringVar(value='')
        self.actions=ttk.Frame(parent)
        self.actions.pack(fill='x',pady=(8,10))
        self.start_button=ttk.Button(self.actions,text='Экспортировать DXF',style='Accent.TButton',command=self.start)
        self.start_button.pack(side='left')
        self.check_button=ttk.Button(self.actions,text='Перейти к проверке',command=on_check)
        self.check_button.pack(side='left',padx=8)
        self.cancel_button=ttk.Button(self.actions,text='Отменить',command=self.cancel,state='disabled')
        self.cancel_button.pack(side='right')
        ttk.Label(parent,textvariable=self.stage,wraplength=920).pack(fill='x')
        self.bar=ttk.Progressbar(parent,mode='determinate')
        self.bar.pack(fill='x',pady=8)
        self.footer=ttk.Frame(parent)
        self.footer.pack(side='bottom',fill='x',pady=(8,0))
        ttk.Label(self.footer,textvariable=self.elapsed).pack(side='left')
        self.log_button=ttk.Button(self.footer,text='Открыть журнал',command=self.open_log,state='disabled')
        self.log_button.pack(side='right')
        self.choice=ttk.Frame(parent)
        ttk.Button(self.choice,text='Вернуться к проверке',command=lambda:self.choose(False)).pack(side='left')
        ttk.Button(self.choice,text='Экспортировать с предупреждениями',command=lambda:self.choose(True)).pack(side='right')
        self.log_area=ttk.Frame(parent)
        self.log_area.pack(fill='both',expand=True)
        self.text=tk.Text(self.log_area,height=6,wrap='word',font=('Segoe UI',10),state='disabled',padx=8,pady=8)
        scroll=ttk.Scrollbar(self.log_area,command=self.text.yview)
        scroll.pack(side='right',fill='y');self.text.pack(side='left',fill='both',expand=True)
        self.text.configure(yscrollcommand=scroll.set)
        self.last_logged=(None,0)

    def theme(self):
        style=ttk.Style(self.parent)
        self.text.configure(background=style.lookup('Treeview','background'),
                            foreground=style.lookup('Treeview','foreground'))

    def append(self, message):
        seconds=int(time.monotonic()-self.started) if self.started else 0
        line=f'[{seconds//60:02d}:{seconds%60:02d}] {message}\n'
        if self.log_path:
            try:
                with self.log_path.open('a',encoding='utf-8') as stream:stream.write(line)
            except OSError:
                line+='Не удалось записать журнал на диск. Текст сохранён в окне.\n'
        follow=self.text.yview()[1]>=.99
        self.text.configure(state='normal');self.text.insert('end',line);self.text.configure(state='disabled')
        if follow:self.text.see('end')

    def start(self):
        if self.busy:return
        import win32api, win32event, winerror
        handle=win32event.CreateMutex(None,False,'Local\\SkladCorelDXFExport')
        if win32api.GetLastError()==winerror.ERROR_ALREADY_EXISTS:
            win32api.CloseHandle(handle)
            self.stage.set('Другой экспорт ещё выполняется. Дождитесь его завершения.')
            return
        self.mutex=handle
        self.request=None;self.source=None;self.log_path=None
        self.cancelled.clear();self.decision_ready.clear();self.accepted=False
        self.return_to_check=False;self.last_logged=(None,0)
        self.started=time.monotonic()
        self.text.configure(state='normal');self.text.delete('1.0','end');self.text.configure(state='disabled')
        self.log_button.configure(state='disabled')
        self.busy=True
        try:
            self.on_busy(True)
            self.start_button.configure(state='disabled');self.check_button.configure(state='disabled')
            self.stage.set('Подготовка временной копии чертежа…')
            self.parent.update_idletasks()
            self.request=Path(self.prepare()).resolve(strict=True)
            self.source,target,scope,_=read_request(self.request)
            self.log_path=self.request.parent/'operations.log'
            self.log_path.write_text('',encoding='utf-8-sig')
            self.log_button.configure(state='normal')
            if target is None:
                target=filedialog.asksaveasfilename(parent=self.parent,title='Сохранить DXF',
                    defaultextension='.dxf',filetypes=[('DXF','*.dxf')])
                if not target:
                    self.append('Экспорт отменён до начала обработки.');self.finish('Экспорт отменён.');return
                target=Path(target)
            self.target=target
            self.append(f'Экспорт: {target}')
            self.append('Область: '+('выделенные объекты.' if scope=='selection' else 'текущая страница.'))
            self.append('Оптимизация: прямые сохраняются прямыми, кривые упрощаются в пределах допуска 0,1 мм. '
                        'Масштаб 1:1, миллиметры. Автоматического замыкания нет.')
            self.append('Перед сохранением: проверка по слоям. Нулевые участки блокируют экспорт; '
                        'пересечения и наложения требуют подтверждения.')
            self.cancel_button.configure(state='normal')
            self.bar.configure(mode='indeterminate');self.bar.start(20)
            threading.Thread(target=self.worker,daemon=False).start()
            self.parent.after(100,self.poll)
        except Exception as error:
            self.append(str(error));self.finish('Не удалось начать экспорт. DXF не заменён.')

    def worker(self):
        import pythoncom
        from import_progress import ConversionCancelled
        pythoncom.CoInitialize()
        last=[None,0]
        def progress(label, **info):
            if self.cancelled.is_set():raise ConversionCancelled('Экспорт отменён.')
            now=time.monotonic()
            if label!=last[0] or now-last[1]>.2:
                self.events.put(('progress',(label,info)));last[:]=[label,now]
        def confirm(summary):
            self.events.put(('confirm',summary))
            while not self.decision_ready.wait(.1):
                if self.cancelled.is_set():return False
            return self.accepted and not self.cancelled.is_set()
        try:
            result=export_request(self.request,self.target,progress,confirm)
            self.events.put(('done',result))
        except ConversionCancelled:self.events.put(('cancelled',None))
        except Exception as error:
            try:(self.request.parent/'error.txt').write_text(traceback.format_exc(),encoding='utf-8')
            except OSError:pass
            self.events.put(('error',str(error)))
        finally:pythoncom.CoUninitialize()

    def cancel(self):
        if not self.busy:return
        self.cancelled.set();self.decision_ready.set()
        self.choice.pack_forget();self.cancel_button.configure(state='disabled')
        self.stage.set('Отмена: ожидаю завершения текущей операции CorelDRAW…')

    def choose(self, accepted):
        self.accepted=accepted
        self.return_to_check=not accepted
        self.choice.pack_forget()
        self.append('Экспорт с предупреждениями подтверждён пользователем.' if accepted else 'Выбран возврат к проверке. DXF не заменён.')
        if not accepted:self.cancelled.set()
        self.decision_ready.set();self.bar.start(20)

    def finish(self, message):
        import win32api
        self.busy=False;self.bar.stop();self.choice.pack_forget()
        seconds=int(time.monotonic()-self.started)
        self.elapsed.set(f'Прошло {seconds//60:02d}:{seconds%60:02d}')
        self.bar.configure(mode='determinate',value=0)
        self.stage.set(message)
        self.start_button.configure(state='normal');self.check_button.configure(state='normal')
        self.cancel_button.configure(state='disabled')
        if self.source and self.request and self.source.name=='snapshot.cdr' and self.source.parent==self.request.parent:
            try:self.source.unlink(missing_ok=True)
            except OSError:pass
        if self.mutex is not None:win32api.CloseHandle(self.mutex);self.mutex=None
        self.on_busy(False)
        if self.return_to_check:self.on_check()

    def poll(self):
        while not self.events.empty():
            kind,data=self.events.get_nowait()
            if kind=='progress' and not self.cancelled.is_set():
                label,info=data;self.stage.set(progress_text(label,info))
                key=(label,info.get('layer'));now=time.monotonic()
                if key!=self.last_logged[0] or now-self.last_logged[1]>=5:
                    self.append(progress_text(label,info));self.last_logged=(key,now)
            elif kind=='confirm' and not self.cancelled.is_set():
                self.append('Найдены предупреждения:\n'+data)
                self.stage.set('DXF ещё не сохранён. Подтвердите экспорт или вернитесь к проверке.')
                self.bar.stop();self.choice.pack(side='bottom',fill='x',pady=8,before=self.log_area)
            elif kind in ('done','cancelled','error'):
                if kind=='done':
                    message=data.get('check_status','DXF сохранён')
                    self.append(message)
                    self.append(f'Сегментов до оптимизации: {data.get("segments_before", "—")}; после: {data.get("segments_after", "—")}.')
                    for warning in data.get('warnings',[]):self.append('Предупреждение: '+warning)
                    if data.get('backup'):self.append('Резервная копия: '+str(data['backup']))
                    self.append('DXF сохранён: '+str(self.target))
                elif kind=='cancelled':
                    message='Экспорт отменён. DXF не заменён.';self.append(message)
                else:
                    message='Экспорт не завершён. Подробности в журнале.'
                    self.append(data);self.append('Технические подробности: '+str(self.request.parent/'error.txt'))
                self.finish(message)
                if kind=='done':self.bar.configure(value=100)
                return
        seconds=int(time.monotonic()-self.started)
        self.elapsed.set(f'Прошло {seconds//60:02d}:{seconds%60:02d}')
        self.parent.after(100,self.poll)

    def open_log(self):
        if self.log_path:
            try:os.startfile(str(self.log_path))
            except OSError as error:self.stage.set('Не удалось открыть журнал: '+str(error))
