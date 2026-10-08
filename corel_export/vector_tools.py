"""Corel vector inspection and explicit, undoable joining; no silent repairs."""
import hashlib
import json
import math
import os
from pathlib import Path
import queue
import re
import sys
import threading
import uuid

if not getattr(sys, 'frozen', False):
    sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'engine'))
from corel_geometry import decode_subpath
from vector_audit import audit, ambiguous_endpoints

MARK_PREFIX = '_DXF_CHECK_'


class CorelSession:
    def __init__(self):
        import win32com.client
        self.app = win32com.client.gencache.EnsureDispatch('CorelDRAW.Application.25')
        if not self.app.Documents.Count:
            raise ValueError('Откройте документ CorelDRAW.')
        self.doc = self.app.ActiveDocument
        self.page = self.doc.ActivePage
        self.scale = self.app.ConvertUnits(1, self.doc.Unit, 3)
        self.refs = {}
        self.contours = []
        self.mark_layer = None

    def current(self):
        if not self.app.Documents.Count or self.app.ActiveDocument._oleobj_ != self.doc._oleobj_ or self.doc.ActivePage.Index != self.page.Index:
            raise ValueError('Документ или страница изменились. Откройте проверку заново.')

    def shape_contours(self, shape, progress=None):
        sid,layer=int(shape.StaticID),shape.Layer.Name
        result=[]
        for number,p in enumerate(shape.DisplayCurve.SubPaths):
            if progress:progress(layer,number)
            count=p.Segments.Count
            if count:
                closed=bool(p.Closed)
                result.append(dict(id=sid,layer=layer,closed=closed,
                    segments=decode_subpath(p.GetCurveInfo(),self.scale,count,closed)))
        return result

    def snapshot(self, selection_only, progress=None, layers=None):
        self.current()
        shapes = list(self.app.ActiveSelectionRange) if selection_only else list(self.page.Shapes)
        if selection_only and not shapes:
            raise ValueError('Выделите объекты в CorelDRAW.')
        refs, contours, skipped = {}, [], []
        def visit(items):
            for s in items:
                name = s.Layer.Name
                if name.startswith(MARK_PREFIX) or name.strip().upper()=='INFO' or not s.Layer.Visible:
                    continue
                if s.Type==7:
                    visit(list(s.Shapes));continue
                if layers is not None and name not in layers:
                    continue
                if progress:
                    progress(name,len(refs))
                if s.Type not in (1,2,3):
                    skipped.append(int(s.StaticID));continue
                if s.PowerClip is not None:
                    skipped.append(int(s.StaticID));continue
                contours.extend(self.shape_contours(s,progress))
                refs[int(s.StaticID)]=s
        visit(shapes)
        self.refs, self.contours = refs, contours
        self.signature = self.digest(contours)
        return contours, skipped

    @staticmethod
    def digest(contours):
        return hashlib.sha256(json.dumps(contours,sort_keys=True).encode()).hexdigest()

    def unchanged(self):
        self.current()
        try:
            now=[c for s in self.refs.values() for c in self.shape_contours(s)]
        except Exception as error:
            raise ValueError('Объекты изменились. Повторите проверку.') from error
        if self.digest(now)!=self.signature:
            raise ValueError('Геометрия или слои изменились. Повторите проверку.')

    def focus(self, issue):
        self.unchanged()
        selected=self.app.CreateShapeRange()
        for sid in issue['ids']:selected.Add(self.refs[sid])
        selected.CreateSelection()
        x,y=issue['point']; width=30/self.scale
        self.doc.ActiveWindow.ActiveView.SetViewArea(x/self.scale-width/2,y/self.scale-width/2,width,width)

    def clear_marks(self):
        self.current()
        layers=[l for l in self.page.Layers if re.fullmatch(r'_DXF_CHECK_[0-9a-f]{8}',l.Name)
                and not l.Printable and l.Shapes.Count>0 and all(s.Name=='DXF_CHECK_MARKER_V1' for s in l.Shapes)]
        for layer in layers:
            self.doc.BeginCommandGroup('Убрать метки проверки DXF')
            try:
                layer.Editable=True
                layer.Delete()
                self.mark_layer=None
            finally:
                self.doc.EndCommandGroup()

    def mark(self, issues):
        self.unchanged()
        self.clear_marks()
        if not issues:return
        self.doc.BeginCommandGroup('Метки проверки DXF')
        try:
            layer=self.page.CreateLayer(MARK_PREFIX+uuid.uuid4().hex[:8])
            self.mark_layer=layer
            layer.Printable=False
            colors={'intersection':(239,55,60),'overlap':(170,40,220),'open':(245,130,20),'zero':(0,160,205)}
            for number,issue in enumerate(issues,1):
                x,y=(v/self.scale for v in issue['point']); radius=1/self.scale
                circle=layer.CreateEllipse2(x,y,radius,radius)
                circle.Name='DXF_CHECK_MARKER_V1'
                circle.Fill.ApplyNoFill()
                circle.Outline.Width=.15/self.scale
                circle.Outline.Color.RGBAssign(*colors[issue['kind']])
                text=layer.CreateArtisticText(x+radius,y+radius,str(number))
                text.Name='DXF_CHECK_MARKER_V1'
                text.Fill.UniformColor.RGBAssign(*colors[issue['kind']])
            layer.Editable=False
        finally:
            self.doc.EndCommandGroup()

    def prepare_join(self,tolerance):
        if not math.isfinite(tolerance) or not 0<tolerance<=1:
            raise ValueError('Допуск должен быть больше 0 и не больше 1 мм.')
        self.snapshot(True)
        if not self.refs:raise ValueError('Нет выбранных кривых.')
        # Combining across layers or appearances could erase production semantics.
        styles=set()
        for s in self.refs.values():
            if s.Type!=3 or s.Fill.Type!=0:
                raise ValueError('Для соединения выделите кривые без заливки. Прямоугольники и текст сначала преобразуйте в кривые вручную.')
            color=s.Outline.Color.GetCopy();color.ConvertToRGB()
            styles.add((s.Layer.Name,s.Outline.Type,s.Outline.Width,color.RGBRed,color.RGBGreen,color.RGBBlue))
        if len(styles)!=1:raise ValueError('Соединяйте по одному слою, с одинаковыми цветом и толщиной линии.')
        if ambiguous_endpoints(self.contours,tolerance):
            raise ValueError('У некоторых концов больше одного соседа в допуске. Соединение неоднозначно: выделите меньше объектов или уменьшите допуск.')
        curve=self.app.CreateCurve(self.doc)
        for s in self.refs.values():curve.AppendCurve(s.Curve)
        before=(sum(c['closed'] for c in self.contours),sum(not c['closed'] for c in self.contours))
        curve.JoinTouchingSubpaths(True,tolerance/self.scale)
        after=(sum(bool(p.Closed) for p in curve.SubPaths),sum(not p.Closed for p in curve.SubPaths))
        self.joined=curve
        return before,after

    def apply_join(self):
        self.unchanged()
        # Undo includes both the replacement and removal of the selected originals.
        self.doc.BeginCommandGroup('Соединить контуры DXF')
        changed=False
        try:
            result=next(iter(self.refs.values())).Duplicate(0,0)
            changed=True
            result.Curve.CopyAssign(self.joined)
            originals=self.app.CreateShapeRange()
            for s in self.refs.values():originals.Add(s)
            originals.Delete()
            result.CreateSelection()
        except Exception:
            self.doc.EndCommandGroup()
            if changed:self.doc.Undo()
            raise
        else:
            self.doc.EndCommandGroup()
        self.refs={}


def main():
    import tkinter as tk
    from tkinter import ttk,messagebox
    root=tk.Tk(); root.title('DXF: соединение контуров' if '--join' in sys.argv else 'DXF: проверка векторов')
    root.geometry('790x780');root.minsize(750,720)
    frame=ttk.Frame(root,padding=16);frame.pack(fill='both',expand=True)
    try:session=CorelSession()
    except Exception as error:
        messagebox.showerror('DXF',str(error),parent=root);root.destroy();return
    status=tk.StringVar(value='Выберите допуск и нажмите «Проверить». Чертёж не изменяется.')
    ttk.Label(frame,textvariable=status,wraplength=710).pack(anchor='w',pady=8)
    def safe(action):
        try:action()
        except Exception as error:messagebox.showerror('DXF',str(error),parent=root)
    if '--join' in sys.argv:
        status.set('Исходные векторы не изменяются при предпросмотре.')
        root.minsize(630,350)
        root.geometry('630x350')
        ttk.Label(frame,text='Только выделенные кривые. Предпросмотр не меняет чертёж.').pack(anchor='w',pady=10)
        tolerance=tk.StringVar(value='0,1')
        ttk.Label(frame,text='Допуск соединения, мм:').pack(anchor='w')
        ttk.Entry(frame,textvariable=tolerance,width=12).pack(anchor='w',pady=6)
        preview_text=tk.StringVar(value='Нажмите «Предпросмотр».')
        ttk.Label(frame,textvariable=preview_text).pack(anchor='w',pady=12)
        def preview():
            apply.configure(state='disabled')
            before,after=session.prepare_join(float(tolerance.get().replace(',','.')))
            preview_text.set(f'До: замкнутых {before[0]}, открытых {before[1]}\nПосле: замкнутых {after[0]}, открытых {after[1]}')
            state['tolerance']=tolerance.get()
            if before!=after:apply.configure(state='normal')
            else:status.set('Соединений в заданном допуске не найдено. Чертёж не изменён.')
        def join():
            if tolerance.get()!=state.get('tolerance'):raise ValueError('Допуск изменён. Повторите предпросмотр.')
            if messagebox.askyesno('Соединить контуры','Заменить выбранные кривые результатом соединения?\nОтмена изменения: Ctrl+Z в Corel.',parent=root):
                session.apply_join();apply.configure(state='disabled');status.set('Соединено. Чертёж не сохранён автоматически. Отмена: Ctrl+Z.')
        state={}
        ttk.Button(frame,text='Предпросмотр',command=lambda:safe(preview)).pack(side='left',pady=14)
        apply=ttk.Button(frame,text='Соединить',command=lambda:safe(join),state='disabled');apply.pack(side='left',padx=8)
    else:
        selected=tk.BooleanVar(value=bool(session.app.ActiveSelectionRange.Count))
        scope=ttk.Checkbutton(frame,text='Только выделенные (иначе видимые объекты текущей страницы)',variable=selected)
        scope.pack(anchor='w')
        layer_box=ttk.LabelFrame(frame,text='Какие слои проверять',padding=6)
        layer_box.pack(fill='x',pady=6)
        layer_row=ttk.Frame(layer_box);layer_row.pack(fill='x')
        layer_canvas=tk.Canvas(layer_box,height=85,highlightthickness=0)
        layer_scroll=ttk.Scrollbar(layer_box,orient='vertical',command=layer_canvas.yview)
        layer_scroll.pack(side='right',fill='y');layer_canvas.pack(fill='x',expand=True)
        layer_canvas.configure(yscrollcommand=layer_scroll.set)
        layer_list=ttk.Frame(layer_canvas)
        layer_canvas.create_window((0,0),window=layer_list,anchor='nw')
        layer_list.bind('<Configure>',lambda event:layer_canvas.configure(scrollregion=layer_canvas.bbox('all')))
        layer_vars={};layer_controls=[]
        for layer in session.page.Layers:
            name=layer.Name
            if not layer.Visible or name.startswith(MARK_PREFIX) or name.strip().upper()=='INFO':continue
            if name in layer_vars:continue
            value=tk.BooleanVar(value=True);layer_vars[name]=value
            button=ttk.Checkbutton(layer_list,text=name,variable=value)
            index=len(layer_controls);button.grid(row=index//2,column=index%2,sticky='w',padx=(0,24),pady=2)
            layer_controls.append(button)
        for caption,value in [('Выбрать все',True),('Снять все',False)]:
            button=ttk.Button(layer_row,text=caption,command=lambda value=value:[v.set(value) for v in layer_vars.values()])
            button.pack(side='left',padx=(0,6));layer_controls.append(button)
        ttk.Label(layer_row,text='Скрытые слои, INFO и метки исключены.').pack(side='left',padx=8)
        settings=ttk.Frame(frame);settings.pack(fill='x',pady=8)
        ttk.Label(settings,text='Допуск проверки кривых, мм:').pack(side='left')
        tolerance=tk.StringVar(value='0,1')
        tolerance_entry=ttk.Entry(settings,textvariable=tolerance,width=10)
        tolerance_entry.pack(side='left',padx=8)
        ttk.Label(frame,text='Каждый слой проверяется отдельно. Между слоями пересечения не ищем.\nДопуск задаёт точность кривых: меньше — точнее, но дольше.\nОткрытые концы отмечаются независимо от допуска; автоматического замыкания нет.\nINFO и служебные метки пропускаются. Найденные места нужно проверить визуально.',wraplength=710).pack(anchor='w',pady=6)
        progress=ttk.Progressbar(frame,maximum=100)
        progress.pack(fill='x',pady=(0,8))
        tree=ttk.Treeview(frame,columns=('type','layer','x','y'),show='headings',height=11)
        for name,title,width in [('type','Проблема',230),('layer','Слой',170),('x','X, мм',100),('y','Y, мм',100)]:
            tree.heading(name,text=title);tree.column(name,width=width)
        tree.pack(fill='both',expand=True)
        messages=queue.Queue();commands=queue.Queue();state={'issues':[],'busy':False}
        cancel_event=threading.Event()
        def background():
            import pythoncom
            pythoncom.CoInitialize()
            try:
                # COM references stay in their owning apartment for all operations.
                worker_session=CorelSession()
                while True:
                    command,data=commands.get()
                    if command=='stop':return
                    try:
                        if command=='scan':
                            value,only_selected,checked_layers=data
                            def reading(layer,count):
                                if cancel_event.is_set():raise InterruptedError('Проверка отменена.')
                                if count%20==0:messages.put(('reading',(layer,count)))
                            contours,skipped=worker_session.snapshot(only_selected,reading,checked_layers)
                            def report(*args):messages.put(('progress',args))
                            result=audit(contours,tolerance=value,progress=report,cancelled=cancel_event.is_set)
                            messages.put(('result',(result,skipped,len(contours),value)))
                        else:
                            if command=='focus':worker_session.focus(data)
                            elif command=='mark':worker_session.mark(data)
                            elif command=='clear':worker_session.clear_marks()
                            messages.put(('action_done',command))
                    except Exception as error:messages.put(('error',str(error)))
            except Exception as error:messages.put(('error',str(error)))
            finally:pythoncom.CoUninitialize()
        threading.Thread(target=background,daemon=True).start()
        names={'intersection':'Пересечение/касание','overlap':'Наложение','open':'Открытый конец','zero':'Нулевой участок'}
        def scan():
            if state['busy']:return
            try:value=float(tolerance.get().replace(',','.'))
            except ValueError:raise ValueError('Введите допуск числом, например 0,1 мм.')
            if not math.isfinite(value) or not .001<=value<=1:
                raise ValueError('Допуск проверки: от 0,001 до 1 мм.')
            checked_layers={name for name,var in layer_vars.items() if var.get()}
            if not checked_layers:raise ValueError('Отметьте хотя бы один слой для проверки.')
            state['issues']=[]; tree.delete(*tree.get_children())
            state['busy']=True;cancel_event.clear()
            for control in (check,marks,clear,scope,tolerance_entry,*layer_controls):control.configure(state='disabled')
            status.set('Чтение векторов из CorelDRAW…');progress.configure(mode='indeterminate');progress.start()
            cancel.configure(state='normal')
            commands.put(('scan',(value,selected.get(),checked_layers)))
        def finish():
            state['busy']=False;progress.stop();progress.configure(mode='determinate')
            for control in (check,clear,scope,tolerance_entry,*layer_controls):control.configure(state='normal')
            marks.configure(state='normal' if state['issues'] else 'disabled')
            cancel.configure(state='disabled')
        def action(command,data=None):
            if state['busy']:return
            state['busy']=True
            for control in (check,marks,clear,scope,tolerance_entry,*layer_controls):control.configure(state='disabled')
            status.set('Выполняется действие в CorelDRAW…')
            commands.put((command,data))
        def pick(event=None):
            rows=tree.selection()
            if rows:action('focus',state['issues'][int(rows[0])])
        tree.bind('<<TreeviewSelect>>',pick)
        def poll():
            while not messages.empty():
                kind,data=messages.get()
                if kind=='progress':
                    progress.stop();progress.configure(mode='determinate')
                    layer,index,total,stage,done,count=data
                    fraction=done/max(1,count)
                    portion=(.3*fraction if stage=='Подготовка кривых' else .3+.7*fraction)
                    progress['value']=100*(index-1+portion)/max(1,total)
                    status.set(f'Слой {index}/{total}: «{layer}». {stage}: {done}/{count}.')
                elif kind=='reading':
                    layer,count=data
                    status.set(f'Чтение CorelDRAW: слой «{layer}», контур/объект {count+1}…')
                elif kind=='action_done':
                    finish();status.set({'focus':'Место выделено и приближено в CorelDRAW.','mark':'Метки добавлены на отдельный непечатный слой.','clear':'Служебные метки удалены.'}[data])
                elif kind=='error':
                    finish();status.set(data);progress['value']=0
                elif kind=='result':
                    result,skipped,count,value=data
                    if cancel_event.is_set():
                        finish();status.set('Проверка отменена. Частичные результаты не показаны.');continue
                    state['issues']=result['issues']
                    for i,issue in enumerate(result['issues']):
                        x,y=issue['point'];tree.insert('', 'end',iid=str(i),values=(str(i+1)+'. '+names[issue['kind']]+(' *' if issue['approx'] else ''),issue['layer'],f'{x:.3f}',f'{y:.3f}'))
                    finish();progress['value']=100 if result['complete'] else 0
                    status.set(f'Проверено контуров: {count}. Допуск: {value:g} мм. Найдено мест: {len(result["issues"])}. Пропущено объектов: {len(skipped)}. '+(('Выберите строку — Corel выделит и приблизит место.' if result['issues'] else 'По проверенным критериям проблем не найдено.') if result['complete'] else 'Достигнут лимит: результат НЕ полный.'))
            root.after(150,poll)
        row=ttk.Frame(frame);row.pack(fill='x',pady=10)
        check=ttk.Button(row,text='Проверить',command=lambda:safe(scan));check.pack(side='left')
        marks=ttk.Button(row,text='Показать метки',command=lambda:action('mark',state['issues']),state='disabled');marks.pack(side='left',padx=6)
        clear=ttk.Button(row,text='Убрать мои метки',command=lambda:action('clear'));clear.pack(side='left')
        cancel=ttk.Button(row,text='Отменить проверку',command=cancel_event.set,state='disabled');cancel.pack(side='right')
        def invalidate(*args):
            if state['busy']:return
            state['issues']=[];tree.delete(*tree.get_children());marks.configure(state='disabled')
            progress['value']=0;status.set('Параметры изменены. Нажмите «Проверить».')
        for variable in (*layer_vars.values(),tolerance,selected):variable.trace_add('write',invalidate)
        root.protocol('WM_DELETE_WINDOW',lambda:(cancel_event.set(),commands.put(('stop',None)),root.destroy()))
        ttk.Label(frame,text='Метки — отдельный непечатный слой; в DXF не экспортируются. Сохранение CDR — вручную.',wraplength=710).pack(anchor='w')
        poll()
    root.mainloop()


if __name__=='__main__':main()
