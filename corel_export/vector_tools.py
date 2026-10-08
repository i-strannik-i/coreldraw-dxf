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

    def shape_contours(self, shape):
        return [dict(id=int(shape.StaticID), layer=shape.Layer.Name, closed=bool(p.Closed),
                     segments=decode_subpath(p.GetCurveInfo(),self.scale,p.Segments.Count,bool(p.Closed)))
                for p in shape.DisplayCurve.SubPaths if p.Segments.Count]

    def snapshot(self, selection_only):
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
                if s.Type not in (1,2,3):
                    skipped.append(int(s.StaticID));continue
                if s.PowerClip is not None:
                    skipped.append(int(s.StaticID));continue
                contours.extend(self.shape_contours(s))
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
    root.geometry('750x520');root.minsize(650,460)
    frame=ttk.Frame(root,padding=16);frame.pack(fill='both',expand=True)
    try:session=CorelSession()
    except Exception as error:
        messagebox.showerror('DXF',str(error),parent=root);root.destroy();return
    status=tk.StringVar(value='Исходные векторы не изменяются при проверке.')
    ttk.Label(frame,textvariable=status,wraplength=710).pack(anchor='w',pady=8)
    def safe(action):
        try:action()
        except Exception as error:messagebox.showerror('DXF',str(error),parent=root)
    if '--join' in sys.argv:
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
        across=tk.BooleanVar(value=False)
        ttk.Checkbutton(frame,text='Только выделенные (иначе видимые объекты текущей страницы)',variable=selected).pack(anchor='w')
        ttk.Checkbutton(frame,text='Искать также между разными слоями',variable=across).pack(anchor='w')
        ttk.Label(frame,text='INFO и служебные метки пропускаются. Кривые: приближение 0,005 мм;\nпересечения/касания кривых — кандидаты для визуальной проверки.',wraplength=710).pack(anchor='w',pady=6)
        tree=ttk.Treeview(frame,columns=('type','layer','x','y'),show='headings',height=11)
        for name,title,width in [('type','Проблема',230),('layer','Слой',170),('x','X, мм',100),('y','Y, мм',100)]:
            tree.heading(name,text=title);tree.column(name,width=width)
        tree.pack(fill='both',expand=True)
        messages=queue.Queue();state={'issues':[],'busy':False}
        names={'intersection':'Пересечение/касание','overlap':'Наложение','open':'Открытый конец','zero':'Нулевой участок'}
        def scan():
            if state['busy']:return
            state['issues']=[]; tree.delete(*tree.get_children())
            contours,skipped=session.snapshot(selected.get())
            state['busy']=True;status.set('Проверка…');all_layers=across.get()
            def work():
                try:messages.put((audit(contours,all_layers),skipped,None))
                except Exception as error:messages.put((None,skipped,str(error)))
            threading.Thread(target=work,daemon=True).start()
        def pick(event=None):
            rows=tree.selection()
            if rows:safe(lambda:session.focus(state['issues'][int(rows[0])]))
        tree.bind('<<TreeviewSelect>>',pick)
        def poll():
            if not messages.empty():
                result,skipped,error=messages.get();state['busy']=False
                if error:status.set(error)
                else:
                    state['issues']=result['issues']
                    for i,issue in enumerate(result['issues']):
                        x,y=issue['point'];tree.insert('', 'end',iid=str(i),values=(str(i+1)+'. '+names[issue['kind']]+(' *' if issue['approx'] else ''),issue['layer'],f'{x:.3f}',f'{y:.3f}'))
                    status.set(f'Мест для проверки: {len(result["issues"])}. Пропущено объектов: {len(skipped)}. '+('Выберите строку для приближения.' if result['complete'] else 'Достигнут лимит: проверьте меньше объектов, результат НЕ полный.'))
            root.after(150,poll)
        row=ttk.Frame(frame);row.pack(fill='x',pady=10)
        ttk.Button(row,text='Проверить',command=lambda:safe(scan)).pack(side='left')
        ttk.Button(row,text='Показать метки',command=lambda:safe(lambda:session.mark(state['issues']))).pack(side='left',padx=6)
        ttk.Button(row,text='Убрать мои метки',command=lambda:safe(session.clear_marks)).pack(side='left')
        ttk.Label(frame,text='Метки — отдельный непечатный слой; в DXF не экспортируются. Сохранение CDR — вручную.',wraplength=710).pack(anchor='w')
        poll()
    root.mainloop()


if __name__=='__main__':main()
