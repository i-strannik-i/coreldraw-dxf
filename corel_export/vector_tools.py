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
import time
import uuid

if not getattr(sys, 'frozen', False):
    sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'engine'))
from corel_geometry import decode_subpath
from vector_audit import audit, ambiguous_endpoints

MARK_PREFIX = '_DXF_CHECK_'
ISSUE_STYLES = {
    'zero': ('Нулевой участок', 'Блокирует экспорт', '#c5363d'),
    'overlap': ('Наложение', 'Предупреждение: проверьте визуально', '#7843a3'),
    'intersection': ('Пересечение / касание', 'Предупреждение: проверьте визуально', '#a65300'),
    'open': ('Открытый контур', 'Предупреждение, экспорт разрешён', '#006a9e'),
}


def excluded_layer(name):
    name=name.strip().upper()
    return name.startswith(MARK_PREFIX) or name in {
        'INFO','GUIDES','GUIDELINES','DESKTOP','DOCUMENT GRID','GRID',
        'НАПРАВЛЯЮЩИЕ','РАБОЧИЙ СТОЛ','СЕТКА ДОКУМЕНТА','СЕТКА',
    }


def draw_issue_symbol(canvas, kind, color):
    line='#73818d'
    if kind=='zero':
        canvas.create_line(9,9,36,36,fill=line,width=2)
        canvas.create_rectangle(19,19,27,27,fill=color,outline=color)
        canvas.create_oval(13,13,33,33,outline=color,width=2)
    elif kind=='overlap':
        canvas.create_rectangle(7,8,29,29,outline=line,width=2)
        canvas.create_rectangle(20,20,42,41,outline=line,width=2)
        canvas.create_line(20,29,29,29,29,20,fill=color,width=4)
    elif kind=='intersection':
        canvas.create_line(7,9,41,39,fill=line,width=2)
        canvas.create_line(7,39,41,9,fill=line,width=2)
        canvas.create_oval(17,17,31,31,outline=color,width=3)
        canvas.create_line(24,13,24,35,fill=color,width=1)
        canvas.create_line(13,24,35,24,fill=color,width=1)
    elif kind=='open':
        canvas.create_line(18,9,8,9,8,39,40,39,40,9,30,9,fill=line,width=2)
        for x in (18,30):canvas.create_oval(x-3,6,x+3,12,fill=color,outline=color)
        canvas.create_line(21,9,27,9,fill=color,width=2,dash=(2,2))


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
                if excluded_layer(name) or not s.Layer.Visible:
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

    def marker_layers(self):
        return [l for l in self.page.Layers if re.fullmatch(r'_DXF_CHECK_[0-9a-f]{8}',l.Name)
                and not l.Printable and l.Shapes.Count>0 and all(s.Name=='DXF_CHECK_MARKER_V1' for s in l.Shapes)]

    def clear_marks(self):
        self.current()
        for layer in self.marker_layers():
            self.doc.BeginCommandGroup('Убрать метки проверки DXF')
            try:
                layer.Editable=True
                layer.Delete()
                self.mark_layer=None
            finally:
                self.doc.EndCommandGroup()

    def refresh_marks(self, result, cancelled=None):
        if cancelled and cancelled():raise InterruptedError('Проверка отменена. Прежние метки сохранены.')
        if not result['complete']:return False
        self.mark(result['issues'],cancelled)
        return True

    def mark(self, issues, cancelled=None):
        self.unchanged()
        old_layers=self.marker_layers()
        if not issues and not old_layers:return
        def check_cancel():
            if cancelled and cancelled():raise InterruptedError('Проверка отменена. Прежние метки сохранены.')
        check_cancel()
        self.doc.BeginCommandGroup('Обновить метки проверки DXF')
        changed=False;layer=None
        try:
            if issues:
                layer=self.page.CreateLayer(MARK_PREFIX+uuid.uuid4().hex[:8])
                changed=True
                layer.Printable=False
            colors={kind:tuple(int(style[2][i:i+2],16) for i in (1,3,5)) for kind,style in ISSUE_STYLES.items()}
            for number,issue in enumerate(issues,1):
                check_cancel()
                x,y=(v/self.scale for v in issue['point']); radius=1/self.scale
                circle=layer.CreateEllipse2(x,y,radius,radius)
                circle.Name='DXF_CHECK_MARKER_V1'
                circle.Fill.ApplyNoFill()
                circle.Outline.Width=.15/self.scale
                circle.Outline.Color.RGBAssign(*colors[issue['kind']])
                text=layer.CreateArtisticText(x+radius,y+radius,str(number))
                text.Name='DXF_CHECK_MARKER_V1'
                text.Fill.UniformColor.RGBAssign(*colors[issue['kind']])
            if layer is not None:layer.Editable=False
            check_cancel()
            # Replace only after the new set is ready; one Undo restores the old set.
            for old in old_layers:
                old.Editable=True;changed=True
                old.Delete()
        except Exception:
            self.doc.EndCommandGroup()
            if changed:self.doc.Undo()
            raise
        else:
            self.doc.EndCommandGroup()
            self.mark_layer=layer

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
        if {int(s.StaticID) for s in self.app.ActiveSelectionRange} != set(self.refs):
            raise ValueError('Выделение изменилось. Пересчитайте предпросмотр соединения.')
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
            for layer in self.marker_layers():
                layer.Editable=True
                layer.Delete()
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
    try:
        from .ui_shell import apply_theme, load_theme, save_theme, feedback, THEMES
        from .version import VERSION
    except ImportError:
        from ui_shell import apply_theme, load_theme, save_theme, feedback, THEMES
        from version import VERSION
    root=tk.Tk(); root.title('CorelDRAW → DXF · '+VERSION+' beta')
    root.geometry('790x780');root.minsize(750,720)
    frame=ttk.Frame(root,padding=16);frame.pack(fill='both',expand=True)
    try:session=CorelSession()
    except Exception as error:
        messagebox.showerror('DXF',str(error),parent=root);root.destroy();return
    header=ttk.Frame(frame);header.pack(fill='x',pady=(0,12))
    ttk.Label(header,text='CorelDRAW → DXF',style='Title.TLabel').pack(side='left')
    ttk.Label(header,text=VERSION+' beta · Strannik',style='Muted.TLabel').pack(side='left',padx=14)
    theme_value=tk.StringVar(value='Тёмная' if load_theme()=='dark' else 'Светлая')
    theme_picker=ttk.Combobox(header,textvariable=theme_value,values=('Светлая','Тёмная'),state='readonly',width=9)
    theme_picker.pack(side='right')
    ttk.Label(header,text='Тема').pack(side='right',padx=6)
    notebook=ttk.Notebook(frame);notebook.pack(fill='both',expand=True)
    export_tab=ttk.Frame(notebook,padding=20)
    check_tab=ttk.Frame(notebook,padding=12)
    help_tab=ttk.Frame(notebook,padding=20)
    notebook.add(export_tab,text='Экспорт');notebook.add(check_tab,text='Проверка');notebook.add(help_tab,text='Слои и помощь')
    notebook.select(check_tab if '--check' in sys.argv or '--join' in sys.argv else export_tab)
    if '--help' in sys.argv:notebook.select(help_tab)
    status=tk.StringVar(value='Готово. Проверка не изменяет исходные векторы.')
    ttk.Label(frame,textvariable=status,wraplength=1000,style='Muted.TLabel').pack(side='bottom',anchor='w',pady=(10,0),before=notebook)
    frame=check_tab
    def safe(action):
        try:action()
        except Exception as error:messagebox.showerror('DXF',str(error),parent=root)
    def macro(name):
        session.app.GMSManager.RunMacro('SkladCorelDXF','ExportDXF.'+name)
    ttk.Label(export_tab,text='Сохранить чертёж в DXF',style='Title.TLabel').pack(anchor='w',pady=(4,16))
    export_scope=tk.StringVar(value='selection' if session.app.ActiveSelectionRange.Count else 'page')
    ttk.Radiobutton(export_tab,text='Выделенные объекты',variable=export_scope,value='selection').pack(anchor='w',pady=6)
    ttk.Radiobutton(export_tab,text='Весь текущий лист',variable=export_scope,value='page').pack(anchor='w',pady=6)
    ttk.Separator(export_tab).pack(fill='x',pady=20)
    ttk.Label(export_tab,text='Масштаб 1:1   ·   Миллиметры   ·   Допуск оптимизации 0,1 мм').pack(anchor='w')
    ttk.Label(export_tab,text='DXF сохраняется рядом с CDR под тем же именем.\nСуществующий DXF получит резервную копию.\nПересечения и наложения требуют подтверждения.\nОткрытые контуры не замыкаются автоматически.',style='Muted.TLabel').pack(anchor='w',pady=16)
    export_path=tk.StringVar(value=str(session.doc.FullFileName or 'Несохранённый документ: папку предложим выбрать при экспорте.'))
    ttk.Label(export_tab,textvariable=export_path,wraplength=820).pack(anchor='w',pady=(0,20))
    def begin_export():
        if state['busy']:return
        session.current()
        if export_scope.get()=='selection' and not session.app.ActiveSelectionRange.Count:
            raise ValueError('Сначала выделите объекты в CorelDRAW.')
        status.set('Подготовка экспорта. Ход операции появится в окне журнала.')
        session.app.GMSManager.RunMacro('SkladCorelDXF','ExportDXF.RunExport',export_scope.get()=='selection')
    ttk.Button(export_tab,text='Экспортировать DXF',style='Accent.TButton',command=lambda:safe(begin_export)).pack(anchor='w')
    ttk.Button(export_tab,text='Проверить векторы перед экспортом',command=lambda:notebook.select(check_tab)).pack(anchor='w',pady=12)
    ttk.Label(help_tab,text='Памятка по слоям',style='Title.TLabel').pack(anchor='w',pady=(0,12))
    rows=[('INFO','Информационный текст. Без резки и гравировки.'),('#','Артикулы: гравировка на глубину 0,1 мм.'),
          ('CUT_OUT','Сквозной рез снаружи контура.'),('CUT_IN','Сквозной рез внутри контура.'),('CUT_ON','Рез по линии контура.'),
          ('D / D4','Отверстия. D4: диаметр 4 мм.'),('B','Круглая фреза: паз под сгиб материала.'),
          ('P / P_2.5','Выборка. P_2.5: глубина 2,5 мм.'),('V / V_3','Гравировка. V_3: ширина 3 мм, не глубина.'),
          ('BACK_…','Обратная сторона. Например, BACK_CUT_OUT.')]
    help_table=ttk.Treeview(help_tab,columns=('layer','meaning'),show='headings',height=10)
    help_table.heading('layer',text='Слой');help_table.heading('meaning',text='Назначение')
    help_table.column('layer',width=130,stretch=False);help_table.column('meaning',width=650)
    for row_values in rows:help_table.insert('','end',values=row_values)
    help_table.pack(fill='x')
    ttk.Label(help_tab,text='Названия слоёв — латиницей. Размеры — в мм.\nСоединение: выберите кривые одного слоя → задайте допуск → проверьте «До / После» → примените.\nТема Corel автоматически не считывается: выберите светлую или тёмную тему вверху.',style='Muted.TLabel',wraplength=850).pack(anchor='w',pady=15)
    help_buttons=ttk.Frame(help_tab);help_buttons.pack(fill='x')
    ttk.Button(help_buttons,text='Обратная связь',style='Accent.TButton',command=lambda:feedback(root)).pack(side='left')
    ttk.Button(help_buttons,text='Обновления',command=lambda:safe(lambda:macro('ShowUpdates'))).pack(side='left',padx=8)
    ttk.Button(help_buttons,text='Горячая клавиша',command=lambda:safe(lambda:macro('ShowHotkeys'))).pack(side='left')
    if '--legacy-join' in sys.argv:
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
        root.geometry('1080x800');root.minsize(1000,760)
        body=ttk.Frame(frame);body.pack(fill='both',expand=True,pady=(8,0))
        layer_box=ttk.LabelFrame(body,text='Слои для проверки',padding=10)
        layer_box.pack(side='left',fill='y',padx=(0,16))
        frame=ttk.Frame(body);frame.pack(side='left',fill='both',expand=True)
        selected=tk.BooleanVar(value=bool(session.app.ActiveSelectionRange.Count))
        scope=ttk.Checkbutton(frame,text='Только выделенные (иначе видимые объекты текущей страницы)',variable=selected)
        scope.pack(anchor='w')
        layer_row=ttk.Frame(layer_box);layer_row.pack(fill='x')
        layer_count=tk.StringVar()
        ttk.Label(layer_box,textvariable=layer_count).pack(anchor='w',pady=(10,8))
        ttk.Label(layer_box,text='Технические и скрытые слои,\nINFO и метки исключены.\nВ CDR ничего не удаляется.',wraplength=210).pack(side='bottom',anchor='w',pady=(10,0))
        layer_area=ttk.Frame(layer_box);layer_area.pack(fill='both',expand=True)
        layer_canvas=tk.Canvas(layer_area,width=190,highlightthickness=0)
        layer_scroll=ttk.Scrollbar(layer_area,orient='vertical',command=layer_canvas.yview)
        layer_scroll.pack(side='right',fill='y');layer_canvas.pack(side='left',fill='both',expand=True)
        layer_canvas.configure(yscrollcommand=layer_scroll.set)
        layer_list=ttk.Frame(layer_canvas)
        layer_canvas.create_window((0,0),window=layer_list,anchor='nw')
        layer_list.bind('<Configure>',lambda event:layer_canvas.configure(scrollregion=layer_canvas.bbox('all')))
        def wheel(event):layer_canvas.yview_scroll(-int(event.delta/120),'units')
        layer_canvas.bind('<MouseWheel>',wheel);layer_list.bind('<MouseWheel>',wheel)
        layer_vars={};layer_controls=[]
        for layer in session.page.Layers:
            name=layer.Name
            if not layer.Visible or excluded_layer(name) or layer.Shapes.Count == 0:continue
            if name in layer_vars:continue
            value=tk.BooleanVar(value=True);layer_vars[name]=value
            button=ttk.Checkbutton(layer_list,text=name,variable=value)
            index=len(layer_controls);button.grid(row=index,column=0,sticky='w',padx=(0,8),pady=6)
            button.bind('<MouseWheel>',wheel)
            layer_controls.append(button)
        for caption,value in [('Выбрать все',True),('Снять все',False)]:
            button=ttk.Button(layer_row,text=caption,command=lambda value=value:[v.set(value) for v in layer_vars.values()])
            button.pack(side='left',padx=(0,6));layer_controls.append(button)
        layer_count.set(f'Выбрано: {len(layer_vars)} из {len(layer_vars)}')
        join_box=ttk.LabelFrame(frame,text='Соединение и замыкание · только выделенные кривые одного слоя',padding=10)
        join_box.pack(side='bottom',fill='x',pady=(12,0))
        join_fields=ttk.Frame(join_box);join_fields.pack(fill='x')
        ttk.Label(join_fields,text='Допуск, мм:').pack(side='left')
        join_tolerance=tk.StringVar(value='0,1')
        join_entry=ttk.Entry(join_fields,textvariable=join_tolerance,width=8)
        join_entry.pack(side='left',padx=8)
        join_text=tk.StringVar(value='Выделите кривые и нажмите\n«Пересчитать выделение».')
        ttk.Label(join_fields,textvariable=join_text,wraplength=390).pack(side='left',padx=10)
        settings=ttk.Frame(frame);settings.pack(fill='x',pady=8)
        ttk.Label(settings,text='Допуск проверки кривых, мм:').pack(side='left')
        tolerance=tk.StringVar(value='0,1')
        tolerance_entry=ttk.Entry(settings,textvariable=tolerance,width=10)
        tolerance_entry.pack(side='left',padx=8)
        ttk.Label(frame,text='Проверка по слоям. Пересечения и касания требуют визуального подтверждения.',style='Muted.TLabel',wraplength=710).pack(anchor='w',pady=6)
        legend=ttk.Frame(frame)
        legend.pack(fill='x',pady=(2,10))
        filter_buttons={}
        filter_names={'zero':'Нулевые','overlap':'Наложения','intersection':'Пересечения','open':'Открытые'}
        for index,kind in enumerate(filter_names):
            control=ttk.Button(legend,text=filter_names[kind]+' · 0',style=kind+'.TButton',command=lambda kind=kind:filter_issues(kind))
            control.grid(row=0,column=index,sticky='ew',padx=(0,6));legend.columnconfigure(index,weight=1)
            filter_buttons[kind]=control
        progress=ttk.Progressbar(frame,maximum=100)
        progress.pack(fill='x',pady=(0,8))
        tree=ttk.Treeview(frame,columns=('type','layer','x','y'),show='headings',height=6)
        for name,title,width in [('type','Проблема',230),('layer','Слой',170),('x','X, мм',100),('y','Y, мм',100)]:
            tree.heading(name,text=title);tree.column(name,width=width)
        tree.pack(fill='both',expand=True)
        for kind,style in ISSUE_STYLES.items():tree.tag_configure(kind,foreground=style[2])
        messages=queue.Queue();commands=queue.Queue();state={'issues':[],'busy':False}
        state['filter']=None
        def render_issues():
            tree.delete(*tree.get_children())
            for kind,control in filter_buttons.items():
                count=sum(issue['kind']==kind for issue in state['issues'])
                control.configure(text=filter_names[kind]+f' · {count}'+(' ✓' if state['filter']==kind else ''))
            for i,issue in enumerate(state['issues']):
                if state['filter'] and issue['kind']!=state['filter']:continue
                x,y=issue['point']
                tree.insert('','end',iid=str(i),tags=(issue['kind'],),values=(str(i+1)+'. '+names[issue['kind']]+(' *' if issue['approx'] else ''),issue['layer'],f'{x:.3f}',f'{y:.3f}'))
        def filter_issues(kind):
            if state['busy']:return
            state['filter']=None if state['filter']==kind else kind
            render_issues()
        cancel_event=threading.Event()
        def background():
            import pythoncom
            pythoncom.CoInitialize()
            try:
                # COM references stay in their owning apartment for all operations.
                worker_session=CorelSession()
                join_session=CorelSession()
                while True:
                    command,data=commands.get()
                    if command=='stop':return
                    try:
                        if command=='join_preview':
                            revision,value=data
                            try:
                                before,after=join_session.prepare_join(value)
                                messages.put(('join_preview',(revision,before,after,None)))
                            except Exception as error:
                                messages.put(('join_preview',(revision,None,None,str(error))))
                        elif command=='join_apply':
                            join_session.apply_join()
                            messages.put(('join_applied',None))
                        elif command=='scan':
                            value,only_selected,checked_layers,update_marks=data
                            def reading(layer,count):
                                if cancel_event.is_set():raise InterruptedError('Проверка отменена.')
                                if count%20==0:messages.put(('reading',(layer,count)))
                            contours,skipped=worker_session.snapshot(only_selected,reading,checked_layers)
                            def report(*args):messages.put(('progress',args))
                            result=audit(contours,tolerance=value,progress=report,cancelled=cancel_event.is_set)
                            messages.put(('marking',None))
                            result['marks_updated']=worker_session.refresh_marks(result,cancel_event.is_set) if update_marks else False
                            result['after_join']=not update_marks
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
        def scan(update_marks=True):
            if state['busy']:return
            try:value=float(tolerance.get().replace(',','.'))
            except ValueError:raise ValueError('Введите допуск числом, например 0,1 мм.')
            if not math.isfinite(value) or not .001<=value<=1:
                raise ValueError('Допуск проверки: от 0,001 до 1 мм.')
            checked_layers={name for name,var in layer_vars.items() if var.get()}
            if not checked_layers:raise ValueError('Отметьте хотя бы один слой для проверки.')
            state['issues']=[]; render_issues()
            state['busy']=True;cancel_event.clear()
            join_apply.configure(state='disabled');join_refresh.configure(state='disabled');join_entry.configure(state='disabled')
            for control in (check,marks,clear,scope,tolerance_entry,*layer_controls):control.configure(state='disabled')
            status.set('Чтение векторов из CorelDRAW…');progress.configure(mode='indeterminate');progress.start()
            cancel.configure(state='normal')
            commands.put(('scan',(value,selected.get(),checked_layers,update_marks)))
        def finish():
            state['busy']=False;progress.stop();progress.configure(mode='determinate')
            for control in (check,clear,scope,tolerance_entry,*layer_controls):control.configure(state='normal')
            marks.configure(state='normal' if state['issues'] else 'disabled')
            cancel.configure(state='disabled')
            join_refresh.configure(state='normal');join_entry.configure(state='normal')
        def action(command,data=None):
            if state['busy']:return
            state['busy']=True
            join_apply.configure(state='disabled');join_refresh.configure(state='disabled');join_entry.configure(state='disabled')
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
                elif kind=='marking':
                    status.set('Обновление меток на чертеже… Прежние метки заменяются, а не накапливаются.')
                elif kind=='action_done':
                    finish();status.set({'focus':'Место выделено и приближено в CorelDRAW.','mark':'Метки добавлены на отдельный непечатный слой.','clear':'Служебные метки удалены.'}[data])
                elif kind=='error':
                    finish();status.set(data);progress['value']=0
                    join_apply.configure(state='disabled')
                elif kind=='join_preview':
                    revision,before,after,error=data
                    finish()
                    if revision!=join_state['revision']:
                        schedule_join()
                    elif error:
                        join_text.set(error)
                    else:
                        join_text.set(f'До: замкнутых {before[0]},\nоткрытых {before[1]}\nПосле: замкнутых {after[0]},\nоткрытых {after[1]}')
                        join_state['ready']=revision
                        join_apply.configure(state='normal' if before!=after else 'disabled')
                        status.set('Предпросмотр готов. Векторы не изменены. Кнопка применения подтверждает соединение.')
                elif kind=='join_applied':
                    finish();join_state['ready']=None
                    join_text.set('Соединение применено.\nОтмена: Ctrl+Z в CorelDRAW.')
                    join_apply.configure(state='disabled')
                    safe(lambda:scan(False))
                elif kind=='result':
                    result,skipped,count,value=data
                    state['issues']=result['issues']
                    render_issues()
                    finish();progress['value']=100 if result['complete'] else 0
                    status.set(f'Проверено контуров: {count}. Допуск: {value:g} мм. Найдено мест: {len(result["issues"])}. Пропущено объектов: {len(skipped)}. '+(('Выберите строку — Corel выделит и приблизит место.' if result['issues'] else 'По проверенным критериям проблем не найдено.') if result['complete'] else 'Достигнут лимит: результат НЕ полный.'))
                    status.set(status.get()+(' После соединения: «Показать метки» добавит метки. Ctrl+Z отменяет соединение.' if result.get('after_join') else (' Метки обновлены.' if result['marks_updated'] else ' Прежние метки сохранены.')))
            root.after(150,poll)
        row=ttk.Frame(frame);row.pack(side='bottom',fill='x',pady=10,before=tree)
        check=ttk.Button(row,text='Проверить',style='Accent.TButton',command=lambda:safe(scan));check.pack(side='left')
        marks=ttk.Button(row,text='Показать метки',command=lambda:action('mark',state['issues']),state='disabled');marks.pack(side='left',padx=6)
        clear=ttk.Button(row,text='Убрать мои метки',command=lambda:action('clear'));clear.pack(side='left')
        cancel=ttk.Button(row,text='Отменить проверку',command=cancel_event.set,state='disabled');cancel.pack(side='right')
        join_state={'revision':0,'ready':None,'timer':None}
        def preview_join():
            join_state['timer']=None
            if state['busy']:
                join_state['timer']=root.after(400,preview_join)
                return
            try:
                value=float(join_tolerance.get().replace(',','.'))
                if not math.isfinite(value) or not 0<value<=1:raise ValueError()
            except ValueError:
                join_text.set('Введите допуск больше 0\nи не больше 1 мм.');return
            join_text.set('Пересчитываю…')
            action('join_preview',(join_state['revision'],value))
        def schedule_join(*args):
            join_state['revision']+=1;join_state['ready']=None
            join_apply.configure(state='disabled')
            if join_state['timer'] is not None:root.after_cancel(join_state['timer'])
            join_state['timer']=root.after(400,preview_join)
        def apply_join_preview():
            if state['busy'] or join_state['ready']!=join_state['revision']:return
            action('join_apply')
        join_refresh=ttk.Button(join_box,text='Пересчитать выделение',command=schedule_join)
        join_refresh.pack(side='left',pady=(8,0))
        join_apply=ttk.Button(join_box,text='Применить соединение и замыкание',command=apply_join_preview,state='disabled')
        join_apply.pack(side='right',pady=(8,0))
        join_tolerance.trace_add('write',schedule_join)
        def invalidate(*args):
            if state['busy']:return
            layer_count.set(f'Выбрано: {sum(v.get() for v in layer_vars.values())} из {len(layer_vars)}')
            state['issues']=[];render_issues();marks.configure(state='disabled')
            progress['value']=0;status.set('Параметры изменены. Нажмите «Проверить».')
        for variable in (*layer_vars.values(),tolerance,selected):variable.trace_add('write',invalidate)
        root.protocol('WM_DELETE_WINDOW',lambda:(cancel_event.set(),commands.put(('stop',None)),root.destroy()))
        ttk.Label(frame,text='Метки не печатаются и не входят в DXF. Сохранение CDR — вручную.',style='Muted.TLabel',wraplength=710).pack(side='bottom',anchor='w',before=tree)
        poll()
        def refresh_theme(event=None):
            name='dark' if theme_value.get()=='Тёмная' else 'light'
            apply_theme(root,name,tree,layer_canvas)
            try:save_theme(name)
            except OSError:status.set('Не удалось сохранить тему. В текущем окне она применена.')
        theme_picker.bind('<<ComboboxSelected>>',refresh_theme)
        refresh_theme()
        last_request=[0]
        def switch_requested_tab():
            try:
                request=Path(os.environ['LOCALAPPDATA'])/'SkladCorelDXF/open-tab.json'
                stamp=request.stat().st_mtime_ns
                if stamp!=last_request[0]:
                    last_request[0]=stamp
                    data=json.loads(request.read_text(encoding='utf-8'))
                    tab={'export':export_tab,'check':check_tab,'help':help_tab}.get(data.get('tab'))
                    if tab is not None and abs(time.time()-data.get('created',0))<10:
                        notebook.select(tab);root.deiconify();root.lift()
            except (OSError,ValueError,TypeError):pass
            root.after(350,switch_requested_tab)
        switch_requested_tab()
    root.mainloop()


if __name__=='__main__':
    import win32api
    import win32event
    import winerror
    mutex=win32event.CreateMutex(None,False,'Local\\SkladCorelDXFMainWindow')
    try:
        if win32api.GetLastError()==winerror.ERROR_ALREADY_EXISTS:
            request=Path(os.environ['LOCALAPPDATA'])/'SkladCorelDXF/open-tab.json'
            request.parent.mkdir(parents=True,exist_ok=True)
            mode='check' if '--check' in sys.argv or '--join' in sys.argv else ('help' if '--help' in sys.argv else 'export')
            temp=request.with_name('open-tab-'+str(os.getpid())+'.tmp')
            temp.write_text(json.dumps({'tab':mode,'created':time.time()}),encoding='utf-8')
            os.replace(temp,request)
        else:
            main()
    finally:
        win32api.CloseHandle(mutex)
