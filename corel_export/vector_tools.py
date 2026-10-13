"""Corel inspection with automatic duplicate quarantine and undoable joining."""
import copy
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
from duplicate_geometry import duplicate_groups, duplicate_removals, DUPLICATE_PREFIX
from marker_layout import place_label
from route_order import join_summary
from continuous_vectors import continuous_vectors, continuous_summary, corel_segments, same_vectors
try:
    from .corel_performance import paused_redraw
except ImportError:
    from corel_performance import paused_redraw

MARK_PREFIX = '_DXF_CHECK_'
ISSUE_STYLES = {
    'duplicate': ('Дубли', 'Копии автоматически переносятся на технические слои', '#27733a'),
    'zero': ('Нулевой участок', 'Блокирует экспорт', '#c5363d'),
    'overlap': ('Наложение', 'Предупреждение: проверьте визуально', '#7843a3'),
    'intersection': ('Пересечение / касание', 'Предупреждение: проверьте визуально', '#a65300'),
    'open': ('Открытый контур', 'Предупреждение, экспорт разрешён', '#006a9e'),
}


def check_summary(result, count, skipped, tolerance, selected, layers):
    if not result['complete']:
        outcome='Проверка НЕ завершена. Отсутствие ошибок не подтверждено.'
    elif not count:
        outcome='Нет контуров для проверки. Отсутствие ошибок не подтверждено.'
    elif skipped:
        outcome='Проверка выполнена с пропусками. Не все объекты проверены.'
    elif result['issues']:
        outcome=f'Проверка завершена. Найдено проблем: {len(result["issues"])}.'
    else:
        outcome='Проверка завершена. Ошибок и предупреждений не найдено.'
    scope='Только выделенные объекты' if selected else 'Видимые объекты текущей страницы'
    return (outcome+f'\nКонтуров: {count}; пропущено объектов: {len(skipped)}; допуск: {tolerance:g} мм.'+
            '\n'+scope+'; слои: '+', '.join(sorted(layers))+'.')


def layer_check_summary(part):
    counts={kind:sum(i['kind']==kind for i in part['issues']) for kind in ISSUE_STYLES}
    status=('проверен (результат без изменений использован повторно)' if part.get('reused') else 'проверен') if part['complete'] else 'проверен НЕ полностью (достигнут лимит)'
    return (f"Слой «{part['layer']}» {status}. Контуров: {part['count']}. "
            f"Перенесено дублей: {part.get('duplicates_moved',0)}.\n"
            f"Нулевые: {counts['zero']}; наложения: {counts['overlap']}; "
            f"пересечения/касания: {counts['intersection']}; метки открытых концов: {counts['open']}.")


def excluded_layer(name):
    name=name.strip().upper()
    return name.startswith((MARK_PREFIX,DUPLICATE_PREFIX)) or name in {
        'INFO','#','GUIDES','GUIDELINES','DESKTOP','DOCUMENT GRID','GRID',
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
    elif kind=='duplicate':
        canvas.create_rectangle(10,8,31,29,outline=color,width=2)
        canvas.create_rectangle(17,15,38,36,outline=color,width=2)
        canvas.create_text(24,23,text='=',fill=color)


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
                result.append(dict(id=sid,layer=layer,subpath=number,closed=closed,
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

    def unchanged(self,progress=None):
        self.current()
        try:
            now=[]
            for index,s in enumerate(self.refs.values()):
                if progress:progress('Проверка исходной геометрии',index,len(self.refs))
                now.extend(self.shape_contours(s))
            if progress:progress('Проверка исходной геометрии',len(self.refs),len(self.refs))
        except InterruptedError:raise
        except Exception as error:
            raise ValueError('Объекты изменились. Повторите проверку.') from error
        if self.digest(now)!=self.signature:
            raise ValueError('Геометрия или слои изменились. Повторите проверку.')

    def focus(self, issue):
        # Navigation only depends on the objects participating in this issue.
        # Re-reading every curve here made each click wait for a full snapshot.
        self.current()
        ids=set(issue['ids'])
        try:
            shapes=[self.refs[sid] for sid in sorted(ids)]
            if any(not shape.Layer.Visible for shape in shapes):
                raise ValueError('Слой скрыт.')
            now=[c for shape in shapes for c in self.shape_contours(shape)]
            expected=[c for c in self.contours if c['id'] in ids]
            ordered=lambda contours:sorted(contours,key=lambda c:(c['id'],c.get('subpath',0)))
            if not expected or self.digest(ordered(now))!=self.digest(ordered(expected)):
                raise ValueError('Геометрия изменилась.')
        except Exception as error:
            raise ValueError('Объекты выбранного места изменились или недоступны. Повторите проверку.') from error
        selected=self.app.CreateShapeRange()
        for shape in shapes:selected.Add(shape)
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
        with paused_redraw(self.app):
            self._replace_marks(issues,cancelled)

    def _replace_marks(self, issues, cancelled=None):
        old_layers=self.marker_layers()
        if not issues and not old_layers:return
        selection=[shape for shape in self.app.ActiveSelectionRange
                   if not shape.Layer.Name.startswith(MARK_PREFIX)]
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
            occupied=[]
            marker_boxes=[(i['point'][0]-1,i['point'][1]-1,i['point'][0]+1,i['point'][1]+1) for i in issues]
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
                text.Text.Story.Size=8
                text.Fill.UniformColor.RGBAssign(*colors[issue['kind']])
                width,height=text.SizeWidth*self.scale,text.SizeHeight*self.scale
                box=place_label(issue['point'],issue['kind'],width,height,occupied,marker_boxes)
                occupied.append(box)
                text.Move(box[0]/self.scale-text.LeftX,box[1]/self.scale-text.BottomY)
                px,py=issue['point']
                end_x=min(max(px,box[0]),box[2]);end_y=min(max(py,box[1]),box[3])
                leader=layer.CreateLineSegment(x,y,end_x/self.scale,end_y/self.scale)
                leader.Name='DXF_CHECK_MARKER_V1'
                leader.Outline.Width=.1/self.scale
                leader.Outline.Color.RGBAssign(*colors[issue['kind']])
            if layer is not None:layer.Editable=False
            check_cancel()
            # Replace only after the new set is ready; one Undo restores the old set.
            for old in old_layers:
                old.Editable=True;changed=True
                old.Delete()
            restored=self.app.CreateShapeRange()
            for shape in selection:restored.Add(shape)
            restored.CreateSelection()
        except Exception:
            self.doc.EndCommandGroup()
            if changed:self.doc.Undo()
            raise
        else:
            self.doc.EndCommandGroup()
            self.mark_layer=layer

    def move_duplicates(self, groups, preserve_selection=False,progress=None):
        report=progress or (lambda *args:None)
        self.unchanged(report)
        expected=duplicate_groups(self.contours,lambda done,total:report("Сверка дублей",done,total))
        order=lambda group:(group['layer'],group['keep'])
        if not groups or sorted(groups,key=order)!=sorted(expected,key=order):
            raise ValueError('Список дублей устарел. Повторите полную проверку.')
        copy_ids,partial=duplicate_removals(groups)
        count=sum(len(group['copies']) for group in groups)
        plans={}
        for number,(sid,indexes) in enumerate(partial.items()):
            report("Подготовка переноса частей кривых",number,len(partial))
            shape=self.refs[sid]
            if shape.Type!=3:
                raise ValueError('Для переноса части объекта сначала преобразуйте его в кривые.')
            keep_curve=self.app.CreateCurve(self.doc)
            copy_curve=self.app.CreateCurve(self.doc)
            for index,path in enumerate(shape.Curve.SubPaths):
                (copy_curve if index in indexes else keep_curve).AppendCurve(path.GetCopy())
            if not keep_curve.SubPaths.Count:copy_ids.add(sid)
            else:plans[sid]=(keep_curve,copy_curve)
        copies=[self.refs[sid] for sid in sorted(copy_ids|set(partial))]
        selection=[shape for shape in self.app.ActiveSelectionRange
                   if int(shape.StaticID) not in copy_ids] if preserve_selection else None
        for number,shape in enumerate(copies):
            report("Проверка доступности дублей",number,len(copies))
            if shape.ParentGroup is not None or shape.Locked or not shape.Layer.Editable:
                raise ValueError('Есть дубли в группе или заблокированные объекты. Проверьте их вручную; ничего не перенесено.')
        # Own quarantine events must not cancel the automatic preview midway.
        previous_events=self.app.EventsEnabled
        self.app.EventsEnabled=False
        try:
            self.doc.BeginCommandGroup('Перенести дубли DXF на проверку')
            changed=False
            try:
                layers={}
                for number,shape in enumerate(copies):
                    report("Перенос дублей",number,len(copies))
                    original=shape.Layer.Name
                    if original not in layers:
                        target=self.page.CreateLayer(DUPLICATE_PREFIX+'__'+original+'__'+uuid.uuid4().hex[:8])
                        changed=True;target.Printable=False;layers[original]=target
                    sid=int(shape.StaticID)
                    if sid in plans:
                        keep_curve,copy_curve=plans[sid]
                        isolated=shape.Duplicate(0,0)
                        isolated.Curve.CopyAssign(copy_curve)
                        isolated.MoveToLayer(layers[original])
                        shape.Curve.CopyAssign(keep_curve)
                    else:
                        shape.MoveToLayer(layers[original])
                for layer in self.marker_layers():
                    layer.Editable=True;layer.Delete()
                originals=self.app.CreateShapeRange()
                if selection is not None:
                    for shape in selection:originals.Add(shape)
                else:
                    for group in groups:originals.Add(self.refs[group['keep']])
                originals.CreateSelection()
            except Exception:
                self.doc.EndCommandGroup()
                if changed:self.doc.Undo()
                raise
            else:self.doc.EndCommandGroup()
        finally:
            self.app.EventsEnabled=previous_events
        self.refs={};self.contours=[]
        return count

    def quarantine_duplicates(self, cancelled=None,progress=None):
        """Preserve the original scan scope, even when moving changes selection."""
        self.duplicates_by_layer={}
        def report(stage,done,total):
            if cancelled and cancelled():raise InterruptedError("Проверка отменена.")
            if progress:progress(stage,done,total)
        report("Поиск дублей",0,len(self.contours))
        groups = duplicate_groups(self.contours,lambda done,total:report("Поиск дублей",done,total))
        if cancelled and cancelled():
            raise InterruptedError('Проверка отменена до переноса дублей.')
        if not groups:
            return 0
        copies,partial = duplicate_removals(groups)
        for sid,indexes in partial.items():
            if len(indexes)==sum(c['id']==sid for c in self.contours):copies.add(sid)
        remaining = [c for c in self.contours if c['id'] not in copies]
        refs = {sid:shape for sid,shape in self.refs.items() if sid not in copies}
        count = self.move_duplicates(groups, preserve_selection=True,progress=report)
        for group in groups:
            layer=group["layer"];self.duplicates_by_layer[layer]=self.duplicates_by_layer.get(layer,0)+len(group["copies"])
        if partial:
            remaining=[c for sid,shape in refs.items() for c in
                       (self.shape_contours(shape) if sid in partial else
                        [old for old in remaining if old['id']==sid])]
        self.contours, self.refs = remaining, refs
        self.signature = self.digest(remaining)
        return count

    def join_selection(self,progress=None,selection_only=True):
        """Keep selected containers separate from the curves they contain."""
        report=progress or (lambda *args:None)
        visited=[0]
        self.current()
        roots=list(self.app.ActiveSelectionRange) if selection_only else list(self.page.Shapes)
        leaves={}
        self.join_skipped_by_layer={}
        def skip(shape,reason):
            self.join_skipped_by_layer.setdefault(shape.Layer.Name,[]).append((int(shape.StaticID),reason))
        def visit(shapes):
            for shape in shapes:
                visited[0]+=1;report("Чтение выделенных объектов",visited[0],0)
                layer=shape.Layer
                if excluded_layer(layer.Name) or not layer.Visible:continue
                if not layer.Editable or shape.PowerClip is not None:
                    skip(shape,'заблокирован или PowerClip');continue
                if shape.Type==7:
                    visit(list(shape.Shapes))
                elif shape.Type in (1,2,3) and shape.Fill.Type==0:
                    leaves[int(shape.StaticID)]=shape
                else:
                    skip(shape,'не кривая без заливки')
        visit(roots)
        return roots,leaves

    def join_skipped_summary(self):
        return '\n'.join(f'Слой «{layer}»: пропущено объектов {len(items)} ({", ".join(sorted({reason for sid,reason in items}))}). Они не изменяются.'
                         for layer,items in getattr(self,'join_skipped_by_layer',{}).items())

    def _prepare_join_parts(self,parts,tolerance,optimize_path,report,selection_only):
        outputs=[];delete_refs={};before=[0,0];after=[0,0];summaries=[]
        network=dict(connectors=0,max_connector=0);changed=False;incomplete=False
        for number,(layer,items) in enumerate(parts,1):
            child=copy.copy(self)
            child.contours=items
            ids={c['id'] for c in child.contours}
            child.refs={sid:shape for sid,shape in self.refs.items() if sid in ids}
            child.join_open_refs={sid:shape for sid,shape in self.join_open_refs.items() if sid in ids}
            child.join_source=None;child.join_skipped_by_layer={}
            def layer_progress(stage,done,total):report(f'Слой {number}/{len(parts)} «{layer}»: '+stage,done,total)
            try:
                b,a=child.prepare_join(tolerance,False,optimize_path,layer_progress,selection_only)
            except ValueError as error:
                b=a=(sum(c['closed'] for c in items),sum(not c['closed'] for c in items))
                child.join_changed=False;child.join_incomplete=True
                child.join_network_summary='НЕ рассчитан: '+str(error)+' Исходные объекты сохраняются.'
                child.join_network=dict(connectors=0,max_connector=0)
            incomplete|=child.join_incomplete
            before=[x+y for x,y in zip(before,b)];after=[x+y for x,y in zip(after,a)]
            changed|=child.join_changed
            if child.join_changed:
                outputs.extend(child.join_outputs);delete_refs.update(child.join_delete_refs)
            text=child.join_network_summary or join_summary(b,a,tolerance,child.join_route)
            summaries.append(f'Слой «{layer}»:\n'+text+('\nБез изменений.' if not child.join_changed else ''))
            if optimize_path:
                network['connectors']+=child.join_network['connectors']
                network['max_connector']=max(network['max_connector'],child.join_network['max_connector'])
            report(f'Слой «{layer}» рассчитан',number,len(parts))
        self.join_outputs=outputs;self.join_delete_refs=delete_refs;self.join_changed=changed
        self.join_counts=(tuple(before),tuple(after));self.join_route=None;self.join_warning=''
        self.join_network=network;self.join_layer_reports=summaries;self.join_incomplete=incomplete
        layer_count=len({c['layer'] for label,items in parts for c in items})
        self.join_network_summary=f'Всего: {sum(before)} → {sum(after)} векторов; слоёв: {layer_count}.'+(' НЕ все слои рассчитаны.' if incomplete else '')+'\n\n'+'\n\n'.join(summaries)
        return self.join_counts

    def prepare_join(self,tolerance,refresh=True,optimize_path=False,progress=None,selection_only=True,duplicates_moved=None):
        report=progress or (lambda *args:None)
        report("Чтение выделения CorelDRAW",0,0)
        self.join_optimize_path=optimize_path;self.join_incomplete=False
        if not math.isfinite(tolerance) or not 0<tolerance<=1:
            raise ValueError('Допуск должен быть больше 0 и не больше 1 мм.')
        if refresh or not hasattr(self,'join_source'):
            self.join_source=None
            self.join_selection_only=selection_only
            roots,leaves=self.join_selection(report,selection_only)
            contours=[]
            for number,shape in enumerate(leaves.values()):
                def reading(layer,part):report(f'Чтение геометрии «{layer}», контур {part+1}',number,len(leaves))
                contours.extend(self.shape_contours(shape,reading))
            self.refs=leaves;self.contours=contours;self.signature=self.digest(contours)
            self.join_roots={int(s.StaticID):s for s in roots}
            self.join_leaf_ids=set(leaves)
            open_ids={c['id'] for c in self.contours if not c['closed']}
            self.join_open_refs={sid:s for sid,s in self.refs.items() if sid in open_ids}
        else:
            self.current()
        if not self.refs:raise ValueError('Нет доступных рабочих контуров в выбранной области. '+self.join_skipped_summary())
        # Duplicates must be removed before any layer/style optimization.
        # Use the same transactional quarantine as the shared check pipeline.
        self.join_changed=False
        moved=self.quarantine_duplicates(progress=report)
        if moved:
            if duplicates_moved:duplicates_moved(dict(self.duplicates_by_layer))
            self.join_source=None
            roots,leaves=self.join_selection(report,selection_only)
            self.join_roots={int(s.StaticID):s for s in roots}
            self.join_leaf_ids=set(leaves)
            open_ids={c['id'] for c in self.contours if not c['closed']}
            self.join_open_refs={sid:s for sid,s in self.refs.items() if sid in open_ids}
        if duplicate_groups(self.contours):
            raise ValueError('Остались дубли. Оптимизация остановлена до их переноса.')
        layer_names=list(dict.fromkeys(c['layer'] for c in self.contours))
        if len(layer_names)>1:
            parts=[(layer,[c for c in self.contours if c['layer']==layer]) for layer in layer_names]
            return self._prepare_join_parts(parts,tolerance,optimize_path,report,selection_only)
        self.join_layer_reports=[]
        # Combining across layers or appearances could erase production semantics.
        styles=set();style_groups={}
        style_refs=list((self.refs if optimize_path else self.join_open_refs).values())
        for index,s in enumerate(style_refs):
            report("Чтение оформления",index,len(style_refs))
            if s.Type not in ((1,2,3) if optimize_path else (3,)) or s.Fill.Type!=0:
                raise ValueError('Для соединения выделите кривые без заливки. Прямоугольники и текст сначала преобразуйте в кривые вручную.')
            color=s.Outline.Color.GetCopy();color.ConvertToRGB()
            style=(s.Layer.Name,s.Outline.Type,s.Outline.Width,color.RGBRed,color.RGBGreen,color.RGBBlue)
            styles.add(style);style_groups.setdefault(style,[]).append(s)
        if len(styles)>1 and not optimize_path:
            parts=[]
            for index,sources in enumerate(style_groups.values(),1):
                ids={int(shape.StaticID) for shape in sources}
                parts.append((f'{layer_names[0]}, оформление {index}',[c for c in self.contours if c['id'] in ids]))
            assigned={c['id'] for label,items in parts for c in items}
            retained=[c for c in self.contours if c['id'] not in assigned]
            if retained:parts.append((f'{layer_names[0]}, остальные контуры',retained))
            return self._prepare_join_parts(parts,tolerance,optimize_path,report,selection_only)
        self.join_network_summary=None
        if optimize_path:
            plan=dict(routes=[],connectors=0,max_connector=0)
            route_sources=[]
            for sources in style_groups.values():
                ids={int(s.StaticID) for s in sources}
                part=continuous_vectors([c for c in self.contours if c['id'] in ids],tolerance,self.scale,progress=report)
                plan['routes'].extend(part['routes'])
                plan['connectors']+=part['connectors']
                plan['max_connector']=max(plan['max_connector'],part['max_connector'])
                route_sources.extend([sources[0]]*len(part['routes']))
            before=(sum(c['closed'] for c in self.contours),sum(not c['closed'] for c in self.contours))
            after=(sum(r['closed'] for r in plan['routes']),sum(not r['closed'] for r in plan['routes']))
            outputs=[]
            for source,route in zip(route_sources,plan['routes']):
                route['segments']=corel_segments(route['segments'],self.scale)
                outputs.append((source,route))
            self.join_outputs=outputs;self.join_delete_refs=self.refs
            self.join_route=None;self.join_warning='';self.join_network=plan
            benefit=sum(after)<sum(before) or after[0]>before[0] or plan["connectors"]>0
            self.join_changed=benefit and not same_vectors(self.contours,plan["routes"],self.scale)
            if not self.join_changed:after=before
            self.join_network_summary=continuous_summary(before,after,plan)
            self.join_counts=(before,after)
            if len(styles)>1:self.join_network_summary+=f'\nРазное оформление: {len(styles)} групп, обработаны отдельно.'
            return before,after
        ambiguous=ambiguous_endpoints(self.contours,tolerance)
        self.join_warning='Соединение пропущено: у концов несколько соседей в допуске.' if ambiguous else ''
        if ambiguous and not optimize_path:
            raise ValueError('У некоторых концов больше одного соседа в допуске. Соединение неоднозначно: выделите меньше объектов или уменьшите допуск.')
        if self.join_source is None:
            self.join_source=self.app.CreateCurve(self.doc)
            for s in self.join_open_refs.values():self.join_source.AppendCurve(s.DisplayCurve)
        curve=self.join_source.GetCopy()
        before=(sum(c['closed'] for c in self.contours),sum(not c['closed'] for c in self.contours))
        if self.join_open_refs and not ambiguous:curve.JoinTouchingSubpaths(True,tolerance/self.scale)
        untouched=sum(c['id'] not in self.join_open_refs for c in self.contours)
        after=(untouched+sum(bool(p.Closed) for p in curve.SubPaths),sum(not p.Closed for p in curve.SubPaths))
        self.joined=curve
        self.join_outputs=[(next(iter(self.join_open_refs.values())),curve)] if self.join_open_refs else []
        self.join_delete_refs=self.join_open_refs
        self.join_route=None
        self.join_changed=before!=after
        self.join_counts=(before,after)
        return before,after

    def apply_join(self,progress=None):
        if not self.join_changed:raise ValueError("Дополнительных изменений не требуется. Исходные векторы сохранены.")
        report=progress or (lambda *args:None)
        report('Проверка выделения',0,0)
        roots,leaves=self.join_selection(report,self.join_selection_only)
        current=[]
        for index,shape in enumerate(leaves.values()):
            report('Чтение перед проверкой дублей',index,len(leaves))
            current.extend(self.shape_contours(shape))
        groups=duplicate_groups(current,lambda done,total:report('Сверка дублей перед применением',done,total))
        if groups:
            raise ValueError('Появились дубли. Оптимизация не применена. Пересчитайте предпросмотр: сначала необходимо перенести дубли.')
        if {int(s.StaticID) for s in roots} != set(self.join_roots) or set(leaves)!=self.join_leaf_ids:
            raise ValueError('Выделение изменилось. Пересчитайте предпросмотр соединения.')
        if self.digest(current)!=self.signature:
            raise ValueError('Геометрия или слои изменились. Пересчитайте предпросмотр оптимизации.')
        with paused_redraw(self.app):
            self._apply_join(report)
            report('Обновление окна CorelDRAW',0,0)

    def _apply_join(self,progress=None):
        report=progress or (lambda *args:None)
        # Undo includes both the replacement and removal of the selected originals.
        self.doc.BeginCommandGroup('Соединить контуры DXF')
        changed=False
        try:
            results=self.app.CreateShapeRange()
            previous_by_layer={}
            report('Создание векторов',0,len(self.join_outputs))
            for index,(source,curve) in enumerate(self.join_outputs,1):
                if isinstance(curve,dict):
                    route=curve;curve=self.app.CreateCurve(self.doc)
                    def units(p):return p[0]/self.scale,p[1]/self.scale
                    path=curve.CreateSubPath(*units(route["segments"][0][1][0]))
                    for number,(kind,p) in enumerate(route["segments"]):
                        if number and number%100==0:report("Создание участков вектора "+str(index),number,len(route["segments"]))
                        if kind=="L":path.AppendLineSegment(*units(p[-1]))
                        else:path.AppendCurveSegment2(*units(p[-1]),*units(p[1]),*units(p[2]))
                    if route["closed"]:path.Closed=True
                result=source.Layer.CreateCurve(curve)
                changed=True
                result.Fill.ApplyNoFill()
                result.Outline.CopyAssign(source.Outline)
                result.Name=source.Name
                layer_name=source.Layer.Name
                previous=previous_by_layer.get(layer_name)
                if previous is not None:result.OrderBackOf(previous)
                results.Add(result)
                previous_by_layer[layer_name]=result
                report('Создание векторов',index,len(self.join_outputs))
            originals=self.app.CreateShapeRange()
            report('Подготовка замены исходных объектов',0,len(self.join_delete_refs))
            for index,s in enumerate(self.join_delete_refs.values(),1):
                originals.Add(s)
                report('Подготовка замены исходных объектов',index,len(self.join_delete_refs))
            report('Замена исходных объектов',0,0)
            originals.Delete()
            report('Обновление выделения',0,0)
            results.CreateSelection()
            report('Удаление прежних меток',0,0)
            for layer in self.marker_layers():
                layer.Editable=True
                layer.Delete()
        except Exception:
            report('Отмена незавершённых изменений',0,0)
            self.doc.EndCommandGroup()
            if changed:self.doc.Undo()
            raise
        else:
            self.doc.EndCommandGroup()
        self.refs={}
        self.join_source=None


def main():
    import tkinter as tk
    from tkinter import ttk,messagebox
    try:
        from .ui_shell import apply_theme, load_theme, save_theme, feedback, THEMES
        from .version import VERSION
        from .export_panel import ExportPanel
        from .check_pipeline import SharedCheck
    except ImportError:
        from ui_shell import apply_theme, load_theme, save_theme, feedback, THEMES
        from version import VERSION
        from export_panel import ExportPanel
        from check_pipeline import SharedCheck
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
    join_tab=ttk.Frame(notebook,padding=20)
    help_tab=ttk.Frame(notebook,padding=20)
    notebook.add(export_tab,text='Экспорт');notebook.add(check_tab,text='Проверка')
    notebook.add(join_tab,text='Оптимизация');notebook.add(help_tab,text='Слои и помощь')
    notebook.select(check_tab if '--check' in sys.argv or '--join' in sys.argv else export_tab)
    if '--join' in sys.argv:notebook.select(join_tab)
    if '--help' in sys.argv:notebook.select(help_tab)
    status=tk.StringVar(value='Готово. При проверке копии дублей переносятся на технические слои. Отмена переноса — в CorelDRAW.')
    ttk.Label(frame,textvariable=status,wraplength=1000,style='Muted.TLabel').pack(side='bottom',anchor='w',pady=(10,0),before=notebook)
    frame=check_tab
    def safe(action):
        try:action()
        except Exception as error:messagebox.showerror('DXF',str(error),parent=root)
    def macro(name):
        session.app.GMSManager.RunMacro('SkladCorelDXF','ExportDXF.'+name)
    ttk.Label(export_tab,text='Сохранить чертёж в DXF',style='Title.TLabel').pack(anchor='w',pady=(0,8))
    export_scope=tk.StringVar(value='selection' if session.app.ActiveSelectionRange.Count else 'page')
    scope_row=ttk.Frame(export_tab);scope_row.pack(fill='x',pady=6)
    export_options=[]
    for caption,value in [('Выделенные объекты','selection'),('Весь текущий лист','page')]:
        option=ttk.Radiobutton(scope_row,text=caption,variable=export_scope,value=value)
        option.pack(side='left',padx=(0,20));export_options.append(option)
    ttk.Label(export_tab,text='Масштаб 1:1   ·   Миллиметры   ·   Допуск оптимизации 0,1 мм').pack(anchor='w')
    ttk.Label(export_tab,text='DXF рядом с CDR. Существующий файл получит резервную копию.\nОткрытые контуры сохраняются без автоматического замыкания.',style='Muted.TLabel').pack(anchor='w',pady=8)
    export_path=tk.StringVar(value=str(session.doc.FullFileName or 'Несохранённый документ: папку предложим выбрать при экспорте.'))
    ttk.Label(export_tab,textvariable=export_path,wraplength=820).pack(anchor='w',pady=(0,4))
    def prepare_export():
        if state['busy']:raise ValueError('Дождитесь завершения проверки или соединения векторов.')
        session.current()
        if export_scope.get()=='selection' and not session.app.ActiveSelectionRange.Count:
            raise ValueError('Сначала выделите объекты в CorelDRAW.')
        return session.app.GMSManager.RunMacro('SkladCorelDXF','ExportDXF.PrepareExport',export_scope.get()=='selection')
    def export_busy(busy):
        notebook.tab(check_tab,state='disabled' if busy else 'normal')
        for option in export_options:option.configure(state='disabled' if busy else 'normal')
        status.set('Выполняется экспорт. Не редактируйте чертёж до завершения операции.' if busy else 'Готово. Итог экспорта и журнал находятся во вкладке «Экспорт».')
    exporter=ExportPanel(export_tab,prepare_export,export_busy,lambda:notebook.select(check_tab))
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
    ttk.Label(help_tab,text='Названия слоёв — латиницей. Размеры — в мм.\nОптимизация: весь лист или выделение → допуск → отчёт по слоям → применить. Каждый слой независим.\nТема Corel автоматически не считывается: выберите светлую или тёмную тему вверху.',style='Muted.TLabel',wraplength=850).pack(anchor='w',pady=15)
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
        check_scope=tk.StringVar(value='page')
        check_scope_row=ttk.Frame(frame);check_scope_row.pack(fill='x',pady=(8,4))
        check_scope_controls=[]
        for caption,value in [('Весь текущий лист','page'),('Только выделенное','selection')]:
            option=ttk.Radiobutton(check_scope_row,text=caption,variable=check_scope,value=value)
            option.pack(side='left',padx=(0,20));check_scope_controls.append(option)
        ttk.Label(frame,text='Каждый слой проверяется отдельно. Выделение задаётся в CorelDRAW.',wraplength=950).pack(anchor='w',pady=(0,4))
        ttk.Label(frame,text='Скрытые и технические слои, INFO, # и метки исключены.',style='Muted.TLabel').pack(anchor='w',pady=(0,8))
        ttk.Label(join_tab,text='Оптимизация',style='Title.TLabel').pack(anchor='w',pady=(0,12))
        ttk.Label(join_tab,text='Выберите весь лист или выделенные объекты. Каждый слой обрабатывается отдельно. Проверка и расчёт — по кнопке «Проверить».\nСоединение объединяет близкие концы. Замыкание выполняется только в пределах допуска.\nДубли автоматически переносятся на технические слои до расчёта. Соединение — по «Применить». Отмена — Ctrl+Z.',wraplength=900).pack(anchor='w',pady=(0,16))
        join_box=ttk.Frame(join_tab)
        join_box.pack(fill='x')
        join_fields=ttk.Frame(join_box);join_fields.pack(fill='x')
        ttk.Label(join_fields,text='Допуск, мм:').pack(side='left')
        join_tolerance=tk.StringVar(value='0,1')
        join_entry=ttk.Entry(join_fields,textvariable=join_tolerance,width=8)
        join_entry.pack(side='left',padx=8)
        join_optimize=tk.BooleanVar(value=False)
        join_optimize_control=ttk.Checkbutton(join_fields,text='Объединить в непрерывные векторы',variable=join_optimize)
        join_optimize_control.pack(side='left',padx=16)
        ttk.Label(join_box,text='С галочкой участки соединяются на пересечениях без повторного прохода.\nНовые соединения — только в пределах допуска. Слои INFO и # игнорируются.',wraplength=950,style='Muted.TLabel').pack(anchor='w',pady=(10,4))
        join_scope=tk.StringVar(value='page')
        join_scope_row=ttk.Frame(join_box);join_scope_row.pack(fill='x',pady=6)
        join_scope_controls=[]
        for title,mode in (('Весь текущий лист','page'),('Только выделенное','selection')):
            control=ttk.Radiobutton(join_scope_row,text=title,variable=join_scope,value=mode)
            control.pack(side='left',padx=(0,20));join_scope_controls.append(control)
        join_text=tk.StringVar(value='Выберите область и нажмите «Проверить».')
        from tkinter.scrolledtext import ScrolledText
        join_preview_box=ScrolledText(join_box,height=10,wrap='word',font=('Segoe UI',10),state='disabled')
        join_preview_box.pack(fill='x',pady=8)
        def update_join_preview(*args):
            join_preview_box.configure(state='normal');join_preview_box.delete('1.0','end')
            join_preview_box.insert('1.0',join_text.get());join_preview_box.configure(state='disabled')
        join_text.trace_add('write',update_join_preview)
        update_join_preview()
        join_stage=tk.StringVar(value='')
        join_elapsed=tk.StringVar(value='')
        join_report=tk.StringVar(value='')
        join_stage_row=ttk.Frame(join_box);join_stage_row.pack(fill='x')
        ttk.Label(join_stage_row,textvariable=join_stage).pack(side='left')
        ttk.Label(join_stage_row,textvariable=join_elapsed).pack(side='right')
        join_progress=ttk.Progressbar(join_box,maximum=100)
        join_progress.pack(fill='x',pady=(5,6))
        ttk.Label(join_box,textvariable=join_report,wraplength=950,justify='left').pack(anchor='w')
        settings=ttk.Frame(frame);settings.pack(fill='x',pady=8)
        ttk.Label(settings,text='Допуск проверки кривых, мм:').pack(side='left')
        tolerance=tk.StringVar(value='0,1')
        tolerance_entry=ttk.Entry(settings,textvariable=tolerance,width=10)
        tolerance_entry.pack(side='left',padx=8)
        ttk.Label(frame,text='Сначала копии дублей автоматически переносятся на технические слои. Затем проверяются оставшиеся векторы. Перенос можно отменить в CorelDRAW.',style='Muted.TLabel',wraplength=710).pack(anchor='w',pady=6)
        legend=ttk.Frame(frame)
        legend.pack(fill='x',pady=(2,10))
        filter_buttons={}
        filter_names={'duplicate':'Дубли','zero':'Нулевые','overlap':'Наложения','intersection':'Пересечения','open':'Открытые'}
        for index,kind in enumerate(filter_names):
            control=ttk.Button(legend,text=filter_names[kind]+' · 0',style=kind+'.TButton',command=lambda kind=kind:filter_issues(kind))
            control.grid(row=0,column=index,sticky='ew',padx=(0,6));legend.columnconfigure(index,weight=1)
            filter_buttons[kind]=control
        progress=ttk.Progressbar(frame,maximum=100)
        progress.pack(fill='x',pady=(0,8))
        check_log=ttk.LabelFrame(frame,text='Журнал проверки',padding=4)
        check_log.pack(fill='x',pady=(0,4))
        check_result=tk.StringVar(value='Проверка ещё не выполнена.')
        history=[];history_views=[]
        join_history=[];join_history_views=[]
        def record_join(entry):
            join_history.append(entry)
            for view in list(join_history_views):
                if not view.winfo_exists():join_history_views.remove(view);continue
                follow=view.yview()[1]>=.99
                view.configure(state='normal')
                if len(join_history)==1:view.delete('1.0','end')
                view.insert('end',('\n\n' if len(join_history)>1 else '')+entry);view.configure(state='disabled')
                if follow:view.see('end')
        def record_check(message):
            entry=time.strftime('[%H:%M:%S] ')+message
            history.append(entry)
            check_result.set(entry)
            for view in list(history_views):
                if not view.winfo_exists():history_views.remove(view);continue
                follow=view.yview()[1]>=.99
                view.configure(state='normal')
                if len(history)==1:view.delete('1.0','end')
                view.insert('end',('\n\n' if len(history)>1 else '')+entry);view.configure(state='disabled')
                if follow:view.see('end')
        def show_check_log(entries=None):
            window=tk.Toplevel(root);window.title('Журнал проверок' if entries is None else 'Журнал оптимизации');window.geometry('780x420')
            from tkinter.scrolledtext import ScrolledText
            palette=ttk.Style(root)
            view=ScrolledText(window,wrap='word',font=('Segoe UI',10),
                background=palette.lookup('TFrame','background'),foreground=palette.lookup('TLabel','foreground'))
            view.pack(fill='both',expand=True,padx=10,pady=10)
            view.insert('1.0','\n\n'.join(history if entries is None else entries) or 'Журнал пока пуст.')
            if entries is None:history_views.append(view)
            else:join_history_views.append(view)
            view.configure(state='disabled')
        ttk.Button(check_log,text='История',command=show_check_log).pack(side='right',padx=(8,0))
        ttk.Label(check_log,textvariable=check_result,wraplength=600).pack(side='left',fill='x',expand=True)
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
                count=sum(len(issue['copies']) if kind=='duplicate' else 1 for issue in state['issues'] if issue['kind']==kind)
                if kind=='duplicate':count=state.get('duplicates_moved',0)
                control.configure(text=filter_names[kind]+f' · {count}'+(' ✓' if state['filter']==kind else ''))
            for i,issue in enumerate(state['issues']):
                if state['filter'] and issue['kind']!=state['filter']:continue
                x,y=issue['point']
                title=names[issue['kind']]+(f': копий {len(issue["copies"])}' if issue['kind']=='duplicate' else '')
                tree.insert('','end',iid=str(i),tags=(issue['kind'],),values=(str(i+1)+'. '+title+(' *' if issue['approx'] else ''),issue['layer'],f'{x:.3f}',f'{y:.3f}'))
        def filter_issues(kind):
            if state['busy']:return
            state['filter']=None if state['filter']==kind else kind
            render_issues()
            if kind=='duplicate':
                status.set(f'Перенесено копий: {state.get("duplicates_moved",0)}. Они находятся на непечатных слоях _DXF_DUPLICATES и исключены из дальнейшей проверки и DXF.')
        cancel_event=threading.Event()
        join_cancel_event=threading.Event()
        def background():
            import pythoncom
            pythoncom.CoInitialize()
            try:
                # COM references stay in their owning apartment for all operations.
                worker_session=CorelSession()
                shared_check=SharedCheck(worker_session)
                join_session=CorelSession()
                while True:
                    command,data=commands.get()
                    if command=='stop':return
                    try:
                        if command=='join_preview':
                            revision,value,geometry_revision,optimize_path,selection_only=data
                            try:
                                refresh=geometry_revision!=getattr(join_session,'preview_revision',None)
                                last_progress=[0]
                                def preview_progress(*data):
                                    if join_cancel_event.is_set():raise InterruptedError("Расчёт отменён. Уже перенесённые дубли остаются на технических слоях; перенос можно отменить в CorelDRAW.")
                                    now=time.monotonic()
                                    if now-last_progress[0]>=.1:
                                        messages.put(("join_preview_progress",(revision,data)));last_progress[0]=now
                                before,after=join_session.prepare_join(value,refresh,optimize_path,preview_progress,selection_only,lambda counts:messages.put(("join_duplicates_moved",counts)))
                                join_session.preview_revision=geometry_revision
                                summary=join_session.join_network_summary or join_summary(before,after,value,join_session.join_route)
                                if join_session.join_warning:summary+='\n'+join_session.join_warning
                                skipped_summary=join_session.join_skipped_summary()
                                if skipped_summary:summary+='\n\n'+skipped_summary
                                messages.put(('join_preview',(revision,summary,join_session.join_changed,None,join_session.join_incomplete)))
                            except Exception as error:
                                messages.put(('join_preview',(revision,None,None,str(error),True)))
                        elif command=='join_apply':
                            started=time.monotonic()
                            join_session.apply_join(lambda *data:messages.put(('join_apply_progress',data)))
                            before,after=join_session.join_counts
                            details=(join_session.join_network['connectors'],join_session.join_network['max_connector']) if join_session.join_optimize_path else None
                            messages.put(('join_applied',(before,after,time.monotonic()-started,details,join_session.join_layer_reports+([join_session.join_skipped_summary()] if join_session.join_skipped_summary() else []))))
                        elif command in ('scan','export_check'):
                            value,only_selected,checked_layers,update_marks=data
                            def reading(layer,count):
                                if cancel_event.is_set():raise InterruptedError('Проверка отменена.')
                                if count%20==0:messages.put(('reading',(layer,count)))
                            last_check=[0,None]
                            def report(*args):
                                now=time.monotonic();key=(args[0],args[3])
                                if key!=last_check[1] or now-last_check[0]>=.1 or args[4]==args[5]:
                                    messages.put(('progress',args));last_check[:]=[now,key]
                            def operation(stage,done,total):
                                now=time.monotonic()
                                if stage!=last_check[1] or now-last_check[0]>=.1 or done==total:
                                    messages.put(('check_stage',(stage,done,total)));last_check[:]=[now,stage]
                            payload=shared_check.run(value,only_selected,checked_layers,cancel_event.is_set,
                                reading,report,lambda count:messages.put(('duplicates_moved',count)),
                                operation,lambda part:messages.put(('layer_done',part)))
                            result=dict(payload['result'])
                            if update_marks:messages.put(('marking',None))
                            result['marks_updated']=worker_session.refresh_marks(result,cancel_event.is_set) if update_marks else False
                            result['after_join']=command=='scan' and not update_marks
                            payload['result']=result
                            messages.put(('export_result' if command=='export_check' else 'result',payload))
                        else:
                            if command=='focus':worker_session.focus(data)
                            elif command=='mark':worker_session.mark(data,cancel_event.is_set)
                            elif command=='clear':worker_session.clear_marks()
                            messages.put(('action_done',command))
                    except Exception as error:messages.put(('export_error' if command=='export_check' else 'error',str(error)))
            except Exception as error:messages.put(('error',str(error)))
            finally:pythoncom.CoUninitialize()
        threading.Thread(target=background,daemon=True).start()
        names={'intersection':'Пересечение/касание','overlap':'Наложение','open':'Открытый конец','zero':'Нулевой участок','duplicate':'Дубли'}
        def scan(update_marks=True):
            if state['busy'] or exporter.busy:return
            try:value=float(tolerance.get().replace(',','.'))
            except ValueError:raise ValueError('Введите допуск числом, например 0,1 мм.')
            if not math.isfinite(value) or not .001<=value<=1:
                raise ValueError('Допуск проверки: от 0,001 до 1 мм.')
            state['issues']=[];state['duplicates_moved']=0; render_issues()
            state['complete']=False
            selection_only=check_scope.get()=='selection'
            state['scan_scope']=(selection_only,None)
            state['scanning']=True;state['issues']=[];render_issues()
            record_check('Проверка выполняется. Итог ещё не получен.')
            state['busy']=True;cancel_event.clear()
            join_check.configure(state='disabled');join_apply.configure(state='disabled');join_cancel.configure(state='disabled');join_entry.configure(state='disabled');join_optimize_control.configure(state='disabled')
            for control in join_scope_controls:control.configure(state='disabled')
            for control in (check,marks,clear,tolerance_entry,*check_scope_controls):control.configure(state='disabled')
            status.set('Чтение векторов из CorelDRAW…');progress.configure(mode='indeterminate');progress.start()
            cancel.configure(state='normal')
            commands.put(('scan',(value,selection_only,None,update_marks)))
        def export_preflight():
            if state['busy']:raise ValueError('Дождитесь завершения текущей операции.')
            value=float(tolerance.get().replace(',','.'))
            if not math.isfinite(value) or not .001<=value<=1:
                raise ValueError('Допуск проверки: от 0,001 до 1 мм.')
            state['scanning']=True;state['busy']=True;state['complete']=False
            state['filter']=None;state['issues']=[];render_issues()
            record_check('Общая проверка перед экспортом: сверка чертежа и параметров…')
            cancel_event.clear()
            for control in (check,marks,clear,tolerance_entry,join_check,join_apply,join_cancel,join_entry,join_optimize_control,*join_scope_controls,*check_scope_controls):
                control.configure(state='disabled')
            progress.configure(mode='indeterminate');progress.start()
            commands.put(('export_check',(value,export_scope.get()=='selection',None,False)))
        exporter.preflight=export_preflight
        exporter.cancel_preflight=cancel_event.set
        def finish():
            state['busy']=False;progress.stop();progress.configure(mode='determinate')
            edit_state='disabled' if exporter.busy else 'normal'
            for control in (check,clear,tolerance_entry,*check_scope_controls):control.configure(state=edit_state)
            marks.configure(state='normal' if state['issues'] and not exporter.busy else 'disabled')
            cancel.configure(state='disabled')
            join_check.configure(state=edit_state);join_cancel.configure(state=edit_state);join_entry.configure(state=edit_state);join_optimize_control.configure(state=edit_state)
            for control in join_scope_controls:control.configure(state=edit_state)
        original_export_busy=exporter.on_busy
        def shared_export_busy(busy):
            original_export_busy(busy)
            if not busy and not state['busy']:finish()
        exporter.on_busy=shared_export_busy
        def action(command,data=None):
            if state['busy']:return
            if exporter.busy and not (command=='focus' and exporter.allows_inspection):
                if command=='focus':status.set('Дождитесь подготовки DXF. При просмотре предупреждений переход к найденным местам будет доступен.')
                return
            cancel_event.clear()
            state['busy']=True
            join_check.configure(state='disabled');join_apply.configure(state='disabled');join_cancel.configure(state='disabled');join_entry.configure(state='disabled');join_optimize_control.configure(state='disabled')
            for control in join_scope_controls:control.configure(state='disabled')
            for control in (check,marks,clear,tolerance_entry,*check_scope_controls):control.configure(state='disabled')
            if command=='join_preview':
                join_cancel_event.clear()
                join_state['calculating']=True;join_state['started']=time.monotonic()
                join_stage.set('Чтение выделения CorelDRAW…');join_elapsed.set('Прошло 00:00')
                join_progress.configure(mode='indeterminate');join_progress.start()
                join_entry.configure(state='normal');join_cancel.configure(state='normal');join_optimize_control.configure(state='normal')
                for control in join_scope_controls:control.configure(state='normal')
            if command=='join_apply':
                join_state['applying']=True;join_state['started']=time.monotonic()
                join_stage.set('Начинаю применение…');join_elapsed.set('Прошло 00:00')
                join_report.set('')
                join_progress.configure(mode='indeterminate');join_progress.start()
            status.set('Переход к выбранному месту…' if command=='focus' else 'Выполняется действие в CorelDRAW…')
            if command=='mark':cancel.configure(state='normal')
            commands.put((command,data))
        def pick(event=None):
            rows=tree.selection()
            if rows:action('focus',state['issues'][int(rows[0])])
        tree.bind('<<TreeviewSelect>>',pick)
        def poll():
            while not messages.empty():
                kind,data=messages.get()
                if kind=='join_geometry_changed':
                    schedule_join_geometry()
                elif kind=='progress':
                    if str(progress['mode'])!='indeterminate':
                        progress.configure(mode='indeterminate');progress.start()
                    layer,index,total,stage,done,count=data
                    status.set(f'Слой {index}/{total}: «{layer}». {stage}: {done}/{count}.')
                    if exporter.busy and exporter.request is None:exporter.stage.set(status.get())
                elif kind=='check_stage':
                    stage,done,total=data
                    if str(progress['mode'])!='indeterminate':
                        progress.configure(mode='indeterminate');progress.start()
                    status.set(stage+(f': {done}/{total}' if total else '…'))
                    if exporter.busy and exporter.request is None:exporter.stage.set(status.get())
                elif kind=='layer_done':
                    state['issues'].extend(data['issues']);render_issues()
                    record_check(layer_check_summary(data))
                elif kind=='reading':
                    layer,count=data
                    status.set(f'Чтение CorelDRAW: слой «{layer}», контур/объект {count+1}…')
                    if exporter.busy and exporter.request is None:exporter.stage.set(status.get())
                elif kind=='marking':
                    status.set('Обновление меток на чертеже… Прежние метки заменяются, а не накапливаются.')
                elif kind=='quarantining':
                    status.set('Поиск дублей и перенос копий на технические слои…')
                elif kind=='action_done':
                    finish();status.set({'focus':'Место выделено и приближено в CorelDRAW.','mark':'Метки добавлены на отдельный непечатный слой.','clear':'Служебные метки удалены.'}[data])
                elif kind in ('error','export_error'):
                    finish();status.set(data);progress['value']=0
                    if join_state.get('applying'):
                        join_state['applying']=False;join_progress.stop();join_progress.configure(mode='determinate',value=0)
                        join_stage.set('Применение не завершено')
                        join_report.set('Ошибка: '+data)
                    if state.pop('scanning',False):
                        record_check('Проверка НЕ завершена. '+data+' Итоги завершённых слоёв сохранены выше; остальные слои не подтверждены.'+f' Перенесено дублей: {state.get("duplicates_moved",0)}; перенос можно отменить в CorelDRAW.')
                    join_apply.configure(state='disabled')
                    if kind=='export_error' or exporter.busy and exporter.request is None:
                        exporter.preflight_done(error=data)
                elif kind=='join_duplicates_moved':
                    text='Дубли перенесены на технические слои: '+str(sum(data.values()))+'.\n'+'\n'.join(f'Слой «{layer}»: {count}.' for layer,count in data.items())+'\nПеренос можно отменить в CorelDRAW (Ctrl+Z).'
                    join_report.set(text)
                    record_join(time.strftime('[%H:%M:%S] ')+text)
                    state['issues']=[];state['complete']=False;state['duplicates_moved']=0
                    render_issues();marks.configure(state='disabled')
                    check_result.set('В оптимизации перенесены дубли. Предыдущие результаты проверки устарели.')
                elif kind=='join_preview_progress':
                    revision,event=data
                    if revision==join_state['revision']:
                        stage,done,total=event
                        join_stage.set(stage+(f': {done} / {total}' if total else f': {done}' if done else '…'))
                        join_progress.stop()
                        if total:join_progress.configure(mode='determinate',value=100*done/total)
                        else:join_progress.configure(mode='indeterminate');join_progress.start()
                elif kind=='join_preview':
                    join_state['calculating']=False;join_progress.stop()
                    join_progress.configure(mode='determinate',value=0)
                    revision,summary,changed,error,incomplete=data
                    finish()
                    if revision!=join_state['revision']:
                        join_stage.set('Расчёт отменён. Нажмите «Проверить» для нового расчёта.')
                        record_join(time.strftime('[%H:%M:%S] ')+'Проверка оптимизации отменена или устарела. '+join_state.get('check_context',''))
                    elif error:
                        join_stage.set('Расчёт не завершён');join_text.set(error);status.set(error)
                        record_join(time.strftime('[%H:%M:%S] ')+'Проверка оптимизации НЕ завершена. '+join_state.get('check_context','')+'\n'+error)
                    else:
                        join_stage.set(('Расчёт завершён частично.' if incomplete else 'Расчёт завершён.')+' Соединение ещё не применено.');join_progress.configure(value=100)
                        join_text.set(('Текущая область после применения:\n' if join_report.get().startswith('Выполнено:') else '')+summary+('\nПрименять нечего: есть нерассчитанные слои.' if incomplete and not changed else '\nДополнительных изменений не требуется.' if not changed else ''))
                        record_join(time.strftime('[%H:%M:%S] ')+('Проверка оптимизации завершена частично.' if incomplete else 'Проверка оптимизации завершена.')+'\n'+join_state.get('check_context','')+'\n'+summary+f'\nВремя: {time.monotonic()-join_state["started"]:.1f} с.')
                        join_state['ready']=revision
                        join_apply.configure(state='normal' if changed else 'disabled')
                        status.set('Предпросмотр готов. '+('Не все слои рассчитаны: см. отчёт. Применение затронет только рассчитанные изменения.' if incomplete else 'Дубли проверены. Соединение выполняется по кнопке «Применить».'))
                elif kind=='join_apply_progress':
                    stage,done,total=data
                    join_stage.set(stage+(f': {done} / {total}' if total else '…'))
                    join_progress.stop()
                    if total:join_progress.configure(mode='determinate',value=100*done/total)
                    else:join_progress.configure(mode='indeterminate');join_progress.start()
                elif kind=='join_applied':
                    before,after,seconds,details,layer_reports=data
                    join_state['applying']=False
                    join_progress.stop();join_progress.configure(mode='determinate',value=100)
                    join_stage.set('Применение завершено')
                    join_elapsed.set(f'Время: {int(seconds)//60:02d}:{int(seconds)%60:02d}')
                    report_text=f'Выполнено: {sum(before)} → {sum(after)} векторов.\nЗамкнутых: {before[0]} → {after[0]}; открытых: {before[1]} → {after[1]}.'
                    if details:report_text+=f'\nНовых соединений: {details[0]}; максимальное — {details[1]:.4f} мм.'
                    join_report.set(report_text+'\nОтмена — Ctrl+Z в CorelDRAW. Чертёж автоматически не сохранён.')
                    record_join(time.strftime('[%H:%M:%S] ')+join_report.get()+'\n\n'+'\n\n'.join(layer_reports))
                    finish();join_state['ready']=None
                    join_text.set('Соединение применено.\nОтмена: Ctrl+Z в CorelDRAW.')
                    join_apply.configure(state='disabled')
                    invalidate()
                    schedule_join_geometry()
                elif kind=='duplicates_moved':
                    state['duplicates_moved']=data;render_issues()
                    if data:record_check(f'Перенесено дублей: {data}. Проверка остальных векторов продолжается. Перенос сохранится при отмене проверки; его можно отменить в CorelDRAW.')
                elif kind in ('result','export_result'):
                    result,skipped,count,value=data['result'],data['skipped'],data['count'],data['tolerance']
                    check_scope.set('selection' if data['selection'] else 'page')
                    state['scan_scope']=(data['selection'],data['layers'])
                    state['duplicates_moved']=data['duplicates_moved']
                    state['scanning']=False
                    state['complete']=result['complete']
                    record_check(('Использован результат: чертёж и параметры не изменились.\n' if data['reused'] else '')+check_summary(result,count,skipped,value,*state['scan_scope'])+f'\nПеренесено дублей: {state.get("duplicates_moved",0)}.')
                    state['issues']=result['issues']
                    render_issues()
                    finish();progress['value']=100 if result['complete'] else 0
                    status.set(f'Проверено контуров: {count}. Допуск: {value:g} мм. Найдено мест: {len(result["issues"])}. Пропущено объектов: {len(skipped)}. '+(('Выберите строку — Corel выделит и приблизит место.' if result['issues'] else 'По проверенным критериям проблем не найдено.') if result['complete'] else 'Достигнут лимит: результат НЕ полный.'))
                    status.set(status.get()+(' После соединения: «Показать метки» добавит метки. Ctrl+Z отменяет соединение.' if result.get('after_join') else (' Метки обновлены.' if result['marks_updated'] else ' Прежние метки сохранены.')))
                    if kind=='export_result':
                        error=None
                        if not result['complete']:error='Проверка не завершена. Экспорт остановлен.'
                        elif skipped:error='Есть непроверенные объекты. Экспорт остановлен; проверьте журнал.'
                        elif any(i['kind']=='zero' for i in result['issues']):error='Обнаружены нулевые участки. Исправьте их во вкладке «Проверка».'
                        exporter.preflight_done(dict(result,duplicates_moved=data['duplicates_moved'],reused=data['reused']),error)
                        if exporter.busy:
                            notebook.tab(check_tab,state='normal')
                            for control in (check,marks,clear,tolerance_entry,join_check,join_apply,join_cancel,join_entry,join_optimize_control,*join_scope_controls,*check_scope_controls):
                                control.configure(state='disabled')
                            status.set('Выполняется экспорт. Результаты общей проверки доступны. Не редактируйте чертёж до завершения операции.')
            if exporter.busy and exporter.request is None:
                seconds=int(time.monotonic()-exporter.started)
                exporter.elapsed.set(f'Прошло {seconds//60:02d}:{seconds%60:02d}')
            if join_state.get('applying') or join_state.get('calculating'):
                seconds=int(time.monotonic()-join_state['started'])
                join_elapsed.set(f'Прошло {seconds//60:02d}:{seconds%60:02d}')
            root.after(150,poll)
        row=ttk.Frame(frame);row.pack(side='bottom',fill='x',pady=10,before=tree)
        check=ttk.Button(row,text='Проверить',style='Accent.TButton',command=lambda:safe(scan));check.pack(side='left')
        marks=ttk.Button(row,text='Показать метки',command=lambda:action('mark',state['issues']),state='disabled');marks.pack(side='left',padx=6)
        clear=ttk.Button(row,text='Убрать мои метки',command=lambda:action('clear'));clear.pack(side='left')
        cancel=ttk.Button(row,text='Отменить проверку',command=cancel_event.set,state='disabled');cancel.pack(side='right')
        join_state={'revision':0,'geometry_revision':0,'ready':None,'timer':None,'cancelled':False}
        def preview_join():
            join_state['timer']=None
            if state['busy'] or exporter.busy:return
            try:
                value=float(join_tolerance.get().replace(',','.'))
                if not math.isfinite(value) or not 0<value<=1:raise ValueError()
            except ValueError:
                join_text.set('Введите допуск больше 0\nи не больше 1 мм.');return
            join_state['geometry_revision']+=1
            join_state['revision']+=1;join_state['ready']=None;join_state['cancelled']=False
            join_state['check_context']=('Весь текущий лист' if join_scope.get()=='page' else 'Только выделенное')+f'; допуск: {value:g} мм; непрерывные векторы: '+('да' if join_optimize.get() else 'нет')
            record_join(time.strftime('[%H:%M:%S] ')+ 'Проверка оптимизации начата. '+join_state['check_context'])
            join_text.set('Проверяю…')
            action('join_preview',(join_state['revision'],value,join_state['geometry_revision'],join_optimize.get(),join_scope.get()=='selection'))
        def schedule_join(*args):
            if join_state.get('calculating'):join_cancel_event.set()
            join_state['cancelled']=False
            join_state['revision']+=1;join_state['ready']=None
            join_apply.configure(state='disabled')
            if join_state['timer'] is not None:root.after_cancel(join_state['timer'])
            join_state['timer']=None
            if not join_state.get('calculating') and not join_state.get('applying'):
                join_stage.set('Нужна проверка — нажмите «Проверить».')
        def schedule_join_geometry(*args):
            join_state['geometry_revision']+=1
            schedule_join()
        def cancel_join_preview():
            join_cancel_event.set()
            join_stage.set('Останавливаю расчёт…' if join_state.get('calculating') else 'Расчёт отменён')
            join_state['revision']+=1;join_state['ready']=None;join_state['cancelled']=True
            if join_state['timer'] is not None:root.after_cancel(join_state['timer'])
            join_state['timer']=None
            join_apply.configure(state='disabled')
            join_text.set('Предпросмотр отменён. Перенесённые дубли остаются на технических слоях; отмена переноса — Ctrl+Z.')
            status.set('Предпросмотр отменён. Для нового расчёта нажмите «Проверить».')
        def apply_join_preview():
            if state['busy'] or join_state['ready']!=join_state['revision']:return
            action('join_apply')
        join_check=ttk.Button(join_box,text='Проверить',command=preview_join,style='Accent.TButton')
        join_check.pack(side='left',padx=(0,8),pady=(16,0))
        join_apply=ttk.Button(join_box,text='Применить',command=apply_join_preview,state='disabled')
        join_apply.pack(side='left',pady=(16,0))
        join_cancel=ttk.Button(join_box,text='Отменить',command=cancel_join_preview)
        join_cancel.pack(side='left',padx=8,pady=(16,0))
        ttk.Button(join_box,text='Журнал оптимизации',command=lambda:show_check_log(join_history)).pack(side='left',padx=8,pady=(16,0))
        join_tolerance.trace_add('write',schedule_join)
        join_optimize.trace_add('write',schedule_join)
        join_scope.trace_add('write',schedule_join_geometry)
        notebook.bind('<<NotebookTabChanged>>',lambda event:schedule_join_geometry() if notebook.select()==str(join_tab) else None)
        import pythoncom,win32com.client
        class JoinEvents:
            # COM can deliver events inside a native Tk message dispatch. Never
            # re-enter Tk from that callback; the normal UI poll handles updates.
            def changed(self,*args):messages.put(('join_geometry_changed',None))
            OnSelectionChange=changed
            OnShapeChange=changed
            OnShapeTransform=changed
            OnShapeMove=changed
            OnShapeDelete=changed
            OnShapeCreate=changed
            OnShapeDistort=changed
            OnLayerChange=changed
            OnPageActivate=changed
        join_events=win32com.client.WithEvents(session.doc,JoinEvents)
        def pump_join_events():
            pythoncom.PumpWaitingMessages()
            root.after(100,pump_join_events)
        root.after(100,pump_join_events)
        if notebook.select()==str(join_tab):schedule_join_geometry()
        def invalidate(*args):
            if state['busy']:return
            state['issues']=[];state['duplicates_moved']=0;render_issues();marks.configure(state='disabled')
            state['complete']=False
            progress['value']=0;status.set('Параметры изменены. Нажмите «Проверить».')
            check_result.set('Параметры изменены. Нужна новая проверка. Предыдущие итоги доступны в истории.')
        tolerance.trace_add('write',invalidate)
        check_scope.trace_add('write',invalidate)
        closing=[False]
        def close_window():
            if exporter.busy:
                if not closing[0]:exporter.cancel();closing[0]=True
                root.after(100,close_window)
                return
            if state['busy']:
                cancel_event.set();join_cancel_event.set();root.after(100,close_window)
                return
            cancel_event.set();commands.put(('stop',None));root.destroy()
            join_events.close()
        root.protocol('WM_DELETE_WINDOW',close_window)
        ttk.Label(frame,text='Метки не печатаются и не входят в DXF. Сохранение CDR — вручную.',style='Muted.TLabel',wraplength=710).pack(side='bottom',anchor='w',before=tree)
        poll()
        def refresh_theme(event=None):
            name='dark' if theme_value.get()=='Тёмная' else 'light'
            apply_theme(root,name,tree)
            join_preview_box.configure(background=ttk.Style(root).lookup('TEntry','fieldbackground'),foreground=ttk.Style(root).lookup('TEntry','foreground'))
            exporter.theme()
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
                    tab={'export':export_tab,'check':check_tab,'join':join_tab,'help':help_tab}.get(data.get('tab'))
                    if tab is not None and abs(time.time()-data.get('created',0))<10:
                        notebook.select(export_tab if exporter.busy else tab);root.deiconify();root.lift()
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
            mode='join' if '--join' in sys.argv else ('check' if '--check' in sys.argv else ('help' if '--help' in sys.argv else 'export'))
            temp=request.with_name('open-tab-'+str(os.getpid())+'.tmp')
            temp.write_text(json.dumps({'tab':mode,'created':time.time()}),encoding='utf-8')
            os.replace(temp,request)
        else:
            main()
    finally:
        win32api.CloseHandle(mutex)
