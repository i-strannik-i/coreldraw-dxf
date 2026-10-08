"""Read CorelDRAW through COM; export millimetre DXF without flattening curves."""
import argparse
from collections import Counter
import hashlib
import json
import math
from pathlib import Path
import shutil
import tempfile
import traceback
import uuid
import time

DEFAULT_ARC_TOLERANCE = 0.1


def bezier(points, t):
    p = list(points)
    while len(p) > 1:
        p = [tuple((1-t)*a + t*b for a, b in zip(x, y)) for x, y in zip(p, p[1:])]
    return p[0]


def split(points):
    levels = [list(points)]
    while len(levels[-1]) > 1:
        levels.append([tuple((a+b)/2 for a, b in zip(x, y))
                       for x, y in zip(levels[-1], levels[-1][1:])])
    return [p[0] for p in levels], [p[-1] for p in reversed(levels)]


def circle_fit(points, tolerance):
    """Conservative Bernstein bounds certify radial error, not sampled error."""
    a, b, c = points[0], bezier(points, .5), points[-1]
    bx, by, cx, cy = b[0]-a[0], b[1]-a[1], c[0]-a[0], c[1]-a[1]
    det = 2*(bx*cy-by*cx)
    if abs(det) < 1e-12:
        return None
    u = ((bx*bx+by*by)*cy-(cx*cx+cy*cy)*by)/det
    v = (bx*(cx*cx+cy*cy)-cx*(bx*bx+by*by))/det
    center, radius = (a[0]+u, a[1]+v), math.hypot(u, v)
    if radius <= tolerance:
        return None
    q = [(x-center[0], y-center[1]) for x, y in points]
    derivative = [(3*(q[i+1][0]-q[i][0]), 3*(q[i+1][1]-q[i][1])) for i in range(3)]
    angular = [sum(math.comb(3,i)*math.comb(2,k-i)/math.comb(5,k)*
                   (q[i][0]*derivative[k-i][1]-q[i][1]*derivative[k-i][0])
                   for i in range(4) if 0 <= k-i < 3) for k in range(6)]
    if not (min(angular) > 0 or max(angular) < 0):
        return None

    def bounded(p, depth=0):
        coeff = [sum(math.comb(3,i)*math.comb(3,k-i)/math.comb(6,k)*
                     sum(x*y for x,y in zip(p[i],p[k-i]))
                     for i in range(4) if 0 <= k-i < 4) for k in range(7)]
        low, high = (radius-tolerance)**2, (radius+tolerance)**2
        if min(coeff) >= low and max(coeff) <= high:
            return True
        if max(coeff) < low or min(coeff) > high or depth >= 14:
            return False
        left, right = split(p)
        return bounded(left, depth+1) and bounded(right, depth+1)

    if not bounded(q):
        return None
    start = math.degrees(math.atan2(q[0][1],q[0][0])) % 360
    end = math.degrees(math.atan2(q[-1][1],q[-1][0])) % 360
    if max(angular) < 0:
        start, end = end, start
    # Avoid an ambiguous nearly full-circle segment.
    if (end-start) % 360 > 180.000001:
        return None
    return center, radius, start, end


def rgb(color):
    copy = color.GetCopy()
    copy.ConvertToRGB()
    return int(copy.RGBRed), int(copy.RGBGreen), int(copy.RGBBlue)


def add_text_entity(drawing, shape, scale, layer, color):
    from ezdxf.colors import rgb2int
    story = shape.Text.Story
    text = str(story.Text).rstrip('\r\n')
    font = str(story.Font)
    size = abs(float(story.Size))
    if not text or not math.isfinite(size) or size <= 0:
        raise ValueError('Cannot export text with undefined font size on '+layer)
    angle = float(shape.RotationAngle)
    cx,cy = float(shape.CenterX),float(shape.CenterY)
    # Measure the text frame without rotation on a disposable duplicate.
    duplicate = shape.Duplicate()
    try:
        if angle:
            duplicate.RotateEx(-angle,cx,cy)
        x,y = float(duplicate.LeftX),float(duplicate.TopY)
    finally:
        duplicate.Delete()
    radians = math.radians(angle)
    insertion = ((cx+(x-cx)*math.cos(radians)-(y-cy)*math.sin(radians))*scale,
                 (cy+(x-cx)*math.sin(radians)+(y-cy)*math.cos(radians))*scale)
    style_name = 'CDR_'+hashlib.sha256(font.encode('utf-8')).hexdigest()[:12]
    if style_name not in drawing.styles:
        style = drawing.styles.new(style_name)
        style.set_extended_font_data(font)
    escaped = text.replace('\\','\\\\').replace('{','\\{').replace('}','\\}')
    escaped = escaped.replace('\r\n','\n').replace('\r','\n').replace('\n','\\P')
    drawing.modelspace().add_mtext(escaped,dxfattribs={
        'layer':layer,'true_color':rgb2int(color),'style':style_name,
        'char_height':size*25.4/72,'insert':insertion,'rotation':angle,
        'attachment_point':1,'width':0})
    return {'text':text,'font':font,'font_size_pt':size,'rotation':angle,
            'layer':layer,'insert_mm':insertion}


def convert(source, output_dir, tolerance=0, progress=None, optimize_tolerance=0, *, page_index=None, connected=False):
    progress = progress or (lambda stage, **details: None)
    started = time.monotonic()
    progress('Чтение исходного CDR')
    import ezdxf
    from ezdxf.colors import rgb2int
    import win32com.client
    source = Path(source).resolve(strict=True)
    if source.suffix.lower() != '.cdr':
        raise ValueError('Select a CDR file.')
    original_hash = hashlib.sha256(source.read_bytes()).hexdigest()
    if not math.isfinite(tolerance) or tolerance < 0:
        raise ValueError('Tolerance must be a finite non-negative number.')
    if not math.isfinite(optimize_tolerance) or optimize_tolerance < 0:
        raise ValueError('Invalid optimization tolerance')
    from optimize_curves import optimize
    from corel_geometry import decode_subpath
    from contour_dxf import polyline_vertices, requires_closed
    # Both approximations share one total error budget.
    fit_budget = optimize_tolerance / 2 if tolerance else optimize_tolerance
    arc_budget = min(tolerance, optimize_tolerance / 2) if optimize_tolerance and tolerance else tolerance
    output_dir = Path(output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=False)
    progress('Подключение к CorelDRAW')
    app = win32com.client.gencache.EnsureDispatch('CorelDRAW.Application.25')
    previous = app.ActiveDocument if app.Documents.Count else None
    if previous is None:
        # COM can start Corel invisibly; expose startup dialogs before OpenDocument.
        progress('Запуск окна CorelDRAW')
        app.Visible = True
    report = {'source': str(source), 'sha256': original_hash, 'units': 'mm',
              'arc_tolerance_mm': tolerance, 'layers': [], 'warnings': [],
              'text_objects_outlined': 0, 'editable_dxf_text': []}
    report.update(optimization_tolerance_mm=optimize_tolerance, segments_before=0, segments_after=0)
    report['validation'] = {'open_required': [], 'open_allowed': 0, 'closed': 0, 'zero_length': []}
    report['geometry_reader'] = 'GetCurveInfo'
    timings = report['timings_seconds'] = dict(open_document=0, read_geometry=0,
                                             optimize=0, emit_geometry=0)
    drawing = ezdxf.new('R2010')
    drawing.header['$INSUNITS'] = 4
    drawing.header['$MEASUREMENT'] = 1
    model = drawing.modelspace()
    transfer = ['CDR_CONTOURS_1']
    contour_count = 0
    segment_count = 0
    document = None
    try:
        with tempfile.TemporaryDirectory(prefix='cdr_copy_') as temp:
            working = Path(temp)/'source.cdr'
            shutil.copy2(source, working)
            try:
                progress('Открытие временной копии в CorelDRAW')
                phase_started = time.monotonic()
                document = app.OpenDocument(str(working))
                timings['open_document'] += time.monotonic() - phase_started
                if page_index is None and document.Pages.Count != 1:
                    raise ValueError('Prototype supports one-page CDR only; nothing exported.')
                chosen_page = page_index if page_index is not None else 1
                if isinstance(chosen_page, bool) or not isinstance(chosen_page, int) or not 1 <= chosen_page <= document.Pages.Count:
                    raise ValueError('Invalid source page index')
                scale = app.ConvertUnits(1, document.Unit, 3)
                # Master-page objects require a separate page-placement implementation.
                for master in document.MasterPage.Layers:
                    if master.Shapes.Count:
                        raise ValueError('Master-page objects are not supported yet: '+master.Name)

                for layer in document.Pages.Item(chosen_page).Layers:
                    if layer.Name.startswith('_DXF_CHECK_'):
                        continue
                    if layer.Shapes.Count == 0 and layer.Name == 'Guides':
                        continue
                    name = layer.Name
                    progress('Обработка слоя', layer=name, segments=segment_count, contours=contour_count)
                    if any(c in name for c in '<>/\\":;?*|=\t\r\n'):
                        raise ValueError('Layer name cannot be preserved in DXF: '+name)
                    base_color = rgb(layer.Color)
                    layer_info = {'name': name, 'rgb': base_color, 'visible': bool(layer.Visible),
                                  'entities': 0, 'source_shapes': layer.Shapes.Count}
                    dxf_layer = drawing.layers.get(name) if name in drawing.layers else drawing.layers.new(name)
                    dxf_layer.rgb = base_color
                    colors = set()
                    transfer.append('H\t'+name)

                    def emit_shapes(shapes, write_dxf=True):
                        nonlocal contour_count, segment_count
                        for shape in list(shapes):
                            progress('Чтение объектов', layer=name, segments=segment_count, contours=contour_count)
                            if shape.Type == 7:
                                emit_shapes(shape.Shapes,write_dxf)
                                continue
                            if shape.Type not in (1,2,3,6):
                                raise ValueError(f'Unsupported object type {shape.Type} in layer {name}')
                            try:
                                if shape.PowerClip is not None:
                                    raise ValueError('PowerClip is not supported: '+name)
                            except AttributeError:
                                pass
                            shape_dxf = write_dxf
                            if shape.Type == 6 and name.strip().upper() == 'INFO' and write_dxf:
                                text_color = rgb(shape.Fill.UniformColor) if shape.Fill.Type == 1 else base_color
                                report['editable_dxf_text'].append(add_text_entity(drawing,shape,scale,name,text_color))
                                colors.add(text_color)
                                layer_info['entities'] += 1
                                shape_dxf = False
                            if shape.Type == 6:
                                report['text_objects_outlined'] += 1
                            if shape.Type != 3:
                                shape.ConvertToCurves()
                                if shape.Type == 7:
                                    emit_shapes(shape.Shapes,shape_dxf)
                                    continue
                            color = rgb(shape.Outline.Color) if shape.Outline.Type != 0 else base_color
                            colors.add(color)
                            attrs = {'layer':name, 'true_color':rgb2int(color)}
                            curve = shape.Curve
                            for subpath in curve.SubPaths:
                                expected_count = subpath.Segments.Count
                                if not expected_count:
                                    raise ValueError('Empty subpath in layer '+name)
                                contour_count += 1
                                closed = bool(subpath.Closed)
                                progress('Пакетное чтение контура', layer=name,
                                         segments=segment_count, contours=contour_count)
                                phase_started = time.monotonic()
                                original_segments = decode_subpath(
                                    subpath.GetCurveInfo(), scale, expected_count, closed)
                                if closed:
                                    report['validation']['closed'] += 1
                                elif shape_dxf and requires_closed(name):
                                    report['validation']['open_required'].append({'layer': name, 'contour': contour_count,
                                        'gap_mm': math.dist(original_segments[0][1][0], original_segments[-1][1][-1])})
                                else:
                                    report['validation']['open_allowed'] += 1
                                for kind, points in original_segments:
                                    if all(math.dist(points[0], p) <= 1e-9 for p in points):
                                        report['validation']['zero_length'].append({'layer': name, 'contour': contour_count})
                                timings['read_geometry'] += time.monotonic() - phase_started
                                segment_count += len(original_segments)
                                first = original_segments[0][1][0]
                                transfer.append('\t'.join(map(str,['C',*first,int(closed)])))
                                phase_started = time.monotonic()
                                optimized = optimize(original_segments, fit_budget, lambda done,total:
                                    progress('Оптимизация узлов', layer=name, segments=segment_count,
                                             contours=contour_count, optimized=done, total=total))
                                timings['optimize'] += time.monotonic() - phase_started
                                report['segments_before'] += len(original_segments)
                                report['segments_after'] += len(optimized)
                                phase_started = time.monotonic()
                                if shape_dxf and connected:
                                    vertices = polyline_vertices(optimized, closed, arc_budget)
                                    model.add_lwpolyline(vertices, format='xyb', close=closed, dxfattribs=attrs)
                                    layer_info['entities'] += 1
                                for kind, points in optimized:
                                    start, c1, c2, end = points
                                    if kind == 'L':
                                        if shape_dxf and not connected:
                                            model.add_line(start,end,dxfattribs=attrs)
                                        transfer.append('\t'.join(map(str,['L',*end])))
                                    else:
                                        arc = circle_fit(points,arc_budget) if arc_budget else None
                                        if arc:
                                            center,radius,a,b = arc
                                            if shape_dxf and not connected:
                                                model.add_arc(center,radius,a,b,dxfattribs=attrs)
                                            cross = (start[0]-center[0])*(c1[1]-start[1])-(start[1]-center[1])*(c1[0]-start[0])
                                            transfer.append('\t'.join(map(str,['A',*end,*center,int(cross > 0)])))
                                        else:
                                            if shape_dxf and not connected:
                                                spline = model.add_spline(degree=3,dxfattribs=attrs)
                                                spline.control_points = [(x,y,0) for x,y in points]
                                                spline.knots = [0,0,0,0,1,1,1,1]
                                            transfer.append('\t'.join(map(str,['B',*end,*c1,*c2])))
                                    if shape_dxf and not connected:
                                        layer_info['entities'] += 1
                                transfer.append('E')
                                timings['emit_geometry'] += time.monotonic() - phase_started
                    emit_shapes(layer.Shapes)
                    # Aspire uses one display colour per layer; preserve uniform stroke colour.
                    if not layer.OverrideColor and len(colors) == 1:
                        layer_info['rgb'] = next(iter(colors))
                        dxf_layer.rgb = layer_info['rgb']
                    elif len(colors) > 1:
                        report['warnings'].append('Mixed object colours on layer '+name+'; DXF retains object colours, Aspire may display one layer colour.')
                    if not layer.Visible:
                        report['warnings'].append('Hidden source layer imported visibly: '+name)
                    report['layers'].append(layer_info)
            finally:
                if document is not None:
                    document.Dirty = False
                    document.Close()
                    document = None
    finally:
        if previous is not None:
            previous.Activate()
    if hashlib.sha256(source.read_bytes()).hexdigest() != original_hash:
        raise RuntimeError('Source file changed during conversion; output withheld.')
    if not len(model):
        raise ValueError('No supported vectors found.')
    target = output_dir/'drawing.dxf'
    progress('Запись DXF', segments=segment_count, contours=contour_count)
    drawing.saveas(target)
    progress('Проверка DXF', segments=segment_count, contours=contour_count)
    check = ezdxf.readfile(target)
    audit = check.audit()
    if audit.has_errors:
        raise RuntimeError('DXF audit failed: '+str(audit.errors))
    report['entities'] = dict(Counter(e.dxftype() for e in check.modelspace()))
    report['contours'] = contour_count
    report['elapsed_seconds'] = round(time.monotonic() - started, 2)
    report['timings_seconds'] = {key: round(value, 3) for key, value in timings.items()}
    (output_dir/'contours.txt').write_text('\n'.join(transfer),encoding='utf-8')
    report['contour_file'] = str(output_dir/'contours.txt')
    report['dxf'] = str(target)
    if report['text_objects_outlined']:
        report['warnings'].append('Text outlined on temporary copy. Verify fonts visually against CorelDRAW.')
    if report['editable_dxf_text']:
        report['warnings'].append('INFO text is editable MTEXT in DXF; Aspire contour transfer remains outlined. Text uses nominal font size and approximate top-left placement; mixed formatting, mirroring and text effects are not reproduced. Verify typography in the receiving application.')
    (output_dir/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    return report


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--source')
    parser.add_argument('--output')
    parser.add_argument('--arc-tolerance',type=float,default=0)
    parser.add_argument('--request')
    parser.add_argument('--progress')
    parser.add_argument('--cancel')
    parser.add_argument('--optimize-tolerance', type=float, default=0)
    parser.add_argument('--publish-job')
    args = parser.parse_args()
    if args.publish_job:
        folder = Path(args.publish_job).parent
        (folder/'published-job.txt').unlink(missing_ok=True)
        try:
            from save_outputs import publish_job
            publish_job(args.publish_job)
        except Exception:
            (folder/'publish.error').write_text(traceback.format_exc(), encoding='utf-8')
            raise
        return
    if not args.request:
        if args.progress:
            from import_progress import ProgressWriter
            writer = ProgressWriter(args.progress, args.cancel)
            try:
                convert(args.source, args.output, args.arc_tolerance, writer, args.optimize_tolerance)
                writer('Готово', finished=True)
            except Exception:
                Path(args.progress).with_suffix('.error').write_text(traceback.format_exc(), encoding='utf-8')
                raise
            return
        print(json.dumps(convert(args.source,args.output,args.arc_tolerance,optimize_tolerance=args.optimize_tolerance),ensure_ascii=False,indent=2))
        return
    import tkinter as tk
    from tkinter import filedialog, messagebox
    request = Path(args.request)
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost',True)
    try:
        source = filedialog.askopenfilename(title='CDR -> Aspire: select drawing',filetypes=[('CorelDRAW','*.cdr')],parent=root)
        if not source:
            return
        from import_options import choose_options
        options = choose_options(root,source,DEFAULT_ARC_TOLERANCE)
        if options is None:
            return
        output = Path(__file__).resolve().parent/'exports'/uuid.uuid4().hex
        from import_progress import run_with_progress
        result = run_with_progress(root, source, output, options['arc'], request, options['optimize'])
        if result is None:
            return
        from save_outputs import prepare_outputs
        report_path = output/'report.json'
        result = prepare_outputs(report_path)
        lines = ['CDR_IMPORT_RESULT_2',result['saved_dxf'],result['contour_file'],
                 result['aspire_staging'],str(report_path)]
        lines += ['\t'.join([layer['name']]+[str(c) for c in layer['rgb']]) for layer in result['layers']]
        request.with_suffix('.result').write_text('\n'.join(lines),encoding='utf-8')
    except Exception as exc:
        request.with_suffix('.error').write_text(traceback.format_exc(),encoding='utf-8')
        messagebox.showerror('CDR -> Aspire',str(exc),parent=root)
    finally:
        root.destroy()


if __name__ == '__main__':
    main()
