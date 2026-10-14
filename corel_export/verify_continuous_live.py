"""Test the integrated Apply and Undo on disposable Corel documents/copies."""
import json
import math
from pathlib import Path
import sys
import time
import win32com.client
from vector_tools import CorelSession,decode_subpath

app=win32com.client.gencache.EnsureDispatch('CorelDRAW.Application.25')
previous=app.ActiveDocument
folder=Path(__file__).resolve().parent/'build/continuous-test';folder.mkdir(parents=True,exist_ok=True)
working='--working' in sys.argv
if working:
    original=next(d for d in app.Documents if d.Name=='правка.cdr')
    filename,dirty=original.FullFileName,original.Dirty
    try:original.SaveAsCopy(str(folder/'working-copy.cdr'),None)
    finally:original.FullFileName=filename;original.Dirty=dirty
    doc=app.OpenDocumentAsCopy(str(folder/'working-copy.cdr'),None)
    layer=doc.ActivePage.Layers.Item('BACK_V90_3mm')
else:
    from demo_network_route import example
    doc=app.CreateDocument();doc.Unit=3
    layer=doc.ActivePage.CreateLayer('TEST')
    for c in example():
        curve=app.CreateCurve(doc);path=curve.CreateSubPath(*c['segments'][0][1][0])
        for kind,p in c['segments']:path.AppendCurveSegment2(*p[-1],*p[1],*p[2])
        shape=layer.CreateCurve(curve);shape.Fill.ApplyNoFill()
    layer.Shapes.All().Group()
try:
    layer.Shapes.All().CreateSelection()
    counts={l.Name:l.Shapes.Count for l in doc.ActivePage.Layers}
    session=CorelSession();started=time.monotonic()
    snapshot=session.snapshot
    def capture(*args,**kwargs):
        result=snapshot(*args,**kwargs)
        (folder/('working-source.json' if working else 'example-source.json')).write_text(json.dumps(session.contours),encoding='utf-8')
        return result
    session.snapshot=capture
    before,after=session.prepare_join(.1,optimize_path=True)
    signature=session.signature
    expected=session.join_network['routes']
    report=dict(before=before,after=after,connectors=session.join_network['connectors'],
                max_connector=session.join_network['max_connector'],repeat=0,
                preview_seconds=time.monotonic()-started)
    print(json.dumps(report),flush=True)
    assert counts=={l.Name:l.Shapes.Count for l in doc.ActivePage.Layers},'Preview changed drawing'
    if not working:assert after==(0,1)
    print('Applying...',flush=True)
    session.apply_join()
    shapes=list(app.ActiveSelectionRange)
    assert len(shapes)==len(expected)
    # Match by complete control-point sequence; Corel rotates closed subpaths.
    def matches(actual,route):
        wanted=route['segments']
        if len(actual)!=len(wanted):return False
        def eq(a,b):
            if a[0]!=b[0]:return False
            indices=range(4) if a[0]=='B' else (0,3)
            return all(math.dist(a[1][i],b[1][i])<.00015 for i in indices)
        shifts=range(len(wanted)) if route['closed'] else (0,)
        return any(all(eq(a,wanted[(i+shift)%len(wanted)]) for i,a in enumerate(actual)) for shift in shifts)
    pending=list(expected)
    for shape in shapes:
        assert shape.Curve.SubPaths.Count==1
        path=shape.Curve.SubPaths.Item(1)
        actual=decode_subpath(path.GetCurveInfo(),session.scale,path.Segments.Count,bool(path.Closed))
        assert all(any(math.dist(p[0],q)>.00001 for q in p[1:]) for kind,p in actual),'Zero segment in Corel result'
        found=next((i for i,r in enumerate(pending) if r['closed']==bool(path.Closed) and matches(actual,r)),None)
        assert found is not None,'Corel geometry differs from preview'
        pending.pop(found)
    assert not pending
    for other in doc.ActivePage.Layers:
        if other.Name!=layer.Name:assert other.Shapes.Count==counts[other.Name]
    doc.SaveAs(str(folder/('working-result.cdr' if working else 'example-result.cdr')),None)
    doc.Undo();layer.Shapes.All().CreateSelection()
    restored=CorelSession();restored.snapshot(True)
    assert restored.signature==signature,'Undo changed source geometry'
    assert counts=={l.Name:l.Shapes.Count for l in doc.ActivePage.Layers}
    report.update(corel_geometry_verified=True,undo_verified=True,total_seconds=time.monotonic()-started)
    (folder/('working.json' if working else 'example.json')).write_text(json.dumps(report,indent=2),encoding='utf-8')
    print('PASS '+json.dumps(report),flush=True)
finally:
    doc.Dirty=False;doc.Close();previous.Activate()
