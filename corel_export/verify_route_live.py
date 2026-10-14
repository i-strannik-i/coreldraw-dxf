"""Verify route ordering and real drawing geometry on a disposable saved copy."""
import json
from pathlib import Path
import time
import faulthandler
import win32com.client
from vector_tools import CorelSession

app=win32com.client.gencache.EnsureDispatch('CorelDRAW.Application.25')
previous=app.ActiveDocument
folder=Path(__file__).resolve().parent/'build/route-test'
folder.mkdir(parents=True,exist_ok=True)
trace=open(folder/'trace.log','w')
faulthandler.dump_traceback_later(60,repeat=True,file=trace)
original_name=previous.FullFileName;original_dirty=previous.Dirty
try:previous.SaveAsCopy(str(folder/'working-copy.cdr'),None)
finally:previous.FullFileName=original_name;previous.Dirty=original_dirty
doc=app.OpenDocumentAsCopy(str(folder/'working-copy.cdr'),None)
try:
    layer=doc.ActivePage.Layers.Item('BACK_V90_3mm')
    layer.Shapes.All().CreateSelection()
    session=CorelSession()
    layer_counts={l.Name:l.Shapes.Count for l in doc.ActivePage.Layers}
    def geometry_key(contours):
        return sorted(json.dumps((c['closed'],c['segments'])) for c in contours)
    started=time.monotonic()
    before,after=session.prepare_join(.1,optimize_path=True)
    original=list(session.contours);original_signature=session.signature
    assert before==after,'The full working layer should retain ambiguous contours'
    assert session.join_warning
    assert session.join_route[1]<session.join_route[0]
    report=dict(before=before,after=after,travel_mm=session.join_route,
                outputs=len(session.join_outputs),preview_seconds=time.monotonic()-started)
    print(json.dumps(report),flush=True)
    print('Applying route to copy...',flush=True)
    started=time.monotonic();session.apply_join()
    print('Applied. Reading actual geometry...',flush=True)
    checked=CorelSession();result,_=checked.snapshot(True)
    assert geometry_key(original)==geometry_key(result),'Geometry changed while ordering'
    assert len(result)==report['outputs']
    # Verify actual layer traversal (the order used by the DXF converter).
    selected_ids=set(checked.refs)
    actual=[int(s.StaticID) for s in layer.Shapes if int(s.StaticID) in selected_ids]
    assert actual==list(checked.refs),'Layer order differs from selected route'
    from route_order import route_length
    measured=route_length(result,list(range(len(result))))
    assert abs(measured-report['travel_mm'][1])<1e-6,(measured,report)
    for other in doc.ActivePage.Layers:
        if other.Name!=layer.Name:assert other.Shapes.Count==layer_counts[other.Name]
    report['apply_seconds']=time.monotonic()-started
    print('Geometry and order verified. Undo...',flush=True)
    doc.Undo()
    layer.Shapes.All().CreateSelection()
    restored=CorelSession();restored.snapshot(True)
    assert restored.signature==original_signature,'Undo did not restore original groups/geometry'
    assert layer.Shapes.Count==layer_counts[layer.Name]
    report['geometry_preserved']=True;report['undo_restored']=True
    (folder/'result.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    print('PASS: real drawing route, unchanged geometry/direction, actual order, other layers and Undo',flush=True)
finally:
    doc.Dirty=False;doc.Close();previous.Activate()
    faulthandler.cancel_dump_traceback_later();trace.close()
