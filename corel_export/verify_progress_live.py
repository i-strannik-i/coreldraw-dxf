"""Live progress and service-layer exclusion on a disposable document."""
from pathlib import Path
import json
import win32com.client
from demo_network_route import example
from vector_tools import CorelSession

app=win32com.client.gencache.EnsureDispatch('CorelDRAW.Application.25')
previous=app.ActiveDocument;doc=app.CreateDocument();doc.Unit=3
try:
    layer=doc.ActivePage.CreateLayer('CUT');selection=app.CreateShapeRange()
    for c in example():
        curve=app.CreateCurve(doc);path=curve.CreateSubPath(*c['segments'][0][1][0])
        for kind,p in c['segments']:path.AppendCurveSegment2(*p[-1],*p[1],*p[2])
        shape=layer.CreateCurve(curve);shape.Fill.ApplyNoFill();selection.Add(shape)
    info=doc.ActivePage.CreateLayer('INFO');text=info.CreateArtisticText(0,100,'ignored')
    labels=doc.ActivePage.CreateLayer('#');line=labels.CreateLineSegment(200,0,210,0)
    selection.Add(text);selection.Add(line);selection.CreateSelection()
    s=CorelSession();contours,skipped=s.snapshot(True)
    assert len(contours)==17 and not skipped and {c['layer'] for c in contours}=={'CUT'}
    before,after=s.prepare_join(.1,optimize_path=True)
    assert before==(0,17) and after==(0,1)
    events=[];s.apply_join(lambda *data:events.append(data))
    assert layer.Shapes.Count==1 and info.Shapes.Count==labels.Shapes.Count==1
    assert info.Shapes.Item(1).StaticID==text.StaticID and labels.Shapes.Item(1).StaticID==line.StaticID
    assert ('Создание векторов',1,1) in events
    assert ('Проверка исходной геометрии',17,17) in events
    assert events[-1][0]=='Обновление окна CorelDRAW'
    doc.Undo();assert layer.Shapes.Count==17 and info.Shapes.Count==labels.Shapes.Count==1
    folder=Path(__file__).resolve().parent/'build/progress-test';folder.mkdir(parents=True,exist_ok=True)
    (folder/'result.json').write_text(json.dumps(dict(events=events,ignored_layers=['INFO','#'],undo=True),ensure_ascii=False,indent=2),encoding='utf-8')
    print('PASS: live progress, 17 -> 1, INFO and # untouched, Undo restores input.')
finally:
    doc.Dirty=False;doc.Close();previous.Activate()
