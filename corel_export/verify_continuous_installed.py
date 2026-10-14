"""Prepare/check a disposable document for the installed GUI Apply test."""
import json
import os
from pathlib import Path
import subprocess
import sys
import win32com.client
from demo_network_route import example
from version import VERSION

app=win32com.client.gencache.EnsureDispatch('CorelDRAW.Application.25')
folder=Path(__file__).resolve().parent/'build/continuous-test'
if '--verify' in sys.argv:
    doc=app.ActiveDocument
    assert doc.Name=='installed-input.cdr'
    layer=doc.ActivePage.Layers.Item('TEST')
    assert layer.Shapes.Count==1
    shape=layer.Shapes.Item(1)
    assert shape.Type==3 and shape.Curve.SubPaths.Count==1
    assert not shape.Curve.SubPaths.Item(1).Closed
    count=shape.Curve.SubPaths.Item(1).Segments.Count
    assert count==33,count
    doc.SaveAs(str(folder/'installed-result.cdr'),None)
    doc.Undo();assert layer.Shapes.Count==17
    print('PASS: installed GUI produced one open 33-segment vector; Undo restored 17 originals.')
    assert doc.ActivePage.Layers.Item('INFO').Shapes.Count==1
    assert doc.ActivePage.Layers.Item('#').Shapes.Count==1
    (folder/'installed.json').write_text(json.dumps(dict(version=VERSION,open_vectors=1,segments=count,undo=True)),encoding='utf-8')
else:
    doc=app.CreateDocument();doc.Unit=3
    doc.ActivePage.SetSize(210,210)
    layer=doc.ActivePage.CreateLayer('TEST')
    for c in example():
        curve=app.CreateCurve(doc)
        def shift(p):return p[0]+105,p[1]+105
        path=curve.CreateSubPath(*shift(c['segments'][0][1][0]))
        for kind,p in c['segments']:path.AppendCurveSegment2(*shift(p[-1]),*shift(p[1]),*shift(p[2]))
        shape=layer.CreateCurve(curve);shape.Fill.ApplyNoFill()
        shape.Outline.Width=.35;shape.Outline.Color.RGBAssign(30,130,75)
    doc.ActivePage.CreateLayer('INFO').CreateArtisticText(15,195,'Тестовый пример')
    doc.ActivePage.CreateLayer('#').CreateLineSegment(15,20,25,20)
    doc.ActivePage.Shapes.All().CreateSelection()
    doc.ActiveWindow.ActiveView.SetViewArea(0,0,210,210)
    doc.SaveAs(str(folder/'installed-input.cdr'),None)
    runtime=Path(os.environ['APPDATA'])/'Corel/CorelDRAW Graphics Suite 2024/Draw/GMS/SkladCorelDXF_runtime'
    subprocess.Popen([str(runtime/'VectorTools.exe'),'--join'],creationflags=subprocess.CREATE_NO_WINDOW)
    print('Prepared installed GUI test: 17 selected open curves.')
