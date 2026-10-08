"""Manual installed-binary smoke test: run Check, inspect, then close plugin."""
import os
from pathlib import Path
import subprocess
import win32com.client

app=win32com.client.gencache.EnsureDispatch('CorelDRAW.Application.25')
previous=app.ActiveDocument if app.Documents.Count else None
doc=app.CreateDocument();doc.Unit=3
process=None
try:
    layer=doc.ActivePage.CreateLayer('CUT_OUT')
    shape=layer.CreateRectangle2(0,0,40,30);shape.Duplicate(0,0);shape.Duplicate(0,0)
    layer.Shapes.All().CreateSelection()
    executable=Path(os.environ['APPDATA'])/'Corel/CorelDRAW Graphics Suite 2024/Draw/GMS/SkladCorelDXF_runtime/VectorTools.exe'
    process=subprocess.Popen([str(executable),'--check'])
    print('READY: disposable document with 3 identical rectangles. Run Check, inspect, close plugin.',flush=True)
    process.wait(timeout=180)
    assert process.returncode==0
    markers=[l for l in doc.ActivePage.Layers if l.Name.startswith('_DXF_CHECK_')]
    assert len(markers)==1 and markers[0].Shapes.Count==3
    label=next(s for s in markers[0].Shapes if s.Type==6)
    color=label.Fill.UniformColor
    assert (color.RGBRed,color.RGBGreen,color.RGBBlue)==(39,115,58)
    assert layer.Shapes.Count==3
    print('PASS: installed checker identifies full duplicates, draws green marker/number/leader, preserves originals',flush=True)
finally:
    if process is not None and process.poll() is None:
        process.terminate();process.wait(timeout=10)
    doc.Dirty=False;doc.Close()
    if previous is not None:previous.Activate()
