"""Installed UI smoke test. Click Check, inspect results, then close plugin."""
import hashlib
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
    shape=layer.CreateRectangle2(0,0,40,30)
    shape.Duplicate(0,0);shape.Duplicate(0,0)
    line=layer.CreateLineSegment(60,0,80,0)
    layer.Shapes.All().CreateSelection()
    executable=Path(os.environ['APPDATA'])/'Corel/CorelDRAW Graphics Suite 2024/Draw/GMS/SkladCorelDXF_runtime/VectorTools.exe'
    built=Path(__file__).parent/'dist/CorelDXF/VectorTools.exe'
    assert hashlib.sha256(executable.read_bytes()).digest()==hashlib.sha256(built.read_bytes()).digest()
    assert app.InitializeVBA()
    assert app.GMSManager.RunMacro('SkladCorelDXF','ExportDXF.InstallationCheck')=='SKLAD_DXF_OK'
    process=subprocess.Popen([str(executable),'--check'])
    print('READY: installed current build; 3 rectangles and 1 open line. Click Check, inspect, close plugin.',flush=True)
    process.wait(timeout=240)
    assert process.returncode==0
    assert layer.Shapes.Count==2, layer.Shapes.Count
    assert line.StaticID in {s.StaticID for s in app.ActiveSelectionRange}
    targets=[l for l in doc.ActivePage.Layers if l.Name.startswith('_DXF_DUPLICATES')]
    assert len(targets)==1 and targets[0].Shapes.Count==2 and not targets[0].Printable
    markers=[l for l in doc.ActivePage.Layers if l.Name.startswith('_DXF_CHECK_')]
    assert len(markers)==1
    labels=[s for s in markers[0].Shapes if s.Type==6]
    assert len(labels)==2
    print('PASS: installed hash, VBA launch, automatic transfer through UI, remaining open-line markers',flush=True)
finally:
    if process is not None and process.poll() is None:
        process.terminate();process.wait(timeout=10)
    doc.Dirty=False;doc.Close()
    if previous is not None:previous.Activate()
    elif app.Documents.Count==0:app.Quit()
