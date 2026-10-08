"""Manual smoke test: click Export in the installed UI, then close that UI."""
import os
from pathlib import Path
import subprocess
import tempfile
import win32com.client
import ezdxf

app=win32com.client.gencache.EnsureDispatch('CorelDRAW.Application.25')
previous=app.ActiveDocument if app.Documents.Count else None
doc=app.CreateDocument();doc.Unit=3
folder=Path(tempfile.mkdtemp(prefix='coreldxf-frozen-smoke-'))
source=folder/'inline-smoke.cdr'
process=None
try:
    shape=doc.ActivePage.CreateLayer('CUT_OUT').CreateRectangle2(0,0,50,40)
    doc.SaveAs(str(source),app.CreateStructSaveAsOptions());shape.CreateSelection()
    original=source.read_bytes()
    executable=Path(os.environ['APPDATA'])/'Corel/CorelDRAW Graphics Suite 2024/Draw/GMS/SkladCorelDXF_runtime/VectorTools.exe'
    process=subprocess.Popen([str(executable),'--export'])
    print('READY: click Export, inspect the result, close plugin. Fixture: '+str(source),flush=True)
    process.wait(timeout=180)
    assert process.returncode==0,process.returncode
    curves=list(ezdxf.readfile(source.with_suffix('.dxf')).modelspace().query('LWPOLYLINE'))
    assert len(curves)==1 and curves[0].closed
    assert source.read_bytes()==original
    print('PASS: installed frozen exporter produced a closed DXF; original CDR unchanged',flush=True)
finally:
    if process is not None and process.poll() is None:
        process.terminate();process.wait(timeout=10)
    doc.Dirty=False;doc.Close()
    if previous is not None:previous.Activate()
