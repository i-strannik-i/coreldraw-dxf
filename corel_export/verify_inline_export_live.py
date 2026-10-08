"""Exercise VBA snapshots and the inline exporter on a disposable Corel document."""
from pathlib import Path
import tempfile
import time
import tkinter as tk
import win32com.client
import ezdxf
from export_panel import ExportPanel

app=win32com.client.gencache.EnsureDispatch('CorelDRAW.Application.25')
previous=app.ActiveDocument if app.Documents.Count else None
doc=app.CreateDocument();doc.Unit=3
root=tk.Tk();root.geometry('1000x760')
with tempfile.TemporaryDirectory(prefix='coreldxf-inline-live-') as directory:
    folder=Path(directory)
    try:
        layer=doc.ActivePage.CreateLayer('CUT_OUT')
        first=layer.CreateRectangle2(0,0,50,40)
        layer.CreateRectangle2(100,0,50,40)
        source=folder/'source.cdr'
        doc.SaveAs(str(source),app.CreateStructSaveAsOptions())
        original=source.read_bytes()
        layer.CreateEllipse2(200,20,10)
        first.CreateSelection()
        dirty=doc.Dirty
        scope=[True]
        def prepare():
            return app.GMSManager.RunMacro('SkladCorelDXF','ExportDXF.PrepareExport',scope[0])
        panel=ExportPanel(root,prepare,lambda busy:None,lambda:None)
        for selection,expected in ((True,1),(False,3)):
            scope[0]=selection
            panel.start()
            end=time.monotonic()+120
            while panel.busy and time.monotonic()<end:
                root.update();time.sleep(.02)
            if panel.busy:
                panel.cancel()
                while panel.busy:root.update();time.sleep(.02)
                raise AssertionError('Live export exceeded 120 seconds')
            assert panel.stage.get()=='Проверка пройдена',panel.text.get('1.0','end')
            drawing=ezdxf.readfile(folder/'source.dxf')
            curves=list(drawing.modelspace().query('LWPOLYLINE'))
            assert len(curves)==expected,(len(curves),expected)
            assert all(curve.closed for curve in curves)
            assert source.read_bytes()==original
            assert doc.Dirty==dirty
            assert app.ActiveDocument.Name==doc.Name
            assert app.ActiveSelectionRange.Count==1
            assert not panel.source.exists()
            assert not any(isinstance(w,tk.Toplevel) for w in root.winfo_children())
        assert 'Резервная копия:' in panel.text.get('1.0','end')
        print('PASS: inline selection/page export, closed DXF, backup, unchanged saved CDR, unsaved edits and selection preserved')
    finally:
        root.destroy()
        doc.Dirty=False;doc.Close()
        if previous is not None:previous.Activate()
