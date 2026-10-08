"""Live regression: INFO text stays text and never enters contour validation."""
from pathlib import Path
import tempfile
import win32com.client
import ezdxf
from export_dxf import export_request


app=win32com.client.gencache.EnsureDispatch('CorelDRAW.Application.25')
previous=app.ActiveDocument if app.Documents.Count else None
doc=app.CreateDocument();doc.Unit=3
try:
    with tempfile.TemporaryDirectory(prefix='coreldxf-info-test-') as directory:
        folder=Path(directory)
        doc.ActivePage.CreateLayer('CUT_OUT').CreateRectangle2(0,0,50,40)
        doc.ActivePage.CreateLayer('INFO').CreateArtisticText(0,50,'Текст INFO OPR-247-0001')
        source=folder/'snapshot.cdr'
        doc.SaveAs(str(source),app.CreateStructSaveAsOptions())
        before=source.read_bytes()
        request=folder/'request.txt';target=folder/'test.dxf'
        request.write_text(f'{source}\n{target}\npage\n',encoding='utf-16')
        report=export_request(request,target,lambda *args,**kwargs:None)
        assert not report['validation']['zero_length']
        assert report['text_objects_outlined']==0
        drawing=ezdxf.readfile(target)
        assert len(drawing.modelspace().query('MTEXT'))==1
        assert len(drawing.modelspace().query('LWPOLYLINE'))==1
        assert source.read_bytes()==before
        doc.Dirty=False;doc.Close();doc=None
        print('PASS: INFO exported as MTEXT without outlining; cutting contour closed; source unchanged')
finally:
    if doc is not None:doc.Dirty=False;doc.Close()
    if previous is not None:previous.Activate()
