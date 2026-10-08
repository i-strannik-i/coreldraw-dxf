"""Destructive operations are tested only in a disposable Corel document."""
from pathlib import Path
import tempfile
import win32com.client
import ezdxf
from vector_tools import CorelSession
from vector_audit import audit
from export_dxf import export_request

app=win32com.client.gencache.EnsureDispatch('CorelDRAW.Application.25')
previous=app.ActiveDocument if app.Documents.Count else None
doc=app.CreateDocument();doc.Unit=3
try:
    layer=doc.ActivePage.CreateLayer('CUT_OUT')
    first=layer.CreateRectangle2(0,0,50,40)
    first.Duplicate(0,0);first.Duplicate(0,0)
    doc.ActivePage.CreateLayer('D4').CreateRectangle2(0,0,50,40)
    session=CorelSession();contours,_=session.snapshot(False)
    result=audit(contours);groups=[i for i in result['issues'] if i['kind']=='duplicate']
    assert len(groups)==1 and len(groups[0]['copies'])==2
    session.refresh_marks(result)
    assert session.move_duplicates(groups)==2
    assert layer.Shapes.Count==1
    targets=[l for l in doc.ActivePage.Layers if l.Name.startswith('_DXF_DUPLICATES')]
    assert len(targets)==1 and targets[0].Shapes.Count==2 and not targets[0].Printable
    assert len(session.snapshot(False)[0])==2
    doc.Undo()
    assert layer.Shapes.Count==3
    assert not [l for l in doc.ActivePage.Layers if l.Name.startswith('_DXF_DUPLICATES')]
    assert len(session.marker_layers())==1
    contours,_=session.snapshot(False);groups=[i for i in audit(contours)['issues'] if i['kind']=='duplicate']
    first.Move(1,0)
    try:session.move_duplicates(groups)
    except ValueError:pass
    else:raise AssertionError('Stale geometry accepted')
    first.Move(-1,0)
    contours,_=session.snapshot(False);groups=[i for i in audit(contours)['issues'] if i['kind']=='duplicate']
    session.move_duplicates(groups)
    with tempfile.TemporaryDirectory(prefix='coreldxf-duplicate-test-') as directory:
        folder=Path(directory);source=folder/'snapshot.cdr';target=folder/'result.dxf'
        doc.SaveAsCopy(str(source),app.CreateStructSaveAsOptions())
        request=folder/'request.txt';request.write_text(f'{source}\n{target}\npage\n1\n',encoding='utf-16')
        export_request(request,target,lambda *args,**kwargs:None)
        drawing=ezdxf.readfile(target)
        assert len(drawing.modelspace().query('LWPOLYLINE'))==2
        assert not any(l.dxf.name.startswith('_DXF_DUPLICATES') for l in drawing.layers)
        saved=target.read_bytes()
        job=folder/'only-duplicates';job.mkdir()
        quarantine=next(l for l in doc.ActivePage.Layers if l.Name.startswith('_DXF_DUPLICATES'))
        quarantine.Shapes.All().CreateSelection()
        options=app.CreateStructSaveAsOptions();options.Range=win32com.client.constants.cdrSelection
        isolated=job/'snapshot.cdr';doc.SaveAsCopy(str(isolated),options)
        request=job/'request.txt';request.write_text(f'{isolated}\n{target}\nselection\n1\n',encoding='utf-16')
        try:export_request(request,target,lambda *args,**kwargs:None)
        except ValueError as error:assert 'нет объектов' in str(error) or 'Не найдены векторы' in str(error)
        else:raise AssertionError('Empty quarantine export replaced the DXF')
        assert target.read_bytes()==saved
    # Never pull an individual child out of its designer's group.
    remaining=layer.Shapes.Item(1)
    child=remaining.Duplicate(0,0)
    extra=layer.CreateRectangle2(100,100,10,10)
    pair=app.CreateShapeRange();pair.Add(child);pair.Add(extra);pair.Group()
    contours,_=session.snapshot(False);groups=[i for i in audit(contours)['issues'] if i['kind']=='duplicate']
    assert groups
    before=len([l for l in doc.ActivePage.Layers if l.Name.startswith('_DXF_DUPLICATES')])
    try:session.move_duplicates(groups)
    except ValueError as error:assert 'группе' in str(error)
    else:raise AssertionError('Grouped duplicate was moved')
    assert len([l for l in doc.ActivePage.Layers if l.Name.startswith('_DXF_DUPLICATES')])==before
    print('PASS: same-layer full duplicates, move copies only, one Undo, stale detection, DXF exclusion')
finally:
    doc.Dirty=False;doc.Close()
    if previous is not None:previous.Activate()
