"""Exercise grouped joins only on disposable Corel documents."""
from pathlib import Path
import json
import win32com.client
from vector_tools import CorelSession

app=win32com.client.gencache.EnsureDispatch('CorelDRAW.Application.25')
previous=app.ActiveDocument

def select(*shapes):
    items=app.CreateShapeRange()
    for shape in shapes:items.Add(shape)
    items.CreateSelection()
    return items

doc=app.CreateDocument();doc.Unit=3
try:
    layer=doc.ActivePage.CreateLayer('JOIN_GROUP_TEST')
    a=layer.CreateLineSegment(0,0,10,0)
    b=layer.CreateLineSegment(10.05,0,20,0)
    inner=select(a,b).Group()
    c=layer.CreateLineSegment(20.05,0,30,0)
    outer=select(inner,c).Group()
    untouched=layer.CreateLineSegment(100,0,110,0)
    closed=layer.CreateEllipse2(80,80,2,2)
    outer=select(outer,closed).Group()
    select(outer)
    session=CorelSession()
    assert session.prepare_join(.1)==((1,3),(1,1))
    assert session.prepare_join(.01,False)==((1,3),(1,3))
    assert outer.Shapes.Count==2 and inner.Shapes.Count==2
    session.prepare_join(.1,False)
    session.apply_join()
    assert layer.Shapes.Count==3
    assert closed.Type==2
    assert app.ActiveSelectionRange.Item(1).Curve.SubPaths.Count==1
    assert untouched.Curve.SubPaths.Count==1
    doc.Undo()
    groups=[s for s in layer.Shapes if s.Type==7]
    assert len(groups)==1 and groups[0].Shapes.Count==2
    assert any(s.Type==7 and s.Shapes.Count==2 for s in groups[0].Shapes)
    # Selecting only part of a group must preserve its unselected siblings.
    group=groups[0]
    inner=next(s for s in next(s for s in group.Shapes if s.Type==7).Shapes if s.Type==7)
    select(inner)
    session=CorelSession();session.prepare_join(.1);session.apply_join()
    assert group.Shapes.Count==2
    doc.Undo()
    # Same container IDs with newly inserted contents must invalidate preview.
    group=next(s for s in layer.Shapes if s.Type==7)
    select(group);session=CorelSession();session.prepare_join(.1)
    extra=group.Shapes.Item(1).Duplicate(50,0)
    select(group)
    try:session.apply_join()
    except ValueError:pass
    else:raise AssertionError('Changed group contents accepted')
    print('PASS: nested groups, tolerance, non-mutating preview, apply, Undo, unselected siblings, changed contents',flush=True)
finally:
    doc.Dirty=False;doc.Close();previous.Activate()

folder=Path(__file__).resolve().parent/'build/group-join-test'
folder.mkdir(parents=True,exist_ok=True)
copy=folder/'current-drawing-copy.cdr'
original_name=previous.FullFileName
original_dirty=previous.Dirty
try:
    previous.SaveAsCopy(str(copy),None)
finally:
    previous.FullFileName=original_name
    previous.Dirty=original_dirty
doc=app.OpenDocumentAsCopy(str(copy),None)
try:
    layer=doc.ActivePage.Layers.Item('BACK_V90_3mm')
    layer.Shapes.All().CreateSelection()
    session=CorelSession()
    roots,leaves=session.join_selection()
    print(json.dumps({'roots':len(roots),'groups':sum(s.Type==7 for s in roots),'curves':len(leaves),'layers':sorted({s.Layer.Name for s in leaves.values()})}),flush=True)
    try:
        before,after=session.prepare_join(.1)
        print(json.dumps({'before':before,'after':after}),flush=True)
        session.apply_join()
        assert layer.Shapes.Count==1
        doc.Undo()
        assert layer.Shapes.Count==len(roots)
        print('PASS: current drawing copy apply and Undo',flush=True)
    except ValueError as error:
        assert 'неоднозначно' in str(error),str(error)
        assert layer.Shapes.Count==len(roots)
        print('REAL DRAWING VALIDATION: '+str(error),flush=True)
finally:
    doc.Dirty=False;doc.Close();previous.Activate()

