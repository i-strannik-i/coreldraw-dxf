"""Verify joining on a disposable drawing, including preview and Undo."""
import win32com.client
from vector_tools import CorelSession

app = win32com.client.Dispatch('CorelDRAW.Application.25')
previous = app.ActiveDocument if app.Documents.Count else None
doc = app.CreateDocument()
doc.Unit = 3
try:
    layer = doc.ActivePage.CreateLayer('CUT_ON')
    a = layer.CreateLineSegment(0, 0, 10, 0)
    b = layer.CreateLineSegment(10.05, 0, 20, 0)
    a.Fill.ApplyNoFill(); b.Fill.ApplyNoFill()
    shapes = app.CreateShapeRange(); shapes.Add(a); shapes.Add(b); shapes.CreateSelection()
    session = CorelSession()
    before, after = session.prepare_join(.01)
    assert before == after == (0, 2), (before, after)
    before, after = session.prepare_join(.1)
    assert before == (0, 2) and after == (0, 1), (before, after)
    assert layer.Shapes.Count == 2
    session.apply_join()
    assert layer.Shapes.Count == 1
    doc.Undo()
    assert layer.Shapes.Count == 2
    other = doc.ActivePage.CreateLayer('OTHER')
    c = other.CreateLineSegment(20.05, 0, 30, 0)
    c.Fill.ApplyNoFill()
    shapes = app.CreateShapeRange(); shapes.Add(layer.Shapes.Item(1)); shapes.Add(c); shapes.CreateSelection()
    try:
        CorelSession().prepare_join(.1)
    except ValueError as error:
        assert 'одному слою' in str(error)
    else:
        raise AssertionError('Cross-layer join was allowed')
    print('PASS: tolerance recalculation, non-mutating preview, apply, single Undo, cross-layer rejection')
finally:
    doc.Dirty = False
    doc.Close()
    if previous is not None:
        previous.Activate()
