"""Regression checks for automatic quarantine in disposable Corel documents."""
import win32com.client
from vector_tools import CorelSession
from vector_audit import audit


app = win32com.client.gencache.EnsureDispatch('CorelDRAW.Application.25')
previous = app.ActiveDocument if app.Documents.Count else None
doc = app.CreateDocument()
doc.Unit = 3
try:
    layer = doc.ActivePage.CreateLayer('CUT_OUT')
    original = layer.CreateRectangle2(0, 0, 40, 30)
    original.Duplicate(0, 0)
    original.Duplicate(0, 0)
    line = layer.CreateLineSegment(60, 0, 80, 0)
    untouched = layer.CreateRectangle2(100, 0, 10, 10)
    untouched.Duplicate(0, 0)
    other = doc.ActivePage.CreateLayer('D4')
    other.CreateRectangle2(0, 0, 40, 30)
    other.Shapes.Item(1).Duplicate(0, 0)
    selected = app.CreateShapeRange()
    for shape in layer.Shapes:
        if shape.StaticID not in (untouched.StaticID,) and shape.PositionX < 90:
            selected.Add(shape)
    selected.CreateSelection()
    session = CorelSession()
    contours, skipped = session.snapshot(True, layers={'CUT_OUT'})
    assert len(contours) == 4 and not skipped
    assert session.quarantine_duplicates() == 2
    assert len(session.contours) == 2
    assert line.StaticID in session.refs
    session.unchanged()
    assert layer.Shapes.Count == 4 and other.Shapes.Count == 2
    result = audit(session.contours)
    assert [issue['kind'] for issue in result['issues']] == ['open', 'open']
    assert {sid for issue in result['issues'] for sid in issue['ids']} == {line.StaticID}
    assert session.refresh_marks(result)
    assert session.marker_layers()
    assert line.StaticID in {s.StaticID for s in app.ActiveSelectionRange}
    repeated=CorelSession()
    repeated_contours,_=repeated.snapshot(True,layers={'CUT_OUT'})
    assert len(repeated_contours)==2
    assert len(audit(repeated_contours)['issues'])==2
    doc.Undo()  # Marker update is distinct from the transfer.
    assert not session.marker_layers()
    doc.Undo()
    assert layer.Shapes.Count == 6 and other.Shapes.Count == 2
    assert not [l for l in doc.ActivePage.Layers if l.Name.startswith('_DXF_DUPLICATES')]
    # Cancellation before transfer leaves geometry unchanged.
    session.snapshot(False, layers={'CUT_OUT'})
    try:
        session.quarantine_duplicates(lambda: True)
    except InterruptedError:
        pass
    else:
        raise AssertionError('Cancellation ignored')
    assert layer.Shapes.Count == 6
    # Lock a copy in a duplicate group; no group may be moved partially.
    from duplicate_geometry import duplicate_groups
    groups = duplicate_groups(session.contours)
    copy_id = groups[0]['copies'][0]
    session.refs[copy_id].Locked = True
    try:
        session.quarantine_duplicates()
    except ValueError as error:
        assert 'заблокированные' in str(error)
    else:
        raise AssertionError('Locked copy transferred')
    assert layer.Shapes.Count == 6
    assert not [l for l in doc.ActivePage.Layers if l.Name.startswith('_DXF_DUPLICATES')]
    print('PASS: automatic transfer, selected scope, excluded layers, remaining defects, marker Undo, transfer Undo, cancellation, locked-copy atomic refusal', flush=True)
finally:
    doc.Dirty = False
    doc.Close()
    if previous is not None:
        previous.Activate()
