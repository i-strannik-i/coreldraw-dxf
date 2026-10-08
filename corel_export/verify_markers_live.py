"""Opt-in marker replacement integration test in a disposable Corel document."""
import win32com.client
from vector_tools import CorelSession
from vector_audit import audit


def main():
    app=win32com.client.gencache.EnsureDispatch('CorelDRAW.Application.25')
    previous=app.ActiveDocument if app.Documents.Count else None
    doc=app.CreateDocument();doc.Unit=3
    try:
        layer=doc.ActivePage.CreateLayer('CUT_ON')
        layer.CreateLineSegment(0,0,10,10)
        second=layer.CreateLineSegment(0,10,10,0)
        session=CorelSession()
        contours,_=session.snapshot(False)
        result=audit(contours)
        assert len(result['issues'])==1
        session.refresh_marks(result)
        original=session.marker_layers()[0].Name
        session.refresh_marks(result)
        assert len(session.marker_layers())==1
        assert session.marker_layers()[0].Name!=original
        assert session.marker_layers()[0].Shapes.Count==3
        doc.Undo()
        assert session.marker_layers()[0].Name==original
        # Cancellation after creation of a replacement must roll it back.
        checks=iter([False,False,True])
        try:session.mark(result['issues'],lambda:next(checks))
        except InterruptedError:pass
        else:raise AssertionError('Expected cancellation')
        assert len(session.marker_layers())==1
        assert session.marker_layers()[0].Name==original
        second.Move(50,0)
        contours,_=session.snapshot(False)
        clean=audit(contours)
        assert not clean['issues']
        session.refresh_marks(clean)
        assert not session.marker_layers()
        assert layer.Shapes.Count==2
        # All problem kinds at the same point must have non-overlapping numbers.
        from marker_layout import overlaps
        from vector_tools import ISSUE_STYLES
        issues=[dict(kind=kind,point=(5,5),ids=[int(second.StaticID)],layer='CUT_ON',approx=False)
                for kind in list(ISSUE_STYLES)*2]
        session.mark(issues)
        labels=[s for s in session.marker_layers()[0].Shapes if s.Type==6]
        assert len(labels)==10
        boxes=[(s.LeftX,s.BottomY,s.RightX,s.TopY) for s in labels]
        assert all(not overlaps(a,b,0) for i,a in enumerate(boxes) for b in boxes[i+1:])
        print('PASS: replacement, no accumulation, Undo, cancellation rollback, zero-issue cleanup; original vectors retained')
    finally:
        doc.Dirty=False;doc.Close()
        if previous is not None:previous.Activate()


if __name__=='__main__':main()
