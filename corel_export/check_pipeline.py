"""One live-document audit shared by inspection and export."""
from vector_audit import audit


class SharedCheck:
    def __init__(self, session):
        self.session = session
        self.key = None
        self.cached = None

    def run(self, tolerance, selection, layers, cancelled, reading=None, progress=None, moved=None):
        if cancelled():
            raise InterruptedError('Проверка отменена.')
        contours, skipped = self.session.snapshot(selection, reading, layers)
        if cancelled():
            raise InterruptedError('Проверка отменена.')
        names = frozenset(layers if layers is not None else (c['layer'] for c in contours))
        def key():
            return (tolerance, selection, names, self.session.signature, tuple(skipped))
        if self.cached is not None and self.key == key():
            return dict(self.cached, reused=True)
        self.key = None
        self.cached = None
        count = self.session.quarantine_duplicates(cancelled)
        if moved:
            moved(count)
        result = audit(self.session.contours, tolerance=tolerance, progress=progress, cancelled=cancelled)
        result['tolerance_mm'] = tolerance
        payload = dict(result=result, skipped=skipped, count=len(self.session.contours),
                       tolerance=tolerance, selection=selection, layers=set(names),
                       duplicates_moved=count, reused=False)
        if result['complete'] and not cancelled():
            self.key = key()
            self.cached = payload
        return payload
