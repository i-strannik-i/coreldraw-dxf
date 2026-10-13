"""One live-document audit shared by inspection and export."""
from vector_audit import audit


class SharedCheck:
    def __init__(self, session):
        self.session = session
        self.key = None
        self.cached = None

    def run(self, tolerance, selection, layers, cancelled, reading=None, progress=None, moved=None, operation=None, layer_done=None):
        if cancelled():
            raise InterruptedError('Проверка отменена.')
        contours, skipped = self.session.snapshot(selection, reading, layers)
        if cancelled():
            raise InterruptedError('Проверка отменена.')
        names = frozenset(layers if layers is not None else (c['layer'] for c in contours))
        def key():
            return (tolerance, selection, names, self.session.signature, tuple(skipped))
        if self.cached is not None and self.key == key():
            if layer_done:
                for part in self.cached['result'].get('layers',[]):layer_done(dict(part,reused=True))
            return dict(self.cached, reused=True)
        self.key = None
        self.cached = None
        count = self.session.quarantine_duplicates(cancelled,progress=operation)
        if moved:
            moved(count)
        moved_layers=getattr(self.session,'duplicates_by_layer',{})
        def completed(part):
            part['duplicates_moved']=moved_layers.get(part['layer'],0)
            if layer_done:layer_done(part)
        result = audit(self.session.contours, tolerance=tolerance, progress=progress, cancelled=cancelled,layer_done=completed)
        result['tolerance_mm'] = tolerance
        payload = dict(result=result, skipped=skipped, count=len(self.session.contours),
                       tolerance=tolerance, selection=selection, layers=set(names),
                       duplicates_moved=count, reused=False)
        if result['complete'] and not cancelled():
            self.key = key()
            self.cached = payload
        return payload
