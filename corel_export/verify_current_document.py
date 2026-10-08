"""Read-only check of the open drawing, with a temporary export and a time limit."""
import hashlib
import json
from pathlib import Path
import tempfile
import time

from vector_tools import CorelSession, excluded_layer
from vector_audit import audit
from export_dxf import export_request
from import_progress import ConversionCancelled


session = CorelSession()
source = Path(session.doc.FullFileName)
before = hashlib.sha256(source.read_bytes()).hexdigest()
folder = Path(tempfile.mkdtemp(prefix='coreldxf-full-check-'))
print('DIAGNOSTICS', folder, flush=True)
print('SOURCE', source, flush=True)
started = time.monotonic()
last = [0.0]


def progress(*args, **kwargs):
    now = time.monotonic()
    if now - started > 240:
        raise ConversionCancelled('Проверка остановлена по лимиту 240 секунд.')
    if now - last[0] > 10:
        print(args, kwargs, flush=True)
        last[0] = now


try:
    contours, skipped = session.snapshot(False, progress=progress)
    result = audit(contours, progress=progress)
    (folder / 'vector-check.json').write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print('CHECK', len(contours), 'contours;', skipped, 'skipped;', len(result['issues']),
          'issues; complete:', result['complete'], flush=True)
    request = folder / 'request.txt'
    request.write_text(f'{source}\n{folder / "test.dxf"}\npage\n{session.page.Index}\n', encoding='utf-16')
    # Refuse warnings: this verifies that the user's DXF cannot be silently published.
    report = export_request(request, folder / 'test.dxf', progress, lambda summary: False)
    print('EXPORT', report['check_status'], flush=True)
except ConversionCancelled as error:
    print('NOT PUBLISHED:', error, flush=True)
finally:
    assert hashlib.sha256(source.read_bytes()).hexdigest() == before
    print('SOURCE HASH UNCHANGED', flush=True)
