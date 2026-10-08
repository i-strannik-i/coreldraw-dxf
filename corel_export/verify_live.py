"""Opt-in integration test using disposable documents, never user drawings."""
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import time
import os
import win32com.client
import ezdxf


def main():
    app = win32com.client.gencache.EnsureDispatch('CorelDRAW.Application.25')
    previous = app.ActiveDocument if app.Documents.Count else None
    doc = app.CreateDocument()
    doc.Unit = 3
    try:
        with tempfile.TemporaryDirectory(prefix='corel_macro_verify_') as folder:
            folder = Path(folder)
            a = doc.ActivePage.CreateLayer('OUT')
            first = a.CreateRectangle2(10, 20, 30, 40)
            first.Outline.Color.RGBAssign(255, 0, 0)
            b = doc.ActivePage.CreateLayer('IN')
            second = b.CreateRectangle2(100, 200, 15, 25)
            second.Outline.Color.RGBAssign(0, 0, 255)
            doc.AddPages(1)
            doc.Pages.Item(2).Activate()
            doc.ActivePage.CreateLayer('OTHER_PAGE').CreateRectangle2(1, 2, 3, 4)
            doc.Pages.Item(1).Activate()
            doc.SaveAs(str(folder / 'source.cdr'), app.CreateStructSaveAsOptions())
            original = (folder / 'source.cdr').read_bytes()
            # An unsaved edit must be exported without altering the saved source.
            first.SetSize(35, 40)
            first.CreateSelection()
            before = (doc.FullFileName, doc.Dirty, doc.ActivePage.Index, app.ActiveSelectionRange.Count)
            for scope, range_value, count in [('selection', 2, 1), ('page', 2, 2), ('page_empty', 2, 2)]:
                if scope == 'page_empty':
                    doc.ClearSelection()
                    before = (doc.FullFileName, doc.Dirty, doc.ActivePage.Index, 0)
                job = folder / scope
                job.mkdir()
                options = app.CreateStructSaveAsOptions()
                options.Range = range_value
                if '--macro' not in sys.argv:
                    selected = app.ActiveSelectionRange
                    page_number = doc.ActivePage.Index
                    if scope != 'selection':
                        doc.ActivePage.Shapes.All.CreateSelection()
                    doc.SaveAsCopy(str(job / 'snapshot.cdr'), options)
                    doc.Pages.Item(page_number).Activate()
                    selected.CreateSelection()
                target = job / 'result.dxf'
                request = job / 'request.txt'
                export_scope = 'selection' if scope == 'selection' else 'page'
                request.write_text(f'{job / "snapshot.cdr"}\n{target}\n{export_scope}\n', encoding='utf-16')
                if '--macro' in sys.argv:
                    jobs = Path(tempfile.gettempdir()) / 'SkladCorelDXF' / 'jobs'
                    before_jobs = set(jobs.glob('*')) if jobs.exists() else set()
                    app.GMSManager.RunMacro('SkladCorelDXF', 'ExportDXF.RunExport', scope == 'selection')
                    created = set(jobs.glob('*')) - before_jobs
                    assert len(created) == 1, created
                    job = created.pop()
                    target = folder / 'source.dxf'
                    process = None
                else:
                    process = subprocess.Popen([sys.argv[1], str(request)])
                deadline = time.monotonic() + 90
                while (process.poll() is None if process else not (job / 'result.json').exists()) and time.monotonic() < deadline:
                    time.sleep(.25)
                    if (job / 'error.txt').exists():
                        raise RuntimeError((job / 'error.txt').read_text(encoding='utf-8'))
                if process and process.poll() is None:
                    raise RuntimeError('Export did not finish within 90 seconds')
                if process:
                    assert process.returncode == 0
                else:
                    assert (job / 'result.json').exists()
                    time.sleep(3)
                drawing = ezdxf.readfile(target)
                assert len(drawing.modelspace()) == count
                layers = {e.dxf.layer for e in drawing.modelspace()}
                assert layers == ({'OUT'} if scope == 'selection' else {'OUT', 'IN'})
                assert all(e.dxftype() == 'LWPOLYLINE' and e.closed for e in drawing.modelspace())
                points = [p for e in drawing.modelspace() if e.dxf.layer == 'OUT' for p in e.get_points('xy')]
                assert abs(max(p[0] for p in points) - min(p[0] for p in points) - 35) < 1e-6
                for entity in drawing.modelspace():
                    assert entity.rgb == ((255, 0, 0) if entity.dxf.layer == 'OUT' else (0, 0, 255))
                assert (folder / 'source.cdr').read_bytes() == original
                after = (doc.FullFileName, doc.Dirty, doc.ActivePage.Index, app.ActiveSelectionRange.Count)
                assert after == before, (before, after)
                print(scope, count, 'entities; layers/colors/unsaved edits/original verified')
    finally:
        doc.Dirty = False
        doc.Close()
        if previous is not None:
            previous.Activate()


if __name__ == '__main__':
    main()
