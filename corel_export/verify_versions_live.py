"""Opt-in version smoke test using a disposable document only."""
import hashlib
import json
from pathlib import Path
import sys
import tempfile
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'engine'))
import corel_connection
from vector_tools import CorelSession
from check_pipeline import SharedCheck
from export_dxf import export_request
from workspace_install import merge_workspace


def main(major):
    import ezdxf
    with patch.object(corel_connection, 'configured_major', return_value=major):
        app = corel_connection.connect()
        previous = app.ActiveDocument if app.Documents.Count else None
        before = (previous.FullFileName, previous.Dirty, previous.ActivePage.Index,
                  app.ActiveSelectionRange.Count) if previous is not None else None
        doc = app.CreateDocument()
        doc.Unit = 3
        try:
            layer = doc.ActivePage.CreateLayer('OUT')
            original = layer.CreateRectangle2(10, 20, 30, 40)
            original.Fill.ApplyNoFill()
            duplicate = original.Duplicate(0, 0)
            selection = app.CreateShapeRange()
            selection.Add(original)
            selection.Add(duplicate)
            selection.CreateSelection()
            session = CorelSession()
            checked = SharedCheck(session).run(.1, True, None, lambda: False)
            assert checked['duplicates_moved'] == 1, checked
            assert checked['result']['complete'], checked
            assert not checked['result']['issues'], checked
            assert layer.Shapes.Count == 1
            print('PASS duplicate quarantine and vector audit', flush=True)
            with tempfile.TemporaryDirectory(prefix=f'CorelDXF-v{major}-') as temp:
                folder = Path(temp)
                layer.Shapes.Item(1).CreateSelection()
                options = app.CreateStructSaveAsOptions()
                options.Range = 2
                source = folder / 'snapshot.cdr'
                doc.SaveAsCopy(str(source), options)
                before_hash = hashlib.sha256(source.read_bytes()).hexdigest()
                target = folder / 'output.dxf'
                request = folder / 'request.txt'
                request.write_text(f'{source}\n{target}\nselection\n1\n', encoding='utf-16')
                result = export_request(request, target, lambda *a, **kw: None, prechecked=checked['result'])
                drawing = ezdxf.readfile(target)
                entities = list(drawing.modelspace())
                assert len(entities) == 1 and entities[0].dxftype() == 'LWPOLYLINE'
                assert entities[0].closed and entities[0].dxf.layer == 'OUT'
                vertices = list(entities[0].get_points('xy'))
                assert abs(max(p[0] for p in vertices) - min(p[0] for p in vertices) - 30) < 1e-6
                assert hashlib.sha256(source.read_bytes()).hexdigest() == before_hash
                assert not result['validation']['zero_length']
                print('PASS millimetre DXF, closed contour, layer, source unchanged', flush=True)
            doc.Activate()
            lines = doc.ActivePage.CreateLayer('CUT_ON')
            first = lines.CreateLineSegment(0, 0, 10, 0)
            second = lines.CreateLineSegment(10.05, 0, 20, 0)
            selection = app.CreateShapeRange()
            selection.Add(first)
            selection.Add(second)
            selection.CreateSelection()
            joining = CorelSession()
            counts = joining.prepare_join(.1, selection_only=True)
            assert sum(counts[1]) < sum(counts[0]), counts
            joining.apply_join()
            assert lines.Shapes.Count == 1
            doc.Undo()
            assert lines.Shapes.Count == 2
            print('PASS contour joining and Undo', flush=True)
            # Validate the real workspace format in memory; do not write it.
            workspace = Path(app.UserWorkspacePath) / '_default.cdws'
            merged = merge_workspace(workspace.read_bytes(), Path(__file__).parent / 'icons')
            assert merged
            print('PASS real workspace toolbar merge (in memory)', flush=True)
            result = {'major': major, 'version': str(app.Version), 'passed': True,
                      'checks': ['duplicates', 'audit', 'dxf', 'join', 'undo', 'workspace']}
            out = Path(__file__).parent / 'build/compatibility' / f'live{major}.json'
            out.parent.mkdir(parents=True, exist_ok=True)
            out.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
        finally:
            doc.Dirty = False
            doc.Close()
            if previous is not None:
                previous.Activate()
                after = (previous.FullFileName, previous.Dirty, previous.ActivePage.Index,
                         app.ActiveSelectionRange.Count)
                assert before == after, (before, after)
                print('PASS previously open document state preserved', flush=True)


if __name__ == '__main__':
    main(int(sys.argv[1]))
