import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from corel_export.export_dxf import read_request, export_request, GeometryBlocked, progress_text


class CorelExportTests(unittest.TestCase):
    def test_progress_labels_are_readable(self):
        self.assertEqual(progress_text('Оптимизация', dict(layer='D4', segments=12)),
                         'Оптимизация: Слой: D4 | Сегментов: 12')

    def test_geometry_errors_never_publish(self):
        for issue in ('zero_length',):
            for existing in (False, True):
                with self.subTest(issue=issue, existing=existing), tempfile.TemporaryDirectory() as folder:
                    source = Path(folder) / 'snapshot.cdr'
                    source.touch()
                    target = source.with_suffix('.dxf')
                    if existing:
                        target.write_bytes(b'previous')
                    request = Path(folder) / 'request.txt'
                    request.write_text(f'{source}\n{target}\npage\n', encoding='utf-16')
                    report = {'validation': {issue: [{'layer': 'CUT_OUT'}]}}
                    with patch('convert_cdr.convert', return_value=report), patch('save_outputs.publish_file') as publish:
                        with self.assertRaisesRegex(GeometryBlocked, 'CUT_OUT: 1') as error:
                            export_request(request, target, lambda *a, **kw: None)
                        self.assertIn('«Проверить»', str(error.exception))
                        publish.assert_not_called()
                    self.assertEqual(target.read_bytes() if existing else target.exists(),
                                     b'previous' if existing else False)

    def test_open_contour_exports_unchanged_with_warning(self):
        import ezdxf
        for gap in (.05, 2):
            with self.subTest(gap=gap), tempfile.TemporaryDirectory() as folder:
                source = Path(folder) / 'snapshot.cdr'
                source.touch()
                target = source.with_suffix('.dxf')
                request = Path(folder) / 'request.txt'
                request.write_text(f'{source}\n{target}\npage\n', encoding='utf-16')
                drawing = ezdxf.new()
                points = [(0, 0), (10, 0), (10, 10), (0, gap)]
                drawing.modelspace().add_lwpolyline(points, close=False, dxfattribs={'layer': 'CUT_OUT'})
                temporary = Path(folder) / 'converted.dxf'
                drawing.saveas(temporary)
                report = {'dxf': str(temporary), 'vector_check': {'complete': True, 'issues': []}, 'validation': {
                    'open_required': [{'layer': 'CUT_OUT', 'gap_mm': gap}], 'zero_length': []}}
                with patch('convert_cdr.convert', return_value=report):
                    result = export_request(request, target, lambda *a, **kw: None)
                entity = list(ezdxf.readfile(target).modelspace())[0]
                self.assertFalse(entity.closed)
                self.assertEqual(list(entity.get_points('xy')), points)
                self.assertIn('CUT_OUT: 1', result['warnings'][0])
                self.assertIn('сохранены открытыми', result['warnings'][0])

    def test_warning_confirmation_and_incomplete_check(self):
        from import_progress import ConversionCancelled
        cases = [(True, True), (True, False), (True, None), (False, True)]
        for complete, accepted in cases:
            with self.subTest(complete=complete, accepted=accepted), tempfile.TemporaryDirectory() as folder:
                source = Path(folder) / 'snapshot.cdr'
                source.touch()
                target = source.with_suffix('.dxf')
                target.write_bytes(b'previous')
                request = Path(folder) / 'request.txt'
                request.write_text(f'{source}\n{target}\npage\n', encoding='utf-16')
                report = {'dxf': 'temporary.dxf', 'vector_check': {'complete': complete, 'issues': [
                    {'kind': 'intersection', 'layer': 'CUT_OUT'}, {'kind': 'overlap', 'layer': 'P'}]}}
                confirm = (lambda summary: accepted) if accepted is not None else None
                with patch('convert_cdr.convert', return_value=report), patch('save_outputs.publish_file', return_value=None) as publish:
                    if complete and accepted:
                        result = export_request(request, target, lambda *a, **kw: None, confirm)
                        self.assertTrue(result['warnings_accepted'])
                        publish.assert_called_once()
                    else:
                        with self.assertRaises(ConversionCancelled if complete else GeometryBlocked):
                            export_request(request, target, lambda *a, **kw: None, confirm)
                        publish.assert_not_called()
                self.assertEqual(target.read_bytes(), b'previous')

    def test_unicode_request_and_scope(self):
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / 'деталь.cdr'
            source.touch()
            request = Path(folder) / 'request.txt'
            for scope in ('selection', 'page'):
                request.write_text(f'{source}\n{source.with_suffix(".dxf")}\n{scope}\n', encoding='utf-16')
                self.assertEqual(read_request(request), (source.resolve(), source.with_suffix('.dxf'), scope, 1))

    def test_unsaved_destination(self):
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / 'snapshot.cdr'
            source.touch()
            request = Path(folder) / 'request.txt'
            request.write_text(f'{source}\n\nselection\n', encoding='utf-16')
            self.assertIsNone(read_request(request)[1])

    def test_reject_bad_scope(self):
        with tempfile.TemporaryDirectory() as folder:
            request = Path(folder) / 'request.txt'
            request.write_text('source.cdr\nresult.dxf\nall\n', encoding='utf-16')
            with self.assertRaises(ValueError):
                read_request(request)

    def test_failed_conversion_does_not_publish(self):
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / 'snapshot.cdr'
            source.touch()
            target = Path(folder) / 'result.dxf'
            target.write_bytes(b'previous')
            request = Path(folder) / 'request.txt'
            request.write_text(f'{source}\n{target}\npage\n', encoding='utf-16')
            with patch('convert_cdr.convert', side_effect=ValueError('unsupported')):
                with self.assertRaises(ValueError):
                    export_request(request, target, lambda *a, **kw: None)
            self.assertEqual(target.read_bytes(), b'previous')


if __name__ == '__main__':
    unittest.main()
