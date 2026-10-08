import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from corel_export.export_dxf import read_request, export_request


class CorelExportTests(unittest.TestCase):
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
