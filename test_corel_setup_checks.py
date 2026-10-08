import json
from pathlib import Path
import struct
import tempfile
import unittest
from unittest.mock import Mock, patch

from corel_export import setup_checks as checks


class SetupChecksTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)

    def binary(self, machine):
        path = self.folder / 'CorelDRW.exe'
        data = bytearray(128)
        data[:2] = b'MZ'
        data[60:64] = struct.pack('<I', 64)
        data[64:70] = b'PE\0\0' + machine
        path.write_bytes(data)
        return path

    def test_x64_accepted(self):
        checks.validate_binary(self.binary(b'\x64\x86'))

    def test_x86_rejected(self):
        with self.assertRaises(RuntimeError):
            checks.validate_binary(self.binary(b'\x4c\x01'))

    def test_missing_packages_rejected(self):
        with self.assertRaises(RuntimeError):
            checks.vba_packages(self.folder / 'Programs64/CorelDRW.exe')

    def test_tampered_signature_rejected(self):
        response = Mock(stdout=json.dumps({'status': 'HashMismatch', 'signer': 'Microsoft Corporation'}))
        with patch.object(checks.subprocess, 'run', return_value=response):
            with self.assertRaises(RuntimeError):
                checks.verify_package(self.folder / 'vba.msi')

    def test_unknown_signer_rejected(self):
        response = Mock(stdout=json.dumps({'status': 'Valid', 'signer': 'Unknown'}))
        with patch.object(checks.subprocess, 'run', return_value=response):
            with self.assertRaises(RuntimeError):
                checks.verify_package(self.folder / 'vba.msi')

    def test_valid_signature_accepted(self):
        response = Mock(stdout=json.dumps({'status': 'Valid', 'signer': 'CN=Microsoft Corporation'}))
        with patch.object(checks.subprocess, 'run', return_value=response):
            checks.verify_package(self.folder / 'vba.msi')

    def test_probe_does_not_touch_running_corel(self):
        output = self.folder / 'result.json'
        with patch('corel_export.setup_corel.corel_running', return_value=True), \
                patch('win32com.client.DispatchEx') as dispatch:
            checks.probe(output)
        dispatch.assert_not_called()
        self.assertFalse(json.loads(output.read_text(encoding='utf-8'))['ok'])

    def test_probe_verifies_macro_and_closes_only_empty_instance(self):
        app = Mock(VersionMajor=25, Version='25.1')
        app.Documents.Count = 0
        app.GMSManager.Projects.Count = 1
        app.GMSManager.RunMacro.return_value = 'SKLAD_DXF_OK'
        output = self.folder / 'result.json'
        with patch('corel_export.setup_corel.corel_running', return_value=False), \
                patch('win32com.client.DispatchEx', return_value=app):
            checks.probe(output, installed=True)
        self.assertTrue(json.loads(output.read_text(encoding='utf-8'))['ok'])
        app.Quit.assert_called_once()
        app.GMSManager.RunMacro.assert_called_once()

    def test_open_document_never_closed(self):
        app = Mock(VersionMajor=25, Version='25.1')
        app.Documents.Count = 1
        app.GMSManager.Projects.Count = 1
        output = self.folder / 'result.json'
        with patch('corel_export.setup_corel.corel_running', return_value=False), \
                patch('win32com.client.DispatchEx', return_value=app):
            checks.probe(output)
        app.Quit.assert_not_called()
        self.assertEqual(json.loads(output.read_text(encoding='utf-8'))['kind'], 'open')

    def test_missing_vba_classified(self):
        app = Mock(VersionMajor=25, Version='25.1')
        app.InitializeVBA.return_value = False
        app.Documents.Count = 0
        output = self.folder / 'result.json'
        with patch('corel_export.setup_corel.corel_running', return_value=False), \
                patch('win32com.client.DispatchEx', return_value=app):
            checks.probe(output)
        self.assertEqual(json.loads(output.read_text(encoding='utf-8'))['kind'], 'vba')


if __name__ == '__main__':
    unittest.main()
