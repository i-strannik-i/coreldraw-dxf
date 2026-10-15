import json
import io
from pathlib import Path
import sys
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile
from unittest.mock import Mock, patch

sys.path.insert(0, str(Path(__file__).parent / 'engine'))
import corel_connection as connection
from corel_export import setup_checks, update_manager
from corel_export.setup_corel import install, corel_executable
from test_corel_installer import workspace
from corel_export.workspace_install import merge_workspace, BAR, EXPORT, STANDARD


class VersionTests(unittest.TestCase):
    def test_sparse_2026_workspace_keeps_base_ui_and_adds_commands(self):
        with tempfile.TemporaryDirectory() as folder:
            icons = Path(folder)
            for name in ('CorelDXF.ico', 'VectorCheck.ico'):
                (icons / name).write_bytes(b'icon')
            xml = (f'<uiConfig><applicationInfo version="27"/>'
                   f'<states><toolbar guidRef="{STANDARD}" dock="left"/></states>'
                   '<commandBars><commandBarData guid="unrelated"/></commandBars></uiConfig>')
            stream = io.BytesIO()
            with zipfile.ZipFile(stream, 'w') as archive:
                archive.writestr('content/workspace.xml', xml)
                archive.writestr('content/keep.bin', b'keep')
            first = merge_workspace(stream.getvalue(), icons)
            second = merge_workspace(first, icons)
            with zipfile.ZipFile(io.BytesIO(second)) as archive:
                root = ET.fromstring(archive.read('content/workspace.xml'))
                self.assertIsNotNone(root.find(f'items/itemData[@guid="{EXPORT}"]'))
                self.assertEqual(len(root.findall(f'.//toolbar[@guidRef="{BAR}"]')), 1)
                self.assertEqual(root.find(f'.//toolbar[@guidRef="{STANDARD}"]').get('dock'), 'left')
                self.assertEqual(archive.read('content/keep.bin'), b'keep')

    def test_installed_tools_connect_to_their_own_version(self):
        for major in (25, 26, 27):
            with self.subTest(major=major), tempfile.TemporaryDirectory() as folder:
                exe = Path(folder) / 'VectorTools.exe'
                (Path(folder) / 'corel-target.json').write_text(json.dumps({'major': major}))
                app = Mock(VersionMajor=major)
                with patch.object(sys, 'frozen', True, create=True), patch.object(sys, 'executable', str(exe)), \
                        patch('win32com.client.gencache.EnsureDispatch', return_value=app) as dispatch:
                    self.assertIs(connection.connect(), app)
                    dispatch.assert_called_once_with(f'CorelDRAW.Application.{major}')

    def test_wrong_version_is_rejected(self):
        with patch('win32com.client.gencache.EnsureDispatch', return_value=Mock(VersionMajor=25)):
            with self.assertRaisesRegex(RuntimeError, 'другой версии'):
                connection.connect(27)

    def test_development_launch_refuses_ambiguous_running_versions(self):
        with patch.object(connection, 'configured_major', return_value=None), \
                patch('win32com.client.GetActiveObject', return_value=Mock()), \
                patch('win32com.client.gencache.EnsureDispatch') as dispatch:
            with self.assertRaises(RuntimeError):
                connection.connect()
            dispatch.assert_not_called()

    def test_probe_uses_selected_version_and_reports_actual_folders(self):
        with tempfile.TemporaryDirectory() as folder:
            app = Mock(VersionMajor=27, Version='27.2')
            app.Documents.Count = 0
            app.GMSManager.Projects.Count = 1
            app.GMSManager.UserGMSPath = str(Path(folder) / 'custom-gms')
            app.UserWorkspacePath = str(Path(folder) / 'custom-workspace')
            output = Path(folder) / 'probe.json'
            with patch('corel_export.setup_corel.corel_running', return_value=False), \
                    patch('win32com.client.DispatchEx', return_value=app) as dispatch:
                setup_checks.probe(output, major=27)
            dispatch.assert_called_once_with('CorelDRAW.Application.27')
            result = json.loads(output.read_text(encoding='utf-8'))
            self.assertTrue(result['ok'])
            self.assertEqual(result['gms_path'], app.GMSManager.UserGMSPath)
            self.assertEqual(result['workspace_path'], app.UserWorkspacePath)

    def test_installer_isolates_versions_and_writes_runtime_target(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            package = root / 'payload'
            (package / 'runtime').mkdir(parents=True)
            (package / 'runtime/CorelDXF.exe').write_bytes(b'runtime')
            (package / 'SkladCorelDXF.gms').write_bytes(b'macro')
            (package / 'icons').mkdir()
            for name in ('CorelDXF.ico', 'VectorCheck.ico'):
                (package / 'icons' / name).write_bytes(b'icon')
            other = root / '2024/Workspace/_default.cdws'
            other.parent.mkdir(parents=True)
            other.write_bytes(b'untouched-2024')
            ws = root / '2026/Workspace'
            ws.mkdir(parents=True)
            (ws / '_default.cdws').write_bytes(workspace())
            gms = root / 'custom-2026-gms'
            with patch('corel_export.setup_corel.corel_running', return_value=False):
                install(package, ws.parent, major=27, gms_path=gms, workspace_path=ws)
            self.assertEqual(other.read_bytes(), b'untouched-2024')
            self.assertEqual(json.loads((gms / 'SkladCorelDXF_runtime/corel-target.json').read_text()), {'major': 27})
            self.assertFalse((ws.parent / 'GMS').exists())

    def test_new_corel_cannot_receive_legacy_2024_only_update(self):
        info = dict(product='CorelDXF', corel_major=25, version='9.0.0',
                    file='SetupCorelDXF-9.0.0.exe', size=10, sha256='a' * 64)
        with patch.object(connection, 'configured_major', return_value=27):
            with self.assertRaises(ValueError):
                update_manager.validate_info(dict(info), '0.12.0')
            multi = dict(info, corel_majors=[25, 26, 27])
            self.assertTrue(update_manager.validate_info(multi, '0.12.0')['available'])


if __name__ == '__main__':
    unittest.main()
