import io
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
import xml.etree.ElementTree as ET
import zipfile

from corel_export.workspace_install import merge_workspace, BAR, EXPORT, HELP, STANDARD, CHECK, JOIN
from corel_export.setup_corel import install


def workspace(anchor=True):
    root = ET.Element('uiConfig')
    ET.SubElement(root, 'items')
    bars = ET.SubElement(root, 'commandBars')
    ET.SubElement(bars, 'commandBarData', guid='unrelated', userCaption='Keep me')
    host = ET.SubElement(ET.SubElement(root, 'states'), 'dockHost')
    if anchor:
        ET.SubElement(host, 'toolbar', guidRef=STANDARD)
        ET.SubElement(ET.SubElement(bars, 'commandBarData', guid=STANDARD), 'toolbar')
    stream = io.BytesIO()
    with zipfile.ZipFile(stream, 'w') as archive:
        archive.writestr('content/workspace.xml', ET.tostring(root, short_empty_elements=False))
        archive.writestr('content/settings.ini', b'keep-original')
    return stream.getvalue()


class InstallerTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.icons = self.folder / 'icons'
        self.icons.mkdir()
        for name in ('CorelDXF.ico', 'LayerHelp.ico','VectorCheck.ico','VectorJoin.ico'):
            (self.icons / name).write_bytes(b'icon')

    def test_preserves_other_entries_and_is_idempotent(self):
        once = merge_workspace(workspace(), self.icons)
        twice = merge_workspace(once, self.icons)
        with zipfile.ZipFile(io.BytesIO(twice)) as archive:
            self.assertEqual(archive.read('content/settings.ini'), b'keep-original')
            root = ET.fromstring(archive.read('content/workspace.xml'))
            panel = root.find(f'.//commandBarData[@guid="{BAR}"]/toolbar')
            self.assertIsNotNone(panel)
            self.assertEqual(panel.get('itemFace'), 'textRightOfImage')
            self.assertEqual([item.get('guidRef') for item in panel], [EXPORT, CHECK])
            self.assertTrue(all(item.get('itemFace') == 'textRightOfImage' for item in panel))
            self.assertEqual(len(root.findall(f'.//toolbar[@guidRef="{BAR}"]')), 1)
            standard = root.find(f'.//commandBarData[@guid="{STANDARD}"]/toolbar')
            self.assertEqual(len(standard.findall('item')), 0)
            for guid in (EXPORT, CHECK):
                self.assertEqual(len(root.findall(f'.//item[@guidRef="{guid}"]')), 1)
            self.assertIsNotNone(root.find('.//commandBarData[@guid="unrelated"]'))
            for guid in (EXPORT, CHECK):
                self.assertEqual(len(root.findall(f'.//itemData[@guid="{guid}"]')), 1)
                self.assertEqual(archive.read('content/icons/' + guid + '.ico'), b'icon')
                self.assertEqual(root.find(f'.//itemData[@guid="{guid}"]').get('icon'), 'guid://' + guid)

    def test_reinstall_preserves_user_toolbar_position(self):
        data = merge_workspace(workspace(), self.icons)
        out = io.BytesIO()
        with zipfile.ZipFile(io.BytesIO(data)) as source, zipfile.ZipFile(out, 'w') as target:
            for name in source.namelist():
                content = source.read(name)
                if name == 'content/workspace.xml':
                    content = content.replace((f'guidRef="{BAR}" dock="top"').encode(),
                                              (f'guidRef="{BAR}" dock="left"').encode())
                target.writestr(name, content)
        result = merge_workspace(out.getvalue(), self.icons)
        with zipfile.ZipFile(io.BytesIO(result)) as archive:
            root = ET.fromstring(archive.read('content/workspace.xml'))
            self.assertEqual(root.find(f'.//toolbar[@guidRef="{BAR}"]').get('dock'), 'left')

    def test_missing_anchor_fails_without_result(self):
        with self.assertRaises(ValueError):
            merge_workspace(workspace(False), self.icons)

    def test_reinstall_migrates_shortcut_command_id(self):
        out = io.BytesIO()
        with zipfile.ZipFile(io.BytesIO(workspace())) as source, zipfile.ZipFile(out, 'w') as target:
            for name in source.namelist():
                data = source.read(name)
                if name == 'content/workspace.xml':
                    root = ET.fromstring(data)
                    ET.SubElement(root.find('items'), 'itemData', guid='old-dxf',
                                  dynamicCommand='SkladCorelDXF.ExportDXF.ShowExporter')
                    keys = ET.SubElement(ET.SubElement(root, 'shortcutKeyTables'), 'table', tableID='main')
                    ET.SubElement(ET.SubElement(keys, 'keySequence', itemRef='old-dxf'), 'key', ctrl='true').text = 'd'
                    data = ET.tostring(root, short_empty_elements=False)
                target.writestr(name, data)
        result = merge_workspace(out.getvalue(), self.icons)
        with zipfile.ZipFile(io.BytesIO(result)) as archive:
            root = ET.fromstring(archive.read('content/workspace.xml'))
            self.assertEqual(root.find('.//keySequence').get('itemRef'), EXPORT)

    def test_corel_self_closed_tags_survive_reinstall(self):
        data = merge_workspace(workspace(), self.icons)
        out = io.BytesIO()
        with zipfile.ZipFile(io.BytesIO(data)) as source, zipfile.ZipFile(out, 'w') as target:
            for name in source.namelist():
                content = source.read(name)
                if name == 'content/workspace.xml':
                    content = content.replace(b'></itemData>', b'/>').replace(b'></item>', b'/>')
                target.writestr(name, content)
        result = merge_workspace(out.getvalue(), self.icons)
        with zipfile.ZipFile(io.BytesIO(result)) as archive:
            root = ET.fromstring(archive.read('content/workspace.xml'))
            self.assertIsNotNone(root.find('.//commandBarData[@guid="unrelated"]'))

    def test_running_corel_blocks_all_writes(self):
        with patch('corel_export.setup_corel.corel_running', return_value=True):
            with self.assertRaises(RuntimeError):
                install(self.folder, self.folder / 'Draw')
        self.assertFalse((self.folder / 'Draw').exists())

    def test_failure_rolls_back_original_files(self):
        package = self.folder
        (package / 'runtime').mkdir()
        (package / 'runtime/CorelDXF.exe').write_bytes(b'new engine')
        (package / 'SkladCorelDXF.gms').write_bytes(b'new macro')
        draw = self.folder / 'Draw'
        (draw / 'Workspace').mkdir(parents=True)
        original = workspace()
        (draw / 'Workspace/_default.cdws').write_bytes(original)
        (draw / 'GMS/SkladCorelDXF_runtime').mkdir(parents=True)
        (draw / 'GMS/SkladCorelDXF_runtime/CorelDXF.exe').write_bytes(b'old engine')
        (draw / 'GMS/SkladCorelDXF.gms').write_bytes(b'old macro')
        import os
        replace = os.replace

        def fail_workspace(source, destination):
            if Path(source).name == '_default.cdws' and 'SkladDXF-install-' in str(source):
                raise OSError('simulated disk failure')
            return replace(source, destination)

        with patch('corel_export.setup_corel.corel_running', return_value=False), \
                patch('corel_export.setup_corel.os.replace', side_effect=fail_workspace):
            with self.assertRaises(OSError):
                install(package, draw)
        self.assertEqual((draw / 'Workspace/_default.cdws').read_bytes(), original)
        self.assertEqual((draw / 'GMS/SkladCorelDXF.gms').read_bytes(), b'old macro')
        self.assertEqual((draw / 'GMS/SkladCorelDXF_runtime/CorelDXF.exe').read_bytes(), b'old engine')


if __name__ == '__main__':
    unittest.main()
