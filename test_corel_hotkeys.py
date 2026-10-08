import io
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
import zipfile
import xml.etree.ElementTree as ET

from corel_export import hotkey_settings as h


def archive():
    xml = f'''<uiConfig><items><itemData guid="{h.EXPORT}"/><itemData guid="save" userCaption="Save"/></items>
    <shortcutKeyTables><table tableID="{h.MAIN}">
    <keySequence itemRef="save"><key ctrl="true">s</key></keySequence>
    <keySequence itemRef="{h.EXPORT}"><key ctrl="true">d</key></keySequence>
    <keySequence itemRef="chain"><key ctrl="true">k</key><key>c</key></keySequence>
    </table><table tableID="text"><keySequence itemRef="text"><key ctrl="true">s</key></keySequence></table></shortcutKeyTables>
    <unrelated a="1" /></uiConfig>'''
    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w') as z:
        z.writestr('content/workspace.xml', xml)
        z.writestr('untouched.bin', b'original')
    return out.getvalue()


class HotkeyTests(unittest.TestCase):
    def test_numlock_does_not_turn_ctrl_into_alt(self):
        self.assertEqual(h.event_key(83, 4 | 8), ('S', True, False, False))
        self.assertEqual(h.event_key(123, 4 | 1 | 0x20000), ('VK_F12', True, True, True))

    def test_conflict_and_caption(self):
        xml = h.workspace_xml(archive())
        self.assertEqual(h.command_names(xml, ('S', True, False, False), {}), ['Save'])
        self.assertEqual(h.assignments(xml, ('X', True, False, False)), ())

    def test_change_one_key_only_and_preserve_context(self):
        data = archive()
        key = ('S', True, False, False)
        result = h.patch_workspace(data, key, h.assignments(h.workspace_xml(data), key))
        xml = h.workspace_xml(result)
        root, table = h.main_table(xml)
        own = table.findall(f'keySequence[@itemRef="{h.EXPORT}"]')
        self.assertEqual(len(own), 1)
        self.assertEqual(h.key_tuple(own[0].find('key')), key)
        self.assertIsNotNone(root.find('./shortcutKeyTables/table[@tableID="text"]/keySequence[@itemRef="text"]'))
        self.assertIn('<unrelated a="1" />', xml)
        self.assertEqual(zipfile.ZipFile(io.BytesIO(result)).read('untouched.bin'), b'original')

    def test_stale_confirmation_rejected(self):
        with self.assertRaises(ValueError):
            h.patch_workspace(archive(), ('S', True, False, False), ())

    def test_prefix_not_free(self):
        self.assertTrue(h.assignments(h.workspace_xml(archive()), ('K', True, False, False)))

    def test_missing_table_rejected(self):
        with self.assertRaises(ValueError):
            h.main_table('<uiConfig/>')

    def test_running_corel_blocks_writes(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / 'test.cdws'
            path.write_bytes(archive())
            with patch.object(h, 'corel_running', return_value=True), self.assertRaises(ValueError):
                h.apply_closed(path, ('X', True, False, False), ())
            self.assertEqual(path.read_bytes(), archive())

    def test_closed_write_has_backup(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / 'Workspace' / 'test.cdws'
            path.parent.mkdir()
            data = archive()
            path.write_bytes(data)
            with patch.object(h, 'corel_running', return_value=False):
                backup = h.apply_closed(path, ('X', True, False, False), ())
            self.assertEqual((backup / path.name).read_bytes(), data)
            self.assertEqual(h.command_names(h.workspace_xml(path.read_bytes()), ('X', True, False, False), {}), ['Экспорт DXF'])


if __name__ == '__main__':
    unittest.main()
