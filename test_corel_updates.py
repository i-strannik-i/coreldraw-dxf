import hashlib
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
from corel_export import update_manager as u


class UpdateTests(unittest.TestCase):
    def test_automatic_check_is_silent_offline(self):
        with patch.object(u,'check_latest',side_effect=OSError('offline')):
            self.assertIsNone(u.automatic_release('unused'))

    def test_automatic_check_is_silent_when_current(self):
        with patch.object(u,'check_latest',return_value={'available':False}):
            self.assertIsNone(u.automatic_release('unused'))

    def test_automatic_check_offers_newer_release(self):
        release={'available':True,'version':'9.0.0'}
        with patch.object(u,'check_latest',return_value=release), patch.object(u.subprocess,'Popen') as launch:
            self.assertEqual(u.automatic_release('unused'),release)
            launch.assert_not_called()

    def setUp(self):
        self.temp=tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder=Path(self.temp.name)
        self.info={'product':'CorelDXF','corel_major':25,'version':'9.0.0',
                   'file':'SetupCorelDXF-9.0.0.exe','size':4,'sha256':hashlib.sha256(b'test').hexdigest()}
        (self.folder/self.info['file']).write_bytes(b'test')
        self.write()

    def write(self):
        (self.folder/u.MANIFEST).write_text(json.dumps(self.info),encoding='utf-8')

    def test_local_stage(self):
        info=u.check_release(self.folder)
        self.assertTrue(info['available'])
        self.assertEqual(u.stage_release(self.folder,info,self.folder/'cache').read_bytes(),b'test')

    def test_corrupt_installer_rejected(self):
        info=u.check_release(self.folder)
        (self.folder/self.info['file']).write_bytes(b'evil')
        with self.assertRaises(ValueError):u.stage_release(self.folder,info,self.folder/'cache')

    def test_path_traversal_rejected(self):
        self.info['file']='../other.exe';self.write()
        with self.assertRaises(ValueError):u.check_release(self.folder)

    def test_downgrade_not_offered(self):
        self.assertFalse(u.check_release(self.folder,current='10.0.0')['available'])

    def test_github_offline_falls_back(self):
        with patch.object(u,'github_release',side_effect=OSError('offline')):
            self.assertEqual(u.check_latest(self.folder)['version'],'9.0.0')

    def test_github_uses_public_assets_without_api(self):
        with patch.object(u,'read_url',return_value=json.dumps(self.info).encode()) as read:
            result=u.github_release()
        self.assertTrue(result['available'])
        self.assertIn('/download/v9.0.0/',result['_url'])
        self.assertTrue(all('api.github.com' not in call.args[0] for call in read.call_args_list))

    def test_github_changed_manifest_rejected(self):
        changed=dict(self.info,size=5)
        with patch.object(u,'read_url',side_effect=[json.dumps(self.info).encode(),json.dumps(changed).encode()]):
            with self.assertRaises(ValueError):u.github_release()
