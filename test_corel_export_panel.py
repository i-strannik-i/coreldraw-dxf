import tempfile
import time
import tkinter as tk
import unittest
from pathlib import Path
from unittest.mock import patch

from corel_export import export_panel
from corel_export.export_dxf import GeometryBlocked
from import_progress import ConversionCancelled


class InlineExportTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory()
        folder=Path(self.temp.name)
        self.source=folder/'snapshot.cdr';self.source.write_bytes(b'fixture')
        self.target=folder/'output.dxf';self.target.write_bytes(b'previous DXF')
        self.request=folder/'request.txt'
        self.request.write_text(f'{self.source}\n{self.target}\nselection\n1\n',encoding='utf-16')
        self.root=tk.Tk();self.root.withdraw()
        self.changes=[];self.checks=[]
        self.panel=export_panel.ExportPanel(self.root,lambda:str(self.request),self.changes.append,lambda:self.checks.append(True))

    def tearDown(self):
        if self.panel.busy:
            self.panel.cancel();self.wait()
        self.root.destroy();self.temp.cleanup()

    def wait(self, decide=None):
        end=time.monotonic()+8
        decided=False
        while self.panel.busy and time.monotonic()<end:
            self.root.update();time.sleep(.01)
            if decide is not None and self.panel.choice.winfo_manager() and not decided:
                self.panel.choose(decide);decided=True
        self.assertFalse(self.panel.busy)
        self.assertIsNone(self.panel.mutex)
        self.assertFalse(self.source.exists())
        self.assertEqual(self.changes,[True,False])

    def test_success_is_logged_inline_without_toplevel(self):
        def convert(*args):return {'check_status':'Проверка пройдена','segments_before':12,'segments_after':8}
        with patch.object(export_panel,'export_request',convert):
            self.panel.start();self.wait()
        self.assertIn('Проверка пройдена',self.panel.text.get('1.0','end'))
        self.assertEqual(self.panel.bar['value'],100)
        self.assertIn('DXF сохранён',self.panel.log_path.read_text(encoding='utf-8-sig'))
        self.assertFalse(any(isinstance(w,tk.Toplevel) for w in self.root.winfo_children()))

    def test_warning_refusal_returns_to_check(self):
        def convert(request,target,progress,confirm):
            if not confirm('CUT_OUT: пересечение'):raise ConversionCancelled('Отмена')
            raise AssertionError('Publication must not be reached')
        with patch.object(export_panel,'export_request',convert):
            self.panel.start();self.wait(decide=False)
        self.assertEqual(self.checks,[True])
        self.assertEqual(self.target.read_bytes(),b'previous DXF')
        self.assertIn('DXF не заменён',self.panel.stage.get())

    def test_warning_acceptance(self):
        def convert(request,target,progress,confirm):
            self.assertTrue(confirm('CUT_OUT: пересечение'))
            return {'check_status':'Экспортировано с предупреждениями'}
        with patch.object(export_panel,'export_request',convert):
            self.panel.start();self.wait(decide=True)
        self.assertIn('Экспортировано с предупреждениями',self.panel.stage.get())
        self.assertEqual(self.checks,[])

    def test_cancellation_while_waiting_for_confirmation(self):
        def convert(request,target,progress,confirm):
            self.assertFalse(confirm('Предупреждение'))
            raise ConversionCancelled('Отмена')
        with patch.object(export_panel,'export_request',convert):
            self.panel.start();self.panel.cancel();self.wait()
        self.assertEqual(self.target.read_bytes(),b'previous DXF')

    def test_blocked_geometry_is_logged(self):
        with patch.object(export_panel,'export_request',side_effect=GeometryBlocked('Нулевые участки. DXF не сохранён.')):
            self.panel.start();self.wait()
        self.assertIn('Нулевые участки',self.panel.text.get('1.0','end'))
        self.assertEqual(self.target.read_bytes(),b'previous DXF')

    def test_preparation_failure_releases_lock(self):
        self.panel.prepare=lambda:(_ for _ in ()).throw(ValueError('Нет выделения'))
        self.panel.start()
        self.assertFalse(self.panel.busy);self.assertIsNone(self.panel.mutex)
        self.assertIn('Нет выделения',self.panel.text.get('1.0','end'))
