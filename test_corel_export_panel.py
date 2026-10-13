import tempfile
import time
import tkinter as tk
import unittest
import uuid
import win32event
from pathlib import Path
from unittest.mock import patch

from corel_export import export_panel
from corel_export.export_dxf import GeometryBlocked
from import_progress import ConversionCancelled


class InlineExportTests(unittest.TestCase):
    def setUp(self):
        create_mutex=win32event.CreateMutex
        mutex_name='Local\\CorelDXFTest-'+uuid.uuid4().hex
        mutex_patch=patch.object(win32event,'CreateMutex',lambda attributes,owner,name:create_mutex(attributes,owner,mutex_name))
        mutex_patch.start();self.addCleanup(mutex_patch.stop)
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
        for timer in self.root.tk.call('after','info'):self.root.after_cancel(timer)
        self.root.destroy();self.temp.cleanup()

    def wait(self, decide=None):
        end=time.monotonic()+8
        decided=False
        while self.panel.busy and time.monotonic()<end:
            self.root.update();time.sleep(.01)
            if decide is not None and self.panel.warning_button.winfo_manager() and not decided:
                self.assertEqual(str(self.panel.check_button['state']),'normal')
                (self.panel.warning_button if decide else self.panel.check_button).invoke();decided=True
        self.assertFalse(self.panel.busy)
        self.assertIsNone(self.panel.mutex)
        self.assertFalse(self.source.exists())
        self.assertEqual(self.changes,[True,False])
        self.assertFalse(self.panel.warning_button.winfo_manager())
        self.assertFalse(self.panel.awaiting_confirmation)
        self.assertEqual(str(self.panel.check_button['text']),'Перейти к проверке')

    def test_success_is_logged_inline_without_toplevel(self):
        def convert(*args):return {'check_status':'Проверка пройдена','segments_before':12,'segments_after':8}
        with patch.object(export_panel,'export_request',convert):
            self.panel.start();self.wait()
        self.assertIn('Проверка пройдена',self.panel.text.get('1.0','end'))
        self.assertEqual(self.panel.bar['value'],100)
        self.assertIn('DXF сохранён',self.panel.log_path.read_text(encoding='utf-8-sig'))
        self.assertFalse(any(isinstance(w,tk.Toplevel) for w in self.root.winfo_children()))

    def test_inspection_allowed_only_when_idle_or_waiting_for_confirmation(self):
        self.assertTrue(self.panel.allows_inspection)
        self.panel.preflight=lambda:None
        self.panel.start()
        self.assertFalse(self.panel.allows_inspection)
        self.panel.events.put(('confirm','Пересечение'))
        self.panel.poll()
        self.assertTrue(self.panel.allows_inspection)
        self.panel.choose(True)
        self.assertFalse(self.panel.allows_inspection)
        self.panel.preflight_done(error='Остановлено для теста')
        self.assertTrue(self.panel.allows_inspection)

    def test_shared_preflight_delivers_results_before_conversion(self):
        pending=[];prepared=[]
        self.panel.preflight=lambda:pending.append(True)
        self.panel.prepare=lambda:(prepared.append(True) or str(self.request))
        check={'complete':True,'issues':[]}
        with patch.object(export_panel,'export_request',return_value={'check_status':'Готово'}) as convert:
            self.panel.start()
            self.assertEqual(pending,[True]);self.assertEqual(prepared,[])
            self.panel.preflight_done(check);self.wait()
            self.assertIs(convert.call_args.kwargs['prechecked'],check)

    def test_cancel_preflight_does_not_prepare_or_convert(self):
        self.panel.preflight=lambda:None
        cancelled=[];self.panel.cancel_preflight=lambda:cancelled.append(True)
        with patch.object(self.panel,'prepare') as prepare, patch.object(export_panel,'export_request') as convert:
            self.panel.start();self.panel.cancel();self.panel.preflight_done({'complete':True,'issues':[]})
            prepare.assert_not_called();convert.assert_not_called()
        self.assertFalse(self.panel.busy);self.assertIsNone(self.panel.mutex)
        self.assertEqual(cancelled,[True]);self.assertEqual(self.target.read_bytes(),b'previous DXF')

    def test_blocked_preflight_returns_to_populated_check_without_conversion(self):
        self.panel.preflight=lambda:None
        with patch.object(self.panel,'prepare') as prepare:
            self.panel.start();self.panel.preflight_done(error='Нулевые участки')
            prepare.assert_not_called()
        self.assertEqual(self.checks,[True]);self.assertFalse(self.panel.busy)
        self.assertIsNone(self.panel.mutex)

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
