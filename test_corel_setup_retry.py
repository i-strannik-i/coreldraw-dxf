"""Shutdown races and retry in the same installer window."""
from pathlib import Path
import gc
import tempfile
import threading
import time
import tkinter as tk
from tkinter import ttk
import unittest
from unittest.mock import Mock, patch

from corel_export import setup_corel as setup


class ShutdownTests(unittest.TestCase):
    def test_reopened_corel_restarts_shutdown_wait(self):
        cancelled = Mock()
        cancelled.is_set.return_value = False
        cancelled.wait.return_value = False
        status = Mock()
        with patch.object(setup, 'corel_running', side_effect=[True, False, True, False, False, False]):
            self.assertTrue(setup.wait_for_corel_exit(cancelled, status))
        self.assertEqual(sum('продолжится автоматически' in call.args[0] for call in status.call_args_list), 2)

    def test_cancel_stops_wait_without_touching_corel(self):
        cancelled = threading.Event()
        def running():
            cancelled.set()
            return True
        with patch.object(setup, 'corel_running', side_effect=running):
            self.assertFalse(setup.wait_for_corel_exit(cancelled, Mock(), interval=0))

    def exercise_window(self, cancel_first=False, fail_first_probe=False):
        # Tk objects must be collected on their owner thread, including cycles
        # retained by callbacks from a previous test window.
        gc.collect()
        was_enabled = gc.isenabled()
        gc.disable()
        self.addCleanup(gc.collect)
        if was_enabled:
            self.addCleanup(gc.enable)
        with tempfile.TemporaryDirectory() as folder:
            folder = Path(folder)
            state = {'running': True}
            preflight = {'ok': True, 'workspace_path': str(folder / 'Workspace'), 'gms_path': str(folder / 'GMS')}
            results = [preflight, {'ok': True}]
            if fail_first_probe:
                results.insert(0, {'ok': False, 'kind': 'corel', 'message': 'Corel ещё завершает работу.'})
            probe = Mock(side_effect=results)
            install = Mock(return_value=folder / 'backup')

            def drive(root):
                root.withdraw()
                def descendants(widget):
                    for child in widget.winfo_children():
                        yield child
                        yield from descendants(child)
                widgets = list(descendants(root))
                button = next(w for w in widgets if isinstance(w, ttk.Button) and w.cget('text') == 'Установить')
                def until(predicate):
                    end = time.monotonic() + 6
                    while time.monotonic() < end:
                        root.update()
                        if predicate():
                            return
                        time.sleep(.01)
                    self.fail('Installer did not reach expected state: ' + str(button.cget('text')))
                def waiting_message():
                    return any(isinstance(w, ttk.Label) and 'продолжится автоматически' in str(w.cget('text')) for w in widgets)
                try:
                    button.invoke()
                    until(waiting_message)
                    probe.assert_not_called()
                    install.assert_not_called()
                    if cancel_first:
                        button.invoke()
                        until(lambda: button.cget('text') == 'Проверить и установить')
                        install.assert_not_called()
                        state['running'] = False
                        button.invoke()
                    else:
                        state['running'] = False
                    if fail_first_probe:
                        until(lambda: button.cget('text') == 'Проверить и установить')
                        install.assert_not_called()
                        button.invoke()
                    until(lambda: button.cget('text') == 'Открыть CorelDRAW')
                    install.assert_called_once()
                    self.assertEqual(probe.call_args.kwargs['major'], 27)
                finally:
                    for timer in root.tk.call('after', 'info'):
                        root.after_cancel(timer)
                    root.destroy()

            with patch.object(setup, 'discover', return_value=[27]), \
                    patch.object(setup, 'corel_executable', return_value=folder / 'CorelDRW.exe'), \
                    patch.object(setup, 'corel_running', side_effect=lambda: state['running']), \
                    patch.object(setup, 'run_probe', probe), patch.object(setup, 'install', install), \
                    patch.object(setup.zipfile, 'ZipFile') as archive, \
                    patch.object(tk.Tk, 'mainloop', drive):
                archive.return_value.__enter__.return_value.namelist.return_value = []
                setup.main()

    def test_closing_corel_continues_in_original_window(self):
        self.exercise_window()

    def test_cancel_then_retry_in_original_window(self):
        self.exercise_window(cancel_first=True)

    def test_failed_probe_then_retry_in_original_window(self):
        self.exercise_window(fail_first_probe=True)


if __name__ == '__main__':
    unittest.main()
