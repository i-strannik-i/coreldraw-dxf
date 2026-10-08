"""Exercise automatic-update UI without downloads or installer launches."""
import sys
import tkinter as tk
from unittest.mock import patch
import update_manager as updater


for scenario in ('current', 'offline', 'new'):
    original = tk.Tk
    failures = []
    checked = []

    def make_root():
        root = original()
        def verify():
            try:
                assert scenario == 'new', 'Silent check left an unwanted window'
                root.update_idletasks()
                assert root.state() == 'normal'
                def children(widget):
                    for child in widget.winfo_children():
                        yield child
                        yield from children(child)
                buttons = [w for w in children(root) if w.winfo_class() == 'TButton']
                assert any(w.cget('text') == 'Позже' for w in buttons)
                assert any(w.cget('text') == 'Обновить' and not w.instate(['disabled']) for w in buttons)
                checked.append(True)
            except Exception as error:
                failures.append(error)
            finally:
                root.destroy()
        timer = root.after(600, verify)
        destroy = root.destroy
        def close():
            root.after_cancel(timer)
            destroy()
        root.destroy = close
        return root

    result = {'available': scenario == 'new', 'version': '99.0.0'}
    with patch.object(tk, 'Tk', make_root), patch.object(sys, 'argv', ['updater', '--auto']), \
            patch.object(updater, 'check_latest', side_effect=OSError('offline') if scenario == 'offline' else None,
                         return_value=result), patch.object(updater.subprocess, 'Popen') as launch:
        updater.main()
        launch.assert_not_called()
    assert not failures, failures
    assert bool(checked) == (scenario == 'new')
    print('PASS:', scenario)
