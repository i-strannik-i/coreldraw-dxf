"""Bounded Tk smoke test; no Corel calls or production files are used."""
from pathlib import Path
import sys
import tempfile
import tkinter as tk
from unittest.mock import patch

import export_dxf


def verify():
    original_tk = tk.Tk
    errors = []
    checked = []
    with tempfile.TemporaryDirectory(prefix='coreldxf-journal-test-') as directory:
        folder = Path(directory)
        source = folder / 'snapshot.cdr'
        source.touch()
        request = folder / 'request.txt'
        request.write_text(f'{source}\n{folder / "test.dxf"}\npage\n', encoding='utf-16')

        def conversion(request, target, progress, confirm=None):
            progress('Чтение объектов', layer='CUT_OUT', contours=3)
            progress('Оптимизация узлов', layer='CUT_OUT', segments=100)
            raise export_dxf.GeometryBlocked('CUT_OUT: 2. Нажмите «Проверить». DXF не заменён.')

        def root_factory():
            root = original_tk()

            def inspect():
                try:
                    root.geometry('640x400')
                    root.update()
                    def descendants(widget):
                        for child in widget.winfo_children():
                            yield child
                            yield from descendants(child)
                    widgets = list(descendants(root))
                    buttons = [w for w in widgets if w.winfo_class() == 'TButton' and w.winfo_ismapped()]
                    assert {w.cget('text') for w in buttons} == {'Открыть журнал', 'Закрыть'}
                    for button in buttons:
                        assert button.winfo_rooty() + button.winfo_height() <= root.winfo_rooty() + root.winfo_height(), (button.cget('text'), button.winfo_rooty(), button.winfo_height(), root.winfo_rooty(), root.winfo_height())
                    text = next(w for w in widgets if isinstance(w, tk.Text)).get('1.0', 'end')
                    for expected in ('Оптимизация:', 'Слой: CUT_OUT', 'Сегментов: 100', 'Нажмите «Проверить»'):
                        assert expected in text, expected
                    assert text.strip() == (folder / 'operations.log').read_text(encoding='utf-8-sig').strip()
                    assert not (folder / 'test.dxf').exists()
                    checked.append(True)
                except Exception as error:
                    errors.append(error)
                finally:
                    root.destroy()

            root.after(1500, inspect)
            return root

        with patch.object(tk, 'Tk', root_factory), patch.object(export_dxf, 'export_request', conversion), \
                patch.object(sys, 'argv', ['export_dxf.py', str(request)]):
            export_dxf.main()
        if errors:
            raise errors[0]
        assert checked
    print('PASS: journal history, layer labels, blocking error, no extra button, footer at 640x400')


if __name__ == '__main__':
    verify()
