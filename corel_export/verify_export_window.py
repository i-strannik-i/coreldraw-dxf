"""Manual UI smoke test: long error text, no conversion or user-file writes."""
from pathlib import Path
import sys
import tempfile
from unittest.mock import patch
import export_dxf


with tempfile.TemporaryDirectory(prefix='coreldxf-window-test-') as directory:
    folder=Path(directory);source=folder/'snapshot.cdr';source.touch()
    request=folder/'request.txt'
    request.write_text(f'{source}\n{folder / "TEST-long-export-message.dxf"}\nselection\n',encoding='utf-16')
    message='ТЕСТ ИНТЕРФЕЙСА. Экспорт не выполнялся, рабочие файлы не затронуты.\n'+('Длинное диагностическое сообщение для проверки прокрутки. '*18)
    with patch.object(sys,'argv',['export_dxf.py',str(request)]),patch.object(export_dxf,'export_request',side_effect=ValueError(message)):
        export_dxf.main()
