"""Recheck a retained failed-job snapshot without replacing the user's DXF."""
import hashlib
from pathlib import Path
import shutil
import sys
import tempfile
import time
from export_dxf import export_request


def main():
    job=Path(sys.argv[1])
    original=job if job.is_file() else job/'snapshot.cdr'
    before=hashlib.sha256(original.read_bytes()).hexdigest()
    folder=Path(tempfile.mkdtemp(prefix='coreldxf-info-regression-'))
    source=folder/'snapshot.cdr';shutil.copy2(original,source)
    target=folder/'verified.dxf';request=folder/'request.txt'
    request.write_text(f'{source}\n{target}\nselection\n',encoding='utf-16')
    last=[0.]
    def progress(stage,**details):
        if time.monotonic()-last[0]>10:
            print(stage,details,flush=True);last[0]=time.monotonic()
    report=export_request(request,target,progress)
    assert not report['validation']['zero_length']
    assert hashlib.sha256(original.read_bytes()).hexdigest()==before
    import ezdxf
    doc=ezdxf.readfile(target)
    text=list(doc.modelspace().query('MTEXT TEXT'))
    assert text
    print('PASS: INFO text retained, no false zero-length blocker; snapshot unchanged',flush=True)
    print(target,flush=True)


if __name__=='__main__':main()
