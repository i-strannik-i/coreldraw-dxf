"""Prepare an isolated two-layer document for installed scope controls."""
from pathlib import Path
import os
import subprocess
import win32com.client

app=win32com.client.gencache.EnsureDispatch('CorelDRAW.Application.25')
doc=app.CreateDocument();doc.Unit=3
doc.ActivePage.CreateLayer('CUT_A').CreateLineSegment(10,10,20,10)
doc.ActivePage.CreateLayer('CUT_B').CreateLineSegment(40,10,50,10)
doc.ActivePage.CreateLayer('INFO').CreateArtisticText(10,40,'ignored')
doc.ActivePage.CreateLayer('#').CreateLineSegment(60,10,70,10)
hidden=doc.ActivePage.CreateLayer('HIDDEN');hidden.CreateLineSegment(80,10,90,10);hidden.Visible=False
doc.ClearSelection()
folder=Path(__file__).resolve().parent/'build/scope-test';folder.mkdir(parents=True,exist_ok=True)
doc.SaveAs(str(folder/'scope-test.cdr'),None)
runtime=Path(os.environ['APPDATA'])/'Corel/CorelDRAW Graphics Suite 2024/Draw/GMS/SkladCorelDXF_runtime/VectorTools.exe'
subprocess.Popen([str(runtime),'--check'],creationflags=subprocess.CREATE_NO_WINDOW)
print('Prepared no-selection scope test: two visible working contours, INFO/#/hidden excluded.')
