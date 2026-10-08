"""Exercise the checker and its confirmation flow using a disposable document."""
import tkinter as tk
from tkinter import messagebox
from unittest.mock import patch
import win32com.client
import vector_tools

app=win32com.client.gencache.EnsureDispatch('CorelDRAW.Application.25')
previous=app.ActiveDocument if app.Documents.Count else None
doc=app.CreateDocument();doc.Unit=3
layer=doc.ActivePage.CreateLayer('CUT_OUT')
first=layer.CreateRectangle2(0,0,30,20);first.Duplicate(0,0);first.Duplicate(0,0)
doc.ClearSelection()
original=tk.Tk
errors=[];completed=[]


def children(widget):
    for child in widget.winfo_children():
        yield child
        yield from children(child)


def create_root():
    root=original();phase=[0]
    def inspect():
        try:
            widgets=list(children(root))
            notebook=next(w for w in widgets if w.winfo_class()=='TNotebook')
            check_tab=notebook.tabs()[1];notebook.select(check_tab)
            buttons={str(w.cget('text')):w for w in widgets if w.winfo_class()=='TButton'}
            transfer=buttons['Перенести дубли']
            if phase[0]==0:
                root.geometry('1000x760');buttons['Проверить'].invoke();phase[0]=1
            elif phase[0]==1 and str(transfer['state'])!='disabled':
                root.update_idletasks()
                assert any(str(w.cget('text')).startswith('Дубли · 2') for w in widgets if w.winfo_class()=='TButton')
                for button in buttons.values():
                    if button.winfo_ismapped():
                        assert button.winfo_rootx()+button.winfo_width()<=root.winfo_rootx()+root.winfo_width()
                        assert button.winfo_rooty()+button.winfo_height()<=root.winfo_rooty()+root.winfo_height()
                with patch.object(messagebox,'askyesno',return_value=False):transfer.invoke()
                assert layer.Shapes.Count==3
                with patch.object(messagebox,'askyesno',return_value=True):transfer.invoke()
                phase[0]=2
            elif phase[0]==2 and layer.Shapes.Count==1 and str(buttons['Проверить']['state'])!='disabled':
                assert str(transfer['state'])=='disabled'
                doc.Undo();assert layer.Shapes.Count==3
                completed.append(True);root.destroy();return
            root.after(150,inspect)
        except Exception as error:errors.append(error);root.destroy()
    root.after(500,inspect)
    root.after(30000,lambda:(errors.append(TimeoutError('Duplicate UI test')),root.destroy()))
    return root


try:
    with patch.object(tk,'Tk',create_root):vector_tools.main()
    assert completed and not errors,errors
    print('PASS: duplicate count, visible actions, refused transfer, confirmed worker transfer and single Undo')
finally:
    doc.Dirty=False;doc.Close()
    if previous is not None:previous.Activate()
