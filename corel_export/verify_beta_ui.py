"""Check beta layout against a live document without changing its vectors."""
import tkinter as tk
from unittest.mock import patch
import vector_tools
import ui_shell

original=tk.Tk
errors=[]
checked=[]


def create_root():
    root=original()
    def inspect():
        try:
            def descendants(widget):
                for child in widget.winfo_children():
                    yield child
                    yield from descendants(child)
            widgets=list(descendants(root))
            journal=next(w for w in widgets if w.winfo_class()=='TLabelframe' and w.cget('text')=='Журнал проверки')
            result_label=next(w for w in journal.winfo_children() if w.winfo_class()=='TLabel')
            root.setvar(result_label.cget('textvariable'),'[12:00:00] '+vector_tools.check_summary(
                {'complete':True,'issues':[]},11,[],.1,True,{'OUT'}))
            notebook=next(w for w in widgets if w.winfo_class()=='TNotebook')
            assert [notebook.tab(tab,'text') for tab in notebook.tabs()]==['Экспорт','Проверка','Слои и помощь']
            root.geometry('1000x760');root.update()
            for theme in ('light','dark'):
                ui_shell.apply_theme(root,theme)
                for tab in notebook.tabs():
                    notebook.select(tab);root.update()
                    if notebook.tab(tab,'text')=='Проверка':
                        assert journal.winfo_ismapped()
                        result_tree=next(w for w in descendants(root) if w.winfo_class()=='Treeview' and w.winfo_ismapped())
                        assert result_tree.winfo_height()>=80, result_tree.winfo_height()
                    for widget in descendants(root):
                        if widget.winfo_class()=='TButton' and widget.winfo_ismapped():
                            assert widget.winfo_rooty()+widget.winfo_height()<=root.winfo_rooty()+root.winfo_height(),widget.cget('text')
            ui_shell.feedback(root)
            feedback=next(w for w in root.winfo_children() if isinstance(w,tk.Toplevel))
            root.update()
            text_fields=[w for w in descendants(feedback) if isinstance(w,tk.Text)]
            assert len(text_fields)==2
            assert all(w.cget('background')==ui_shell.THEMES['dark']['panel'] for w in text_fields)
            feedback.destroy()
            checked.append(True)
        except Exception as error:errors.append(error)
        finally:root.destroy()
    root.after(900,inspect)
    return root


with patch.object(tk,'Tk',create_root),patch.object(ui_shell,'save_theme'):
    vector_tools.main()
assert checked and not errors, errors
print('PASS: three tabs, two themes, visible actions at minimum window size; no geometry changed')
