"""Shared beta UI styling and explicit, previewed GitHub feedback."""
import json
import os
from pathlib import Path
import re
import tempfile
import tkinter as tk
from tkinter import ttk, messagebox
import urllib.parse
import webbrowser

try:
    from .version import VERSION
except ImportError:
    from version import VERSION

THEMES = {
    'light': dict(bg='#f3f6f7', panel='#ffffff', text='#233342', muted='#586977', line='#d6e0e5', accent='#087f8c', selected='#dceff2',
                  zero='#c5363d', overlap='#7843a3', intersection='#a65300', open='#006a9e', duplicate='#27733a'),
    'dark': dict(bg='#20272e', panel='#29333d', text='#edf3f7', muted='#b5c1cb', line='#455461', accent='#54c5ce', selected='#34515e',
                 zero='#ff858b', overlap='#cda0ff', intersection='#ffbd70', open='#78c8ff', duplicate='#80d596'),
}


def settings_path():
    return Path(os.environ.get('LOCALAPPDATA', tempfile.gettempdir())) / 'SkladCorelDXF/ui-settings.json'


def load_theme():
    try:
        value = json.loads(settings_path().read_text(encoding='utf-8'))['theme']
        return value if value in THEMES else 'light'
    except (OSError, ValueError, KeyError):
        return 'light'


def apply_theme(root, name, tree=None, canvas=None):
    p = THEMES[name]
    style = ttk.Style(root)
    style.theme_use('clam')
    root.configure(background=p['bg'])
    style.configure('.', background=p['bg'], foreground=p['text'], font=('Segoe UI', 10), bordercolor=p['line'], lightcolor=p['line'], darkcolor=p['line'])
    style.configure('TFrame', background=p['bg'])
    style.configure('TLabel', background=p['bg'], foreground=p['text'])
    style.configure('Muted.TLabel', foreground=p['muted'])
    style.configure('Title.TLabel', font=('Segoe UI', 18, 'bold'))
    style.configure('TButton', padding=(10, 7), background=p['panel'], foreground=p['text'], borderwidth=1)
    style.map('TButton', background=[('active', p['selected'])], foreground=[('disabled',p['muted'])])
    style.configure('Accent.TButton', background=p['accent'], foreground='#ffffff' if name=='light' else '#10242a')
    style.map('Accent.TButton',background=[('active',p['accent'])],foreground=[('disabled',p['muted']),('active','#ffffff' if name=='light' else '#10242a')])
    style.configure('TCheckbutton', padding=4)
    style.configure('Horizontal.TProgressbar', background=p['accent'], troughcolor=p['panel'], bordercolor=p['line'], lightcolor=p['accent'], darkcolor=p['accent'])
    style.map('TCheckbutton', background=[('active',p['selected'])])
    style.configure('TEntry', fieldbackground=p['panel'], foreground=p['text'], padding=5, insertcolor=p['text'])
    style.configure('TCombobox', fieldbackground=p['panel'], foreground=p['text'], padding=5)
    style.map('TCombobox', fieldbackground=[('readonly',p['panel'])], foreground=[('readonly',p['text'])])
    style.configure('TNotebook', background=p['bg'], borderwidth=0)
    style.configure('TNotebook.Tab', padding=(20,10), background=p['bg'])
    style.map('TNotebook.Tab', background=[('selected',p['panel'])], foreground=[('selected',p['accent'])])
    style.configure('Treeview', background=p['panel'], fieldbackground=p['panel'], foreground=p['text'], rowheight=30, borderwidth=1)
    style.configure('Treeview.Heading', background=p['bg'], foreground=p['muted'], padding=8)
    style.map('Treeview', background=[('selected',p['selected'])], foreground=[('selected',p['text'])])
    for kind in ('zero','overlap','intersection','open','duplicate'):
        style.configure(kind+'.TButton', foreground=p[kind])
        if tree is not None:
            tree.tag_configure(kind, foreground=p[kind])
    if canvas is not None:
        canvas.configure(background=p['bg'])
    return p


def save_theme(name):
    path=settings_path();path.parent.mkdir(parents=True,exist_ok=True)
    path.write_text(json.dumps({'theme':name}),encoding='utf-8')


def clean_log(text):
    # Withhold whole sensitive lines rather than trying to guess path boundaries.
    result=[]
    for line in text.splitlines():
        if re.search(r'(?i)([a-z]:[\\/]|\\\\|https?://|\b(?:token|password|secret|authorization)\b|\.(?:cdr|dxf|crv3d)\b|[\w.+-]+@[\w.-]+)',line):
            result.append('[строка с путём, именем файла или личными данными скрыта]')
        else:
            result.append(line)
    return '\n'.join(result)[-2500:]


def latest_log():
    root=Path(tempfile.gettempdir())/'SkladCorelDXF/jobs'
    files=list(root.glob('*/error.txt'))+list(root.glob('*/operations.log'))
    if not files:
        return 'Журнал не найден.'
    path=max(files,key=lambda p:p.stat().st_mtime)
    return clean_log(path.read_text(encoding='utf-8-sig',errors='replace'))


def feedback(root):
    win=tk.Toplevel(root);win.title('Обратная связь · GitHub');win.geometry('660x610');win.minsize(600,540)
    style=ttk.Style(root)
    text_style=dict(background=style.lookup('TEntry','fieldbackground'),foreground=style.lookup('TEntry','foreground'),
                    insertbackground=style.lookup('TEntry','foreground'),relief='solid',borderwidth=1,padx=8,pady=6)
    frame=ttk.Frame(win,padding=18);frame.pack(fill='both',expand=True)
    ttk.Label(frame,text='Жалоба или предложение',style='Title.TLabel').pack(anchor='w')
    ttk.Label(frame,text='Обращение будет публичным. Не указывайте данные заказчиков.\nОкончательная публикация — на GitHub под вашей учётной записью.',style='Muted.TLabel').pack(anchor='w',pady=10)
    kind=tk.StringVar(value='Ошибка')
    ttk.Combobox(frame,textvariable=kind,values=('Ошибка','Предложение'),state='readonly',width=20).pack(anchor='w')
    title=tk.StringVar()
    ttk.Label(frame,text='Краткий заголовок').pack(anchor='w',pady=(10,3))
    ttk.Entry(frame,textvariable=title).pack(fill='x')
    ttk.Label(frame,text='Что произошло / что предлагаете?').pack(anchor='w',pady=(10,3))
    text=tk.Text(frame,height=6,wrap='word',font=('Segoe UI',10),**text_style);text.pack(fill='both',expand=True)
    include=tk.BooleanVar(value=False)
    preview=tk.Text(frame,height=6,wrap='word',font=('Consolas',9),**text_style)
    def toggle():
        preview.delete('1.0','end')
        if include.get():
            try:preview.insert('1.0',latest_log())
            except OSError:preview.insert('1.0','Журнал недоступен.')
    ttk.Checkbutton(frame,text='Приложить очищенный журнал — проверьте текст ниже',variable=include,command=toggle).pack(anchor='w',pady=10)
    preview.pack(fill='both',expand=True)
    def send():
        body=text.get('1.0','end').strip()
        if not title.get().strip() or not body:
            messagebox.showwarning('Обратная связь','Заполните заголовок и сообщение.',parent=win);return
        if len(body)>1800 or len(title.get())>120:
            messagebox.showwarning('Обратная связь','Сократите заголовок до 120, а сообщение до 1800 символов.',parent=win);return
        body=f'Версия: {VERSION} beta\nТип: {kind.get()}\n\n'+body
        if include.get():body+='\n\nЖурнал (проверьте перед публикацией):\n```text\n'+clean_log(preview.get('1.0','end'))+'\n```'
        url='https://github.com/i-strannik-i/coreldraw-dxf/issues/new?'+urllib.parse.urlencode({'title':title.get(),'body':body})
        if len(url)>8000:
            messagebox.showwarning('Обратная связь','Сократите журнал или снимите галочку вложения.',parent=win);return
        webbrowser.open(url)
    ttk.Button(frame,text='Просмотреть и отправить на GitHub',command=send,style='Accent.TButton').pack(anchor='e',pady=(14,0))
