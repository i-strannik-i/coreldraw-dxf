"""Compact import settings; no informational confirmation dialogs."""
import math


def parse_tolerance(value):
    number = float(value.strip().replace(',', '.'))
    if not math.isfinite(number) or number <= 0 or number > 1:
        raise ValueError('Введите допуск больше 0 и не больше 1 мм.')
    return number


def choose_options(parent, filename, default_tolerance=0.1):
    import tkinter as tk
    from tkinter import ttk
    from pathlib import Path

    window = tk.Toplevel(parent)
    window.title('Импорт CDR в Aspire')
    window.resizable(False, False)
    window.attributes('-topmost', True)
    mode = tk.StringVar(window, 'arcs')
    tolerance = tk.StringVar(window, str(default_tolerance).replace('.', ','))
    optimize = tk.BooleanVar(window, True)
    error = tk.StringVar(window)
    result = None
    frame = ttk.Frame(window, padding=18)
    frame.grid(sticky='nsew')
    ttk.Label(frame, text='Параметры импорта', font=('Segoe UI', 12, 'bold')).grid(row=0,column=0,columnspan=2,sticky='w')
    ttk.Label(frame, text=Path(filename).name, wraplength=410).grid(row=1,column=0,columnspan=2,sticky='w',pady=(5,14))
    group = ttk.LabelFrame(frame, text='Кривые', padding=10)
    group.grid(row=2,column=0,columnspan=2,sticky='ew')

    def update_mode():
        entry.configure(state='normal' if mode.get()=='arcs' or optimize.get() else 'disabled')
        error.set('')

    ttk.Radiobutton(group, text='Восстанавливать круговые дуги', variable=mode,
                    value='arcs',command=update_mode).grid(row=0,column=0,columnspan=2,sticky='w')
    ttk.Label(group,text='Общий допуск, мм').grid(row=1,column=0,sticky='w',padx=(22,20),pady=8)
    entry = ttk.Spinbox(group,from_=0.001,to=1,increment=0.01,textvariable=tolerance,width=10)
    entry.grid(row=1,column=1,sticky='e')
    ttk.Radiobutton(group, text='Кривые Безье (без замены на дуги)', variable=mode,
                    value='exact',command=update_mode).grid(row=2,column=0,columnspan=2,sticky='w')
    ttk.Checkbutton(group, text='Оптимизировать: объединять соседние сегменты', variable=optimize,
                    command=update_mode).grid(row=3,column=0,columnspan=2,sticky='w',pady=(8,0))
    ttk.Label(group,text='Прямые остаются прямыми. Углы от 30° сохраняются.\nБез оптимизации режим Безье сохраняет исходные узлы.',
              foreground='#555555').grid(row=4,column=0,columnspan=2,sticky='w',pady=(9,0))
    ttk.Label(frame,text='Единицы: мм   •   Слои и цвета: сохранять\nКоординаты: как в CDR, без смещения\nDXF и Aspire: рядом с CDR, с тем же именем.\nПредыдущие файлы: в _cdr_import_backups.',
              foreground='#444444').grid(row=3,column=0,columnspan=2,sticky='w',pady=(12,0))
    ttk.Label(frame,textvariable=error,foreground='#b3261e',wraplength=410).grid(row=4,column=0,columnspan=2,sticky='w',pady=(6,6))

    def accept(event=None):
        nonlocal result
        try:
            value = parse_tolerance(tolerance.get()) if mode.get()=='arcs' or optimize.get() else 0
            result = {'arc': value if mode.get()=='arcs' else 0, 'optimize': value if optimize.get() else 0}
        except ValueError:
            error.set('Введите допуск больше 0 и не больше 1 мм.')
            entry.focus_set()
            return
        window.destroy()

    buttons = ttk.Frame(frame)
    buttons.grid(row=5,column=0,columnspan=2,sticky='e')
    ttk.Button(buttons,text='Импортировать',command=accept).pack(side='left',padx=(0,8))
    ttk.Button(buttons,text='Отмена',command=window.destroy).pack(side='left')
    window.bind('<Return>',accept)
    window.bind('<Escape>',lambda event:window.destroy())
    window.update_idletasks()
    width,height = window.winfo_reqwidth(),window.winfo_reqheight()
    window.geometry(f'+{max(0,(window.winfo_screenwidth()-width)//2)}+{max(0,(window.winfo_screenheight()-height)//2)}')
    window.grab_set()
    window.focus_force()
    parent.wait_window(window)
    return result
