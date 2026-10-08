"""Assign a native Corel shortcut while its workspace is safely closed."""
from pathlib import Path
import ctypes
import datetime
import io
import json
import os
import re
import subprocess
import sys
import tempfile
import threading
import time
import xml.etree.ElementTree as ET
import zipfile

EXPORT = '87260fe1-e755-4c95-9f68-2aa2f3dc5001'
MAIN = 'bc175625-191c-4b95-9053-756e5eee26fe'


def key_tuple(node):
    text = (node.text or '').upper()
    if text.startswith('VK_') and len(text) == 4:
        text = text[3:]
    return (text, node.get('ctrl') == 'true', node.get('alt') == 'true',
            node.get('shift') == 'true')


def key_label(key):
    name, ctrl, alt, shift = key
    return '+'.join(([ 'Ctrl'] if ctrl else []) + (['Alt'] if alt else [])
                    + (['Shift'] if shift else []) + [name.removeprefix('VK_')])


def event_key(vk, state):
    # Windows Tk uses 0x20000 for Alt; 0x8 can be a lock-state bit.
    return (f'VK_F{vk - 111}' if vk >= 112 else chr(vk),
            bool(state & 4), bool(state & 0x20000), bool(state & 1))


def workspace_xml(data):
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return archive.read('content/workspace.xml').decode('utf-8')


def main_table(xml):
    root = ET.fromstring(xml)
    table = root.find(f'./shortcutKeyTables/table[@tableID="{MAIN}"]')
    if table is None:
        raise ValueError('Не найдена основная таблица горячих клавиш Corel.')
    if root.find(f'./items/itemData[@guid="{EXPORT}"]') is None:
        raise ValueError('В этой рабочей среде не установлена команда Sklad DXF.')
    return root, table


def assignments(xml, key):
    _, table = main_table(xml)
    result = []
    for seq in table.findall('keySequence'):
        keys = seq.findall('key')
        # Prefixes of multi-key sequences are also conflicts, not free keys.
        if keys and key_tuple(keys[0]) == key:
            result.append(ET.tostring(seq, encoding='unicode'))
    return tuple(sorted(result))


def command_names(xml, key, captions):
    root, _ = main_table(xml)
    items = {n.get('guid'): n for n in root.findall('./items/itemData')}
    names = []
    for text in assignments(xml, key):
        guid = ET.fromstring(text).get('itemRef', '')
        item = items.get(guid)
        caption = '' if item is None else (item.get('userCaption') or item.get('dynamicCommand') or '')
        names.append('Экспорт DXF' if guid == EXPORT else caption or captions.get(guid) or f'Команда {guid}')
    return list(dict.fromkeys(names))


def patch_workspace(data, key, expected):
    xml = workspace_xml(data)
    if assignments(xml, key) != expected:
        raise ValueError('Назначение изменилось после проверки. Проверьте сочетание заново.')
    main_table(xml)
    pattern = r'(<table\b(?=[^>]*\btableID="' + MAIN + r'")[^>]*>)(.*?)(</table>)'
    match = re.search(pattern, xml, re.S)
    if not match:
        raise ValueError('Не удалось безопасно найти таблицу клавиш.')

    def retain(m):
        seq = ET.fromstring(m.group())
        keys = seq.findall('key')
        if seq.get('itemRef') == EXPORT or (keys and key_tuple(keys[0]) == key):
            return ''
        return m.group()

    body = re.sub(r'<keySequence\b[^>]*>.*?</keySequence>', retain, match[2], flags=re.S)
    name, ctrl, alt, shift = key
    attrs = ''.join(f' {n}="true"' for n, value in [('ctrl', ctrl), ('alt', alt), ('shift', shift)] if value)
    text = name.lower() if len(name) == 1 else name
    body += f'<keySequence itemRef="{EXPORT}"><key{attrs}>{text}</key></keySequence>'
    xml = xml[:match.start(2)] + body + xml[match.end(2):]
    ET.fromstring(xml)
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as source, zipfile.ZipFile(out, 'w') as target:
        target.comment = source.comment
        for entry in source.infolist():
            target.writestr(entry, xml.encode('utf-8') if entry.filename == 'content/workspace.xml' else source.read(entry))
    return out.getvalue()


def corel_running():
    result = subprocess.run(['tasklist', '/FI', 'IMAGENAME eq CorelDRW.exe', '/FO', 'CSV', '/NH'],
                            capture_output=True, creationflags=0x08000000, timeout=10, check=True)
    return b'coreldrw.exe' in result.stdout.lower()


def apply_closed(path, key, expected):
    if corel_running():
        raise ValueError('Сначала закройте CorelDRAW. Документы автоматически не закрываются.')
    data = path.read_bytes()
    patched = patch_workspace(data, key, expected)
    backup = path.parent.parent / 'SkladDXF_backups' / datetime.datetime.now().strftime('hotkey_%Y%m%d_%H%M%S_%f')
    backup.mkdir(parents=True)
    (backup / path.name).write_bytes(data)
    fd, temporary = tempfile.mkstemp(prefix='.sklad-hotkey-', suffix='.tmp', dir=path.parent)
    try:
        with os.fdopen(fd, 'wb') as stream:
            stream.write(patched)
        if corel_running() or path.read_bytes() != data:
            raise ValueError('Corel открыт или настройки изменились. Повторите сохранение.')
        os.replace(temporary, path)
    finally:
        Path(temporary).unlink(missing_ok=True)
    return backup


def run():
    import tkinter as tk
    from tkinter import ttk
    import win32com.client
    if not corel_running():
        raise ValueError('Откройте CorelDRAW и запустите настройки из окна DXF.')
    app = win32com.client.Dispatch('CorelDRAW.Application.25')
    path = Path(app.UserWorkspacePath) / (app.ActiveWorkspace.Name + '.cdws')
    captions = {}
    del app
    root = tk.Tk()
    root.title('Горячая клавиша DXF')
    root.geometry('560x350')
    root.resizable(False, False)
    frame = ttk.Frame(root, padding=22)
    frame.pack(fill='both', expand=True)
    ttk.Label(frame, text='Открывать экспорт сочетанием клавиш', font=('Segoe UI', 13, 'bold')).pack(anchor='w')
    ttk.Label(frame, text='Нажмите в поле и задайте сочетание, например Ctrl+Shift+D.').pack(anchor='w', pady=(10, 8))
    key_text = tk.StringVar(value='Нажмите сочетание…')
    field = ttk.Entry(frame, textvariable=key_text, font=('Segoe UI', 16), justify='center')
    field.pack(fill='x', ipady=8)
    status = tk.StringVar(value='Проверяется основная таблица Corel, вне редактирования текста.')
    ttk.Label(frame, textvariable=status, wraplength=510).pack(anchor='w', pady=12)
    notice = tk.StringVar(value='Назначение применяется сразу в CorelDRAW.\nЗанятое сочетание заменяется только по кнопке «Переназначить на DXF».')
    ttk.Label(frame, textvariable=notice, foreground='#58666b', wraplength=510).pack(anchor='w')
    row = ttk.Frame(frame)
    row.pack(fill='x', pady=(18, 0))
    state = {'key': None, 'expected': None, 'waiting': False, 'job': None, 'native': None, 'vk': 0, 'mod': 0, 'generation': 0}

    def bridge(method, *args):
        if not corel_running():
            raise ValueError('CorelDRAW закрыт.')
        result = []
        def invoke():
            import pythoncom
            pythoncom.CoInitialize()
            try:
                corel = win32com.client.Dispatch('CorelDRAW.Application.25')
                handle = corel.FrameWork.MainFrameWindow.Handle
                process_id = ctypes.c_ulong()
                ctypes.windll.user32.GetWindowThreadProcessId(handle, ctypes.byref(process_id))
                ctypes.windll.user32.AllowSetForegroundWindow(process_id.value)
                corel.FrameWork.MainFrameWindow.Activate()
                ctypes.windll.user32.SetForegroundWindow(handle)
                deadline = time.monotonic() + 2
                while ctypes.windll.user32.GetForegroundWindow() != handle:
                    if time.monotonic() >= deadline:
                        raise ValueError('Не удалось активировать Corel для проверки. Повторите сочетание.')
                    time.sleep(.05)
                # Let activation messages update Corel's input queue before
                # the VBA bridge temporarily supplies modifier state.
                time.sleep(.1)
                result.append(str(corel.GMSManager.RunMacro('SkladCorelDXF', 'HotkeyBridge.' + method, *args)))
                del corel
            except Exception as exc:
                result.append(exc)
            finally:
                pythoncom.CoUninitialize()
        # Keep Corel's COM call separate from Tk's keyboard input queue.
        worker = threading.Thread(target=invoke)
        worker.start()
        worker.join()
        root.lift()
        root.focus_force()
        field.focus_set()
        if isinstance(result[0], Exception):
            raise result[0]
        return result[0]

    def cancel_wait():
        if state['job']:
            root.after_cancel(state['job'])
            state['job'] = None
        state['waiting'] = False

    def choose_other():
        cancel_wait()
        state['key'] = None
        state['native'] = None
        state['generation'] += 1
        key_text.set('Нажмите другое сочетание…')
        save.configure(state='disabled', text='Сохранить')
        status.set('Существующее назначение не изменено.')
        field.configure(state='normal')
        field.focus_set()

    def capture(event):
        if state['waiting']:
            return 'break'
        vk = event.keycode
        if vk in (16, 17, 18):
            return 'break'
        state['generation'] += 1
        generation = state['generation']
        state['key'] = None
        save.configure(state='disabled')
        if not (48 <= vk <= 57 or 65 <= vk <= 90 or 112 <= vk <= 123):
            status.set('Используйте букву, цифру или F1–F12 с Ctrl / Alt / Shift.')
            state['key'] = None
            save.configure(state='disabled')
            return 'break'
        down = lambda k: bool(ctypes.windll.user32.GetKeyState(k) & 0x8000)
        # Use the queued event's modifier snapshot, not keys already released
        # by the time Tk processes this event.
        key = event_key(vk, event.state)
        if (not key[1] and not key[2] and vk < 112) or (key[2] and vk == 115) or down(91) or down(92):
            status.set('Выберите сочетание с Ctrl или Alt. Alt+F4 и Windows не назначаются.')
            state['key'] = None
            save.configure(state='disabled')
            return 'break'
        key_text.set(key_label(key))
        status.set('Отпустите клавиши для проверки…')
        def after_release():
            if generation != state['generation']:
                return
            if any(ctypes.windll.user32.GetAsyncKeyState(k) & 0x8000 for k in (16, 17, 18, vk)):
                root.after(80, after_release)
            else:
                root.after(80, lambda: check_key(key, vk, generation))
        root.after(80, after_release)
        return 'break'

    def check_key(key, vk, generation):
        if generation != state['generation']:
            return
        try:
            modifiers = int(key[3]) + 2 * int(key[1]) + 4 * int(key[2])
            preview = bridge('Preview', vk, modifiers)
            if preview.startswith('OK\t'):
                parts = preview.split('\t')
                if parts[1].lower() != key_label(key).lower():
                    raise ValueError('Corel распознал ' + parts[1] + ' вместо ' + key_label(key) + '. Попробуйте заново.')
                state.update(key=key, native=preview, vk=vk, mod=modifiers)
                key_text.set(parts[1])
                names = parts[2] if len(parts) > 2 else ''
                status.set('Занято: ' + names + '. Заменить назначение или выбрать другое?' if names else 'Свободно. Можно назначить для открытия DXF.')
                notice.set('Применяется сразу, без перезапуска CorelDRAW.\nПеред сохранением назначение проверяется повторно.')
                save.configure(state='normal', text='Переназначить на DXF' if names else 'Сохранить')
                return 'break'
            state['native'] = None
            notice.set('Штатная запись сейчас недоступна. Запасной способ:\nсохранение после закрытия CorelDRAW, с резервной копией.')
            current = workspace_xml(path.read_bytes())
            state['expected'] = assignments(current, key)
            for entry in state['expected']:
                guid = ET.fromstring(entry).get('itemRef')
                try:
                    corel = win32com.client.Dispatch('CorelDRAW.Application.25')
                    captions[guid] = corel.FrameWork.Automation.GetCaptionText(guid).replace('&', '')
                except Exception:
                    pass
            names = command_names(current, key, captions)
            state['key'] = key
            key_text.set(key_label(key))
            status.set('Занято: ' + ', '.join(names) + '. Заменить назначение или выбрать другое?' if names else 'Сочетание свободно в сохранённой рабочей среде Corel.')
            save.configure(state='normal', text='Переназначить на DXF' if names else 'Сохранить')
        except Exception as exc:
            state['key'] = None
            save.configure(state='disabled')
            status.set('Не удалось проверить: ' + str(exc))
        return 'break'

    def poll():
        state['job'] = None
        try:
            if corel_running():
                state['job'] = root.after(1500, poll)
                return
            backup = apply_closed(path, state['key'], state['expected'])
            status.set('Сохранено: ' + key_label(state['key']) + '. Можно запускать CorelDRAW.\nРезервная копия: ' + str(backup))
            state['waiting'] = False
            other.configure(state='disabled')
        except Exception as exc:
            cancel_wait()
            status.set(str(exc) + '\nНичего не переназначено. Выберите сочетание заново.')
            field.configure(state='normal')
            state['key'] = None

    def save_key():
        if state['key'] is None:
            return
        if state['native']:
            try:
                # Keep a record of the live conflict as well as the last saved workspace.
                backup = path.parent.parent / 'SkladDXF_backups' / datetime.datetime.now().strftime('hotkey_live_%Y%m%d_%H%M%S_%f')
                backup.mkdir(parents=True)
                (backup / path.name).write_bytes(path.read_bytes())
                (backup / 'assignment.json').write_text(json.dumps({'key': key_label(state['key']), 'before': state['native']}, ensure_ascii=False), encoding='utf-8')
                result = bridge('Assign', state['vk'], state['mod'], state['native'])
                if not result.startswith('OK\t'):
                    raise ValueError(result.removeprefix('ERROR\t'))
                status.set('Сохранено: ' + result.split('\t')[1] + '. Сочетание уже работает в CorelDRAW.')
                save.configure(state='disabled')
                notice.set('Можно закрыть это окно. Документ и настройки других команд не изменялись,\nкроме подтверждённого переназначения выбранного сочетания.')
            except Exception as exc:
                status.set('Не удалось подтвердить сохранение: ' + str(exc))
                save.configure(state='disabled')
            return
        # The conflict is shown before this explicit replace/save action.
        state['waiting'] = True
        field.configure(state='disabled')
        save.configure(state='disabled')
        status.set('Ожидаю закрытия CorelDRAW. Сохраните документы и закройте Corel.\nНе запускайте его снова до сообщения «Сохранено». Закрытие этого окна отменит ожидание.')
        poll()

    save = ttk.Button(row, text='Сохранить', command=save_key, state='disabled')
    save.pack(side='left')
    other = ttk.Button(row, text='Выбрать другое', command=choose_other)
    other.pack(side='left', padx=8)
    ttk.Button(row, text='Закрыть', command=root.destroy).pack(side='right')
    field.bind('<KeyPress>', capture)
    field.bind('<<Paste>>', lambda e: 'break')
    field.bind('<<Cut>>', lambda e: 'break')
    field.focus_set()
    root.mainloop()


if __name__ == '__main__':
    try:
        run()
    except Exception as exc:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror('Горячая клавиша DXF', str(exc), parent=root)
        root.destroy()
