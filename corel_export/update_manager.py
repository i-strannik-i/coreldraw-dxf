"""Background release discovery and explicitly confirmed installation."""
import hashlib
import json
import os
from pathlib import Path
import queue
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import urllib.request

try:
    from .version import VERSION, CHANNEL
except ImportError:
    from version import VERSION, CHANNEL

MANIFEST = 'CorelDXF-latest.json'
GITHUB_REPO = 'i-strannik-i/coreldraw-dxf'


def read_url(url, limit):
    if not url.startswith(('https://api.github.com/repos/' + GITHUB_REPO + '/',
                           'https://github.com/' + GITHUB_REPO + '/releases/latest/download/',
                           'https://github.com/' + GITHUB_REPO + '/releases/download/')):
        raise ValueError('Недопустимый адрес обновления')
    request = urllib.request.Request(url, headers={'User-Agent': 'CorelDXF-Updater', 'Accept': 'application/vnd.github+json'})
    with urllib.request.urlopen(request, timeout=20) as response:
        data = response.read(limit + 1)
    if len(data) > limit:
        raise ValueError('Превышен размер файла обновления')
    return data


def github_release():
    if CHANNEL == 'beta':
        try:
            releases=json.loads(read_url('https://api.github.com/repos/'+GITHUB_REPO+'/releases?per_page=20',262144))
            tags=[item['tag_name'][1:] for item in releases if not item.get('draft') and re.fullmatch(r'v\d+\.\d+\.\d+',item.get('tag_name',''))]
            if tags:
                version=max(tags,key=version_tuple)
                base='https://github.com/'+GITHUB_REPO+'/releases/download/v'+version+'/'
                info=json.loads(read_url(base+MANIFEST,16384))
                if info.get('version')!=version:raise ValueError('Версия бета-выпуска не совпала.')
                validate_info(info,VERSION)
                info['_url']=base+info['file']
                return info
        except (OSError, ValueError, KeyError, TypeError):
            pass
    # Public latest-release assets avoid shared-IP GitHub API rate limits.
    base = 'https://github.com/' + GITHUB_REPO + '/releases/'
    info = json.loads(read_url(base + 'latest/download/' + MANIFEST, 16384))
    validate_info(info, VERSION)
    tagged = base + 'download/v' + info['version'] + '/'
    pinned = json.loads(read_url(tagged + MANIFEST, 16384))
    validate_info(pinned, VERSION)
    if pinned != info:
        raise ValueError('Выпуск изменился. Повторите проверку обновлений.')
    info['_url'] = tagged + info['file']
    return info


def check_latest(folder):
    releases, failures = [], []
    for check in (github_release, lambda: check_release(folder)):
        try:
            releases.append(check())
        except Exception as error:
            failures.append(str(error))
    if not releases:
        raise ValueError('GitHub и резервная папка недоступны: ' + '; '.join(failures))
    return max(releases, key=lambda info: version_tuple(info['version']))


def automatic_release(folder):
    """A failed background check must not interrupt the designer's work."""
    try:
        release = check_latest(folder)
        return release if release['available'] else None
    except Exception:
        return None


def stage_update(folder, info, cache):
    if '_url' not in info:
        return stage_release(folder, info, cache)
    if github_release() != info or not info['available']:
        raise ValueError('Версия обновления изменилась. Повторите проверку.')
    data = read_url(info['_url'], info['size'])
    if len(data) != info['size'] or hashlib.sha256(data).hexdigest() != info['sha256']:
        raise ValueError('Проверка установщика не пройдена. Запуск запрещён.')
    cache = Path(cache)
    cache.mkdir(parents=True, exist_ok=True)
    descriptor, temporary = tempfile.mkstemp(suffix='.download', dir=cache)
    try:
        with os.fdopen(descriptor, 'wb') as stream:
            stream.write(data)
        target = cache / info['file']
        os.replace(temporary, target)
        return target
    finally:
        Path(temporary).unlink(missing_ok=True)


def version_tuple(value):
    if not isinstance(value, str) or not re.fullmatch(r'\d+\.\d+\.\d+', value):
        raise ValueError('Некорректный номер версии.')
    return tuple(map(int, value.split('.')))


def settings_file():
    return Path(os.environ['LOCALAPPDATA']) / 'SkladCorelDXF/update-settings.json'


def update_folder():
    try:
        return Path(json.loads(settings_file().read_text(encoding='utf-8'))['folder'])
    except (OSError, ValueError, KeyError):
        folder=Path.home() / 'Yandex.Disk/Производство/чертежи фрезер/_плагины'
        return folder/'beta' if CHANNEL=='beta' else folder


def check_release(folder, current=VERSION):
    folder = Path(folder)
    info = json.loads((folder / MANIFEST).read_text(encoding='utf-8'))
    return validate_info(info, current)


def validate_info(info, current):
    if info.get('product') != 'CorelDXF' or info.get('corel_major') != 25:
        raise ValueError('Обновление предназначено для другого продукта или Corel.')
    version = info.get('version')
    version_tuple(version)
    if info.get('file') != 'SetupCorelDXF-' + version + '.exe':
        raise ValueError('Некорректное имя установщика.')
    if not re.fullmatch('[0-9a-f]{64}', str(info.get('sha256', ''))):
        raise ValueError('Некорректная контрольная сумма.')
    if type(info.get('size')) is not int or not 0 < info['size'] < 1024**3:
        raise ValueError('Некорректный размер установщика.')
    info['available'] = version_tuple(version) > version_tuple(current)
    return info


def stage_release(folder, info, cache):
    # Recheck after user confirmation: the shared manifest may have changed.
    fresh = check_release(folder)
    if fresh != info or not fresh['available']:
        raise ValueError('Выпуск изменился. Повторите проверку обновлений.')
    source = Path(folder) / info['file']
    if source.stat().st_size != info['size']:
        raise ValueError('Установщик ещё не синхронизирован или повреждён. Повторите позже.')
    cache = Path(cache)
    cache.mkdir(parents=True, exist_ok=True)
    descriptor, temporary = tempfile.mkstemp(suffix='.download', dir=cache)
    os.close(descriptor)
    try:
        shutil.copyfile(source, temporary)
        digest = hashlib.sha256(Path(temporary).read_bytes()).hexdigest()
        if digest != info['sha256'] or Path(temporary).stat().st_size != info['size']:
            raise ValueError('Контрольная сумма не совпала. Установка запрещена; дождитесь синхронизации.')
        target = cache / info['file']
        os.replace(temporary, target)
        return target
    finally:
        Path(temporary).unlink(missing_ok=True)


def main():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    root = tk.Tk()
    automatic = '--auto' in sys.argv
    if automatic:
        root.withdraw()
    root.title('Обновления CorelDRAW → DXF')
    root.geometry('610x330')
    root.resizable(False, False)
    frame = ttk.Frame(root, padding=22)
    frame.pack(fill='both', expand=True)
    ttk.Label(frame, text='CorelDRAW → DXF  ' + VERSION, font=('Segoe UI', 17, 'bold')).pack(anchor='w')
    folder = tk.StringVar(value=str(update_folder()))
    ttk.Label(frame, textvariable=folder, wraplength=555).pack(anchor='w', pady=12)
    status = tk.StringVar(value='Проверка GitHub и резервной папки…')
    ttk.Label(frame, textvariable=status, wraplength=555).pack(anchor='w', pady=10)
    ttk.Label(frame, text='Обновление запускается только с вашего согласия.\nПеред установкой сохраните документы и закройте Corel.', wraplength=555).pack(anchor='w')
    row = ttk.Frame(frame)
    row.pack(fill='x', pady=16)
    events = queue.Queue()
    state = {'release': None, 'busy': False}

    def work(install=False):
        if state['busy']:
            return
        source = folder.get()
        release = state['release']
        if install and not messagebox.askyesno('Обновление DXF', 'Запустить установщик версии ' + release['version'] + '?\nСохраните документы и закройте Corel перед установкой.', parent=root):
            return
        state['busy'] = True
        apply.configure(state='disabled')
        status.set('Копирование и проверка установщика…' if install else 'Проверка GitHub и резервной папки…')
        def worker():
            try:
                if install:
                    path = stage_update(source, release, Path(os.environ['LOCALAPPDATA']) / 'SkladCorelDXF/updates')
                    subprocess.Popen([str(path)])
                    events.put(('launched', None))
                else:
                    events.put(('release', automatic_release(source) if automatic else check_latest(source)))
            except Exception as error:
                events.put(('error', str(error)))
        threading.Thread(target=worker, daemon=True).start()

    def choose():
        if state['busy']:
            return
        selected = filedialog.askdirectory(parent=root, title='Общая папка _плагины')
        if selected:
            settings_file().parent.mkdir(parents=True, exist_ok=True)
            settings_file().write_text(json.dumps({'folder': selected}, ensure_ascii=False), encoding='utf-8')
            folder.set(selected)
            work()

    def poll():
        while not events.empty():
            kind, value = events.get_nowait()
            state['busy'] = False
            state['release'] = value if kind == 'release' else None
            if kind == 'release':
                if automatic and value is None:
                    root.destroy()
                    return
                if automatic:
                    root.deiconify()
                    root.lift()
                status.set('Доступна версия ' + value['version'] if value['available'] else 'Установлена актуальная версия. Обновление не требуется.')
                apply.configure(state='normal' if value['available'] else 'disabled')
            elif kind == 'launched':
                # Release our executable before the installer replaces the runtime.
                root.destroy()
                return
            else:
                status.set('Не удалось проверить или запустить обновление: ' + value)
        root.after(150, poll)

    ttk.Button(row, text='Проверить', command=work).pack(side='left')
    ttk.Button(row, text='Выбрать папку', command=choose).pack(side='left', padx=6)
    ttk.Button(row, text='Позже', command=lambda: root.destroy() if not state['busy'] else None).pack(side='left')
    apply = ttk.Button(row, text='Обновить', command=lambda: work(True), state='disabled')
    apply.pack(side='right')
    root.protocol('WM_DELETE_WINDOW', lambda: None if state['busy'] else root.destroy())
    work()
    poll()
    root.mainloop()


if __name__ == '__main__':
    import win32api
    import win32event
    import winerror
    mutex = win32event.CreateMutex(None, False, 'Local\\SkladCorelDXFUpdater')
    try:
        if win32api.GetLastError() != winerror.ERROR_ALREADY_EXISTS:
            main()
    finally:
        win32api.CloseHandle(mutex)
