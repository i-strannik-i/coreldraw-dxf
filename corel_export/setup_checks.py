"""Bounded Corel capability checks, isolated from the installer UI."""
import json
import os
from pathlib import Path
import struct
import subprocess
import sys
import time
import winreg


def discover():
    found = []
    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, '') as root:
        for index in range(winreg.QueryInfoKey(root)[0]):
            name = winreg.EnumKey(root, index)
            if not name.startswith('CorelDRAW.Application.'):
                continue
            suffix = name.rsplit('.', 1)[-1]
            if suffix.isdigit():
                found.append(int(suffix))
    return sorted(set(found))


def validate_binary(path):
    with open(path, 'rb') as stream:
        if stream.read(2) != b'MZ':
            raise RuntimeError('Повреждён исполняемый файл CorelDRAW.')
        stream.seek(0x3c)
        offset = struct.unpack('<I', stream.read(4))[0]
        stream.seek(offset)
        if stream.read(4) != b'PE\0\0' or stream.read(2) != b'\x64\x86':
            raise RuntimeError('Нужен CorelDRAW 2024 x64. Эта сборка не поддерживает 32-битный Corel.')


def repair_path(executable):
    path = Path(executable).parent.parent / 'Setup' / 'SetupARP.exe'
    return path if path.is_file() else None


def vba_packages(executable):
    folder = Path(executable).parent.parent / 'Setup' / 'MSIs'
    packages = [folder / 'VBA/Vba71.msi', folder / 'VBA/Vba71_1033.MSI', folder / 'VBA_x64.msi']
    if not all(path.is_file() for path in packages):
        raise RuntimeError('В локальном комплекте Corel нет полного набора VBA. Нужен официальный установочный комплект CorelDRAW 2024.')
    return packages


def verify_package(path):
    # Never elevate an unverified cached installer (including modified MSI files).
    script = "$s=Get-AuthenticodeSignature -LiteralPath $env:SKLAD_MSI; [pscustomobject]@{status=[string]$s.Status; signer=[string]$s.SignerCertificate.Subject} | ConvertTo-Json -Compress"
    env = dict(os.environ, SKLAD_MSI=str(path))
    result = subprocess.run(['powershell.exe', '-NoProfile', '-NonInteractive', '-Command', script],
                            env=env, capture_output=True, text=True, timeout=60, check=True,
                            creationflags=subprocess.CREATE_NO_WINDOW)
    data = json.loads(result.stdout)
    if data['status'] != 'Valid' or not any(name in data['signer'].lower() for name in ('microsoft corporation', 'corel corporation')):
        raise RuntimeError('Не подтверждена подлинность компонента ' + path.name + '. Автоматическая установка остановлена. Нужен оригинальный комплект Corel.')


def install_vba(executable, log_folder, status):
    import win32api
    import win32event
    import win32process
    from win32com.shell import shell, shellcon
    packages = vba_packages(executable)
    for path in packages:
        verify_package(path)
    for path in packages:
        status('Добавление VBA: ' + path.name + '. Подтвердите запрос Windows, если он появится.')
        log = Path(log_folder) / (path.stem + '.log')
        args = subprocess.list2cmdline(['/i', str(path), '/passive', '/norestart', '/L*v', str(log)])
        info = shell.ShellExecuteEx(fMask=shellcon.SEE_MASK_NOCLOSEPROCESS, lpVerb='runas',
                                   lpFile=str(Path(os.environ['WINDIR']) / 'System32/msiexec.exe'),
                                   lpParameters=args, nShow=1)
        handle = info['hProcess']
        try:
            win32event.WaitForSingleObject(handle, win32event.INFINITE)
            code = win32process.GetExitCodeProcess(handle)
        finally:
            win32api.CloseHandle(handle)
        if code == 3010:
            raise RuntimeError('Компонент VBA установлен, Windows требует перезагрузку. После неё повторите установку DXF.')
        if code != 0:
            raise RuntimeError('Установка VBA остановлена, код ' + str(code) + '. Журнал: ' + str(log))


def probe(output, installed=False):
    import pythoncom
    import win32com.client
    try:
        from .setup_corel import corel_running
    except ImportError:
        from setup_corel import corel_running
    result = {'ok': False, 'kind': 'corel', 'message': ''}
    app = None
    pythoncom.CoInitialize()
    try:
        if corel_running():
            raise RuntimeError('Сохраните документы и закройте CorelDRAW. Проверка не закрывает ваши окна.')
        app = win32com.client.DispatchEx('CorelDRAW.Application.25')
        major = int(app.VersionMajor)
        if major != 25:
            raise RuntimeError('COM запустил другую версию CorelDRAW: ' + str(major))
        result['version'] = str(app.Version)
        result['kind'] = 'vba'
        if not app.InitializeVBA():
            raise RuntimeError('CorelDRAW не смог инициализировать VBA.')
        count = int(app.GMSManager.Projects.Count)
        result['projects'] = count
        # Access the runtime, not the VB project editor (which may be policy-blocked).
        if installed:
            result['kind'] = 'macro'
            diagnostic = str(app.GMSManager.RunMacro('SkladCorelDXF', 'ExportDXF.InstallationCheck'))
            if diagnostic != 'SKLAD_DXF_OK':
                raise RuntimeError('Макрос запустился, но не нашёл конвертер: ' + diagnostic)
        result.update(ok=True, kind='ok', message='CorelDRAW 2024 x64 и VBA работают.')
    except Exception as error:
        result['message'] = str(error)
    finally:
        if app is not None:
            try:
                if int(app.Documents.Count) == 0:
                    app.Quit()
                else:
                    result.update(ok=False, kind='open', message='Corel открыл документ. Сохраните его и закройте Corel, затем повторите проверку.')
            except Exception:
                result.update(ok=False, kind='open', message='Закройте проверочный экземпляр CorelDRAW и повторите установку.')
        pythoncom.CoUninitialize()
        Path(output).write_text(json.dumps(result, ensure_ascii=False), encoding='utf-8')


def run_probe(folder, installed=False):
    output = Path(folder) / ('verify.json' if installed else 'preflight.json')
    args = [sys.executable]
    if not getattr(sys, 'frozen', False):
        args.append(str(Path(__file__).with_name('setup_corel.py')))
    args.extend(['--probe', str(output)])
    if installed:
        args.append('--installed')
    try:
        subprocess.run(args, timeout=120, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
    except subprocess.TimeoutExpired:
        return {'ok': False, 'kind': 'timeout', 'message': 'Corel не ответил за 2 минуты. Проверьте окно лицензии или запуска Corel. Документы и процесс Corel принудительно не закрывались.'}
    result = json.loads(output.read_text(encoding='utf-8'))
    # Corel may finish writing its workspace shortly after Quit returns.
    try:
        from .setup_corel import corel_running
    except ImportError:
        from setup_corel import corel_running
    if result['ok']:
        for _ in range(50):
            if not corel_running():
                return result
            time.sleep(.2)
        return {'ok': False, 'kind': 'open', 'message': 'Corel ещё завершает работу. Дождитесь закрытия и повторите проверку.'}
    return result
