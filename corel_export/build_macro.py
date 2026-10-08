"""Build a GMS through Corel's VBA API; never change macro security settings."""
from pathlib import Path
import sys
import win32com.client


def build(destination):
    app = win32com.client.gencache.EnsureDispatch('CorelDRAW.Application.25')
    if not app.InitializeVBA():
        raise RuntimeError('CorelDRAW VBA is unavailable')
    destination = Path(destination).resolve()
    owned = [p for p in app.GMSManager.Projects if p.Name == 'SkladCorelDXF']
    if owned:
        if Path(owned[0].FullFileName).resolve() != destination:
            raise RuntimeError('Another copy of the macro is loaded; unload it before building')
        # Unloading removes live toolbar buttons and shortcut references.
        # Update our already-loaded project in place instead.
    else:
        if destination.exists():
            app.GMSManager.Projects.Load(str(destination))
        else:
            app.GMSManager.Projects.Create('SkladCorelDXF', False, str(destination))
    project = next(p for p in app.VBE.VBProjects if p.Name == 'SkladCorelDXF')
    existing = {item.Name: item for item in project.VBComponents}
    module = existing.get('ExportDXF') or project.VBComponents.Add(1)
    module.Name = 'ExportDXF'
    if module.CodeModule.CountOfLines:
        module.CodeModule.DeleteLines(1, module.CodeModule.CountOfLines)
    source = (Path(__file__).parent / 'ExportDXF.bas').read_text(encoding='utf-8')
    module.CodeModule.AddFromString('\n'.join(source.splitlines()[1:]))
    bridge = existing.get('HotkeyBridge') or project.VBComponents.Add(1)
    bridge.Name = 'HotkeyBridge'
    if bridge.CodeModule.CountOfLines:
        bridge.CodeModule.DeleteLines(1, bridge.CodeModule.CountOfLines)
    source = (Path(__file__).parent / 'HotkeyBridge.bas').read_text(encoding='utf-8')
    bridge.CodeModule.AddFromString('\n'.join(source.splitlines()[1:]))
    validator = existing.get('ContourCheck') or project.VBComponents.Add(1)
    validator.Name = 'ContourCheck'
    if validator.CodeModule.CountOfLines:
        validator.CodeModule.DeleteLines(1, validator.CodeModule.CountOfLines)
    source = (Path(__file__).parent / 'ContourCheck.bas').read_text(encoding='utf-8')
    validator.CodeModule.AddFromString('\n'.join(source.splitlines()[1:]))
    form = existing.get('DXFOptions') or project.VBComponents.Add(3)
    form.Name = 'DXFOptions'
    ui = form.Designer
    if ui is None:
        form.DesignerWindow()
        ui = form.Designer
    for control in list(ui.Controls):
        ui.Controls.Remove(control.Name)
    if form.CodeModule.CountOfLines:
        form.CodeModule.DeleteLines(1, form.CodeModule.CountOfLines)
    for key, value in {'Caption': 'CorelDRAW → DXF', 'Width': 355, 'Height': 264}.items():
        form.Properties.Item(key).Value = value
    controls = [
        ('Forms.Label.1', 'title', 'Что экспортировать?', 14, 12, 240, 20),
        ('Forms.CommandButton.1', 'SettingsButton', '', 271, 8, 26, 26),
        ('Forms.CommandButton.1', 'HelpButton', '?', 303, 8, 26, 26),
        ('Forms.OptionButton.1', 'SelectionOnly', 'Выделенные объекты', 14, 40, 315, 22),
        ('Forms.OptionButton.1', 'WholePage', 'Весь текущий лист', 14, 67, 315, 22),
        ('Forms.Label.1', 'hint', 'Слои и цвета • Масштаб 1:1 • Допуск 0,1 мм', 14, 102, 320, 18),
        ('Forms.Label.1', 'path', 'DXF рядом с CDR. Существующий DXF получит резервную копию.', 14, 126, 320, 30),
        ('Forms.Label.1', 'CheckHint', 'Соединение и замыкание — в окне «Проверить».', 14, 159, 174, 33),
        ('Forms.CommandButton.1', 'UpdateButton', 'Обновления', 199, 162, 130, 26),
        ('Forms.CommandButton.1', 'ExportButton', 'Экспортировать', 105, 200, 130, 26),
        ('Forms.CommandButton.1', 'CancelButton', 'Отмена', 245, 200, 84, 26),
    ]
    for progid, name, caption, left, top, width, height in controls:
        control = ui.Controls.Add(progid, name, True)
        for key, value in dict(Caption=caption, Left=left, Top=top, Width=width, Height=height).items():
            setattr(control, key, value)
        control.Font.Name = 'Segoe UI'
        control.Font.Size = 10
    for index, name in enumerate(('SelectionOnly', 'WholePage', 'ExportButton', 'CancelButton', 'SettingsButton', 'HelpButton')):
        ui.Controls.Item(name).TabIndex = index
    form.CodeModule.AddFromString(r'''Option Explicit
Private Sub UserForm_Initialize()
    On Error Resume Next
    Set SettingsButton.Picture = LoadPicture(Application.GMSManager.UserGMSPath & "SkladCorelDXF_runtime\icons\Settings_20.bmp")
    On Error GoTo 0
    SettingsButton.ControlTipText = "Настройка горячей клавиши"
    SettingsButton.TakeFocusOnClick = False
    SelectionOnly.Enabled = (ActiveSelectionRange.Count > 0)
    SelectionOnly.Value = SelectionOnly.Enabled
    WholePage.Value = Not SelectionOnly.Enabled
End Sub
Private Sub ExportButton_Click()
    Dim selected As Boolean
    selected = SelectionOnly.Value
    Unload Me
    ExportDXF.RunExport selected
End Sub
Private Sub CancelButton_Click()
    Unload Me
End Sub
Private Sub HelpButton_Click()
    ExportDXF.ShowLayerHelp
End Sub
Private Sub SettingsButton_Click()
    Unload Me
    ExportDXF.ShowHotkeys
End Sub
Private Sub UpdateButton_Click()
    Unload Me
    ExportDXF.ShowUpdates
End Sub
''')
    if 'HotkeySettings' in existing:
        project.VBComponents.Remove(existing['HotkeySettings'])
    help_form = existing.get('LayerHelp') or project.VBComponents.Add(3)
    help_form.Name = 'LayerHelp'
    if help_form.Designer is None:
        help_form.DesignerWindow()
    for control in list(help_form.Designer.Controls):
        help_form.Designer.Controls.Remove(control.Name)
    if help_form.CodeModule.CountOfLines:
        help_form.CodeModule.DeleteLines(1, help_form.CodeModule.CountOfLines)
    for key, value in {'Caption': 'Памятка дизайнеру: слои', 'Width': 470, 'Height': 427}.items():
        help_form.Properties.Item(key).Value = value
    rows = [
        ('INFO', 'Информационный текст и пояснения. Без резки и гравировки.'),
        ('#', 'Артикулы: гравировка на глубину 0,1 мм.'),
        ('CUT_OUT', 'Сквозной рез снаружи контура.'),
        ('CUT_IN', 'Сквозной рез внутри контура.'),
        ('CUT_ON', 'Рез по линии контура.'),
        ('D / D4', 'Отверстия. D4: диаметр отверстия 4 мм.'),
        ('B', 'Круглая фреза: паз под сгиб материала.'),
        ('P / P_2.5', 'Выборка. P_2.5: глубина 2,5 мм.'),
        ('V / V_3', 'Гравировка. V_3: ширина 3 мм, не глубина.'),
        ('BACK…', 'Префикс для обратной стороны, например BACK_CUT_OUT.'),
    ]
    for index, (name, explanation) in enumerate(rows):
        for suffix, caption, left, width in [('name', name, 14, 88), ('text', explanation, 106, 330)]:
            label = help_form.Designer.Controls.Add('Forms.Label.1', f'row{index}_{suffix}', True)
            label.Caption = caption
            label.Left, label.Top, label.Width, label.Height = left, 14 + index * 31, width, 29
            label.Font.Name = 'Segoe UI'
            label.Font.Size = 10
            label.Font.Bold = suffix == 'name'
    note = help_form.Designer.Controls.Add('Forms.Label.1', 'Note', True)
    note.Caption = 'Названия — латиницей (B, P, V). Размеры — в мм.\nПамятка не создаёт траектории и не назначает инструмент автоматически.'
    note.Left, note.Top, note.Width, note.Height = 14, 326, 427, 34
    note.Font.Name, note.Font.Size = 'Segoe UI', 9
    close = help_form.Designer.Controls.Add('Forms.CommandButton.1', 'CloseButton', True)
    close.Caption = 'Закрыть'
    close.Left, close.Top, close.Width, close.Height = 341, 363, 100, 23
    close.Cancel = True
    help_form.CodeModule.AddFromString('Private Sub CloseButton_Click()\nUnload Me\nEnd Sub')
    # Save only our new project; do not alter any other loaded macro.
    app.VBE.ActiveVBProject = project
    app.VBE.CommandBars.Item('File').Controls.Item(3).Execute()
    if not destination.is_file():
        raise RuntimeError('Corel did not save the GMS project')
    print(destination)


if __name__ == '__main__':
    build(sys.argv[1])
