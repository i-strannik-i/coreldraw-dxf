Attribute VB_Name = "ExportDXF"
Option Explicit
Private Declare PtrSafe Function OpenMutex Lib "kernel32" Alias "OpenMutexA" (ByVal access As Long, ByVal inherit As Long, ByVal name As String) As LongPtr
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal handle As LongPtr) As Long

Public Function InstallationCheck() As String
    On Error GoTo Failed
    Dim fs As Object, folder As String
    Set fs = CreateObject("Scripting.FileSystemObject")
    folder = fs.BuildPath(Application.GMSManager.UserGMSPath, "SkladCorelDXF_runtime")
    If Not fs.FileExists(fs.BuildPath(folder, "CorelDXF.exe")) Then Err.Raise 513, , "Converter missing"
    If Not fs.FileExists(fs.BuildPath(folder, "HotkeySettings.exe")) Then Err.Raise 513, , "Hotkey helper missing"
    If Not fs.FileExists(fs.BuildPath(folder, "VectorTools.exe")) Then Err.Raise 513, , "Vector tools missing"
    If Not fs.FileExists(fs.BuildPath(folder, "UpdateCorelDXF.exe")) Then Err.Raise 513, , "Updater missing"
    InstallationCheck = "SKLAD_DXF_OK"
    Exit Function
Failed:
    InstallationCheck = "ERROR: " & Err.Description
End Function

Public Function Diagnostics() As String
    On Error Resume Next
    Dim fs As Object, sh As Object, path As String
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set sh = CreateObject("WScript.Shell")
    path = Application.GMSManager.UserGMSPath & "SkladCorelDXF_runtime\CorelDXF.exe"
    Diagnostics = path & " | exists=" & CStr(fs.FileExists(path)) & " | selected=" & CStr(ActiveSelectionRange.Count) & " | dir=" & Dir$(path) & " | folder=" & CStr(fs.FolderExists(fs.GetParentFolderName(path))) & " | error=" & Err.Description
End Function

Private Function ExportBusy() As Boolean
    Dim handle As LongPtr
    handle = OpenMutex(&H100000, 0, "Local\SkladCorelDXFExport")
    If handle <> 0 Then
        CloseHandle handle
        ExportBusy = True
        MsgBox "DXF export is still running. Please wait.", vbInformation, "CorelDRAW -> DXF"
    End If
End Function

Public Sub ShowExporter()
    If Application.Documents.Count = 0 Then Exit Sub
    If ExportBusy() Then Exit Sub
    LaunchVectorTools "--export"
End Sub

Public Sub ShowLayerHelp()
    LaunchVectorTools "--help"
End Sub

Public Sub ShowHotkeys()
    AutoCheckUpdates
    On Error GoTo Failed
    Dim engine As String
    engine = Application.GMSManager.UserGMSPath & "SkladCorelDXF_runtime\HotkeySettings.exe"
    If Dir$(engine) = "" Then Err.Raise 513, , "HotkeySettings.exe is missing. Update Sklad DXF."
    CreateObject("WScript.Shell").Run Chr$(34) & engine & Chr$(34), 1, False
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "DXF"
End Sub

Public Sub AutoCheckUpdates()
    On Error Resume Next
    Dim engine As String
    engine = Application.GMSManager.UserGMSPath & "SkladCorelDXF_runtime\UpdateCorelDXF.exe"
    If Dir$(engine) <> "" Then
        CreateObject("WScript.Shell").Run Chr$(34) & engine & Chr$(34) & " --auto", 0, False
    End If
End Sub

Public Sub ShowUpdates()
    Dim engine As String
    engine = Application.GMSManager.UserGMSPath & "SkladCorelDXF_runtime\UpdateCorelDXF.exe"
    If Dir$(engine) = "" Then
        MsgBox "Компонент обновления не установлен. Переустановите DXF.", vbExclamation, "DXF"
        Exit Sub
    End If
    CreateObject("WScript.Shell").Run Chr$(34) & engine & Chr$(34), 1, False
End Sub

Private Sub LaunchVectorTools(ByVal mode As String)
    AutoCheckUpdates
    Dim engine As String
    engine = Application.GMSManager.UserGMSPath & "SkladCorelDXF_runtime\VectorTools.exe"
    If Dir$(engine) = "" Then
        MsgBox "Компонент проверки отсутствует. Обновите DXF.", vbExclamation, "DXF"
        Exit Sub
    End If
    CreateObject("WScript.Shell").Run Chr$(34) & engine & Chr$(34) & " " & mode, 1, False
End Sub

Public Sub ShowValidator()
    LaunchVectorTools "--check"
End Sub

Public Sub ShowJoiner()
    LaunchVectorTools "--join"
End Sub

Public Sub InstallToolbar()
    Dim bar As CommandBar, button As Control, item As Control
    Dim exportButton As Control, helpButton As Control, icons As String
    On Error Resume Next
    Set bar = Application.CommandBars("DXF Export")
    On Error GoTo 0
    If bar Is Nothing Then Set bar = Application.CommandBars.Add("DXF Export", cuiBarTop, False)
    For Each item In bar.Controls
        If item.Tag = "SkladDXF.Export" Or item.Caption = "ExportDXF.ShowExporter" Or item.Caption = "DXF" Or item.Caption = "Экспорт DXF" Then Set exportButton = item
        If item.Tag = "SkladDXF.Help" Or item.Caption = "ExportDXF.ShowLayerHelp" Or item.Caption = "?" Or item.Caption = "? Слои" Or item.Caption = "Слои ?" Then Set helpButton = item
    Next item
    icons = Application.GMSManager.UserGMSPath & "SkladCorelDXF_runtime\icons\"
    If exportButton Is Nothing Then Set exportButton = bar.Controls.AddCustomButton(cdrCmdCategoryMacros, "SkladCorelDXF.ExportDXF.ShowExporter", 0, False)
    exportButton.Tag = "SkladDXF.Export"
    exportButton.Caption = "Экспорт DXF"
    exportButton.ToolTipText = "Export current page or selection to DXF"
    exportButton.SetIcon2 "guid://87260fe1-e755-4c95-9f68-2aa2f3dc5001"
    If helpButton Is Nothing Then Set helpButton = bar.Controls.AddCustomButton(cdrCmdCategoryMacros, "SkladCorelDXF.ExportDXF.ShowLayerHelp", 0, False)
    helpButton.Tag = "SkladDXF.Help"
    helpButton.Caption = "Слои ?"
    helpButton.ToolTipText = "Layer naming guide"
    helpButton.SetIcon2 "guid://87260fe1-e755-4c95-9f68-2aa2f3dc5002"
    bar.Visible = True
End Sub

Public Sub RunExport(ByVal SelectionOnly As Boolean)
    On Error GoTo Failed
    If ExportBusy() Then Exit Sub
    Dim request As String, engine As String
    request = PrepareExport(SelectionOnly)
    engine = Application.GMSManager.UserGMSPath & "SkladCorelDXF_runtime\CorelDXF.exe"
    CreateObject("WScript.Shell").Run Chr$(34) & engine & Chr$(34) & " " & Chr$(34) & request & Chr$(34), 1, False
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "CorelDRAW -> DXF"
End Sub

Public Function PrepareExport(ByVal SelectionOnly As Boolean) As String
    On Error GoTo Failed
    Dim doc As Document, opt As StructSaveAsOptions
    Dim fs As Object, stream As Object, shell As Object
    Dim root As String, job As String, target As String, engine As String
    Dim pageNumber As Long, selection As ShapeRange, pageShapes As ShapeRange
    Dim failure As String, master As Layer
    Set doc = ActiveDocument
    If SelectionOnly And ActiveSelectionRange.Count = 0 Then Err.Raise 513, , "No selected objects."
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set shell = CreateObject("WScript.Shell")
    root = shell.ExpandEnvironmentStrings("%TEMP%") & "\SkladCorelDXF"
    engine = Application.GMSManager.UserGMSPath & "SkladCorelDXF_runtime\CorelDXF.exe"
    If fs.FileExists(engine) = False Then Err.Raise 513, , "Missing runtime: " & engine
    If fs.FolderExists(root) = False Then fs.CreateFolder root
    If fs.FolderExists(root & "\jobs") = False Then fs.CreateFolder root & "\jobs"
    job = root & "\jobs\" & fs.GetTempName
    fs.CreateFolder job
    If Len(doc.FilePath) > 0 Then target = fs.BuildPath(doc.FilePath, fs.GetBaseName(doc.FileName) & ".dxf")
    Set opt = Application.CreateStructSaveAsOptions
    pageNumber = doc.ActivePage.Index
    Set selection = ActiveSelectionRange
    opt.Filter = cdrCDR
    opt.EmbedVBAProject = False
    opt.Range = cdrSelection
    If SelectionOnly = False Then
        For Each master In doc.MasterPage.Layers
            If master.Shapes.Count > 0 Then Err.Raise 513, , "Master-page objects are not supported. Export an explicit selection instead."
        Next master
        Set pageShapes = doc.ActivePage.Shapes.All
        If pageShapes.Count = 0 Then Err.Raise 513, , "Current page has no objects."
        pageShapes.CreateSelection
        If ActiveSelectionRange.Count <> pageShapes.Count Then Err.Raise 513, , "Cannot select all objects on this page."
    End If
    doc.SaveAsCopy job & "\snapshot.cdr", opt
    doc.Pages(pageNumber).Activate
    selection.CreateSelection
    Set stream = fs.CreateTextFile(job & "\request.txt", True, True)
    stream.WriteLine job & "\snapshot.cdr"
    stream.WriteLine target
    If SelectionOnly Then
        stream.WriteLine "selection"
        stream.WriteLine "1"
    Else
        stream.WriteLine "page"
        stream.WriteLine "1"
    End If
    stream.Close
    ' The caller owns conversion, progress, cancellation and snapshot cleanup.
    PrepareExport = job & "\request.txt"
    Exit Function
Failed:
    failure = Err.Description
    On Error Resume Next
    If Not selection Is Nothing Then
        doc.Pages(pageNumber).Activate
        selection.CreateSelection
    End If
    If Len(job) > 0 Then fs.DeleteFile job & "\snapshot.cdr", True
    On Error GoTo 0
    Err.Raise 513, "PrepareExport", failure
End Function
