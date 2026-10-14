Option Explicit
On Error Resume Next
Dim fs, sh, here, root, app, gms, p, silent
Set fs = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
here = fs.GetParentFolderName(WScript.ScriptFullName)
silent = (WScript.Arguments.Count > 0)
If Not fs.FileExists(here & "\runtime\CorelDXF.exe") Then Fail "Extract the entire ZIP before installing."
Set app = CreateObject("CorelDRAW.Application.25")
If Err.Number <> 0 Then Fail "CorelDRAW 2024 is required. " & Err.Description
If Not app.InitializeVBA Then Fail "Install the VBA component of CorelDRAW first."
If Err.Number <> 0 Then Fail Err.Description
gms = app.GMSManager.UserGMSPath
root = gms & "SkladCorelDXF_runtime"
If Not fs.FolderExists(gms) Then fs.CreateFolder gms
For Each p In app.GMSManager.Projects
    If p.Name = "SkladCorelDXF" Then p.Unload
Next
If Err.Number <> 0 Then Fail "Close the export window before installing. " & Err.Description
If Not fs.FolderExists(root) Then fs.CreateFolder root
fs.CopyFolder here & "\runtime\*", root & "\", True
fs.CopyFile here & "\runtime\CorelDXF.exe", root & "\CorelDXF.exe", True
fs.CopyFolder here & "\icons", root & "\icons", True
If Err.Number <> 0 Then Fail "Runtime installation failed: " & Err.Description
If Not fs.FolderExists(gms) Then fs.CreateFolder gms
If fs.FileExists(gms & "SkladCorelDXF.gms") Then fs.CopyFile gms & "SkladCorelDXF.gms", gms & "SkladCorelDXF.gms.bak", True
fs.CopyFile here & "\SkladCorelDXF.gms", gms & "SkladCorelDXF.gms", True
If Err.Number <> 0 Then Fail "Macro installation failed: " & Err.Description
app.GMSManager.Projects.Load gms & "SkladCorelDXF.gms"
If Err.Number <> 0 Then Fail "Could not load macro: " & Err.Description
If Not silent Then MsgBox "Installed. Run macro SkladCorelDXF.ExportDXF.ShowExporter in CorelDRAW.", vbInformation, "CorelDRAW -> DXF"
WScript.Quit 0

Sub Fail(message)
    If silent Then
        WScript.Echo message
    Else
        MsgBox message, vbExclamation, "CorelDRAW -> DXF"
    End If
    WScript.Quit 1
End Sub
