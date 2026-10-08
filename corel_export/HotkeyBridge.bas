Attribute VB_Name = "HotkeyBridge"
Option Explicit
Private Declare PtrSafe Function GetKeyboardState Lib "user32" (ByRef keys As Byte) As Long
Private Declare PtrSafe Function SetKeyboardState Lib "user32" (ByRef keys As Byte) As Long
Private Const ExportGuid As String = "87260fe1-e755-4c95-9f68-2aa2f3dc5001"

Private Function Source() As Object
    Dim ds As Object, xml As Object, node As Object, button As Object
    Set ds = Application.FrameWork.Application.DataContext.GetDataSource("CustomizationDS")
    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    If Not xml.LoadXML(CStr(ds.GetProperty("CommandCategoriesItems"))) Then Err.Raise 513, , "Cannot read command categories"
    Set node = xml.SelectSingleNode("//itemData[@guid='2cc24a3e-fe24-4708-9a74-9c75406eebcd']")
    If node Is Nothing Then Err.Raise 513, , "Macro category not found"
    ds.SetProperty "SelectedCommandCategoryItem", node.xml
    If Not xml.LoadXML(CStr(ds.GetProperty("CommandItems"))) Then Err.Raise 513, , "Cannot read commands"
    Set node = xml.SelectSingleNode("//itemData[@guid='" & ExportGuid & "']")
    If node Is Nothing Then
        On Error Resume Next
        Set button = Application.CommandBars("Sklad DXF").Controls.Item(1)
        On Error GoTo 0
        If Not button Is Nothing Then
            If button.Caption = "DXF" Or button.Caption = "Экспорт DXF" Then Set node = xml.SelectSingleNode("//itemData[@guid='" & button.ID & "']")
        End If
    End If
    If node Is Nothing Then
        If xml.SelectNodes("//itemData[@text='ExportDXF.ShowExporter']").Length = 1 Then
            Set node = xml.SelectSingleNode("//itemData[@text='ExportDXF.ShowExporter']")
        End If
    End If
    If node Is Nothing Then Err.Raise 513, , "DXF command not installed"
    ds.SetProperty "SelectedCommandItem", node.xml
    Set Source = ds
End Function

Private Sub Capture(ByVal ds As Object, ByVal vk As Long, ByVal modifiers As Long)
    Dim original(0 To 255) As Byte, simulated(0 To 255) As Byte, i As Long, failure As String
    If GetKeyboardState(original(0)) = 0 Then Err.Raise 513, , "Cannot read keyboard state"
    For i = 0 To 255
        simulated(i) = original(i) And 1
    Next i
    If (modifiers And 1) <> 0 Then simulated(16) = 128
    If (modifiers And 2) <> 0 Then simulated(17) = 128
    If (modifiers And 4) <> 0 Then simulated(18) = 128
    If SetKeyboardState(simulated(0)) = 0 Then Err.Raise 513, , "Cannot prepare shortcut"
    On Error GoTo Failed
    ds.SetProperty "VirtualKey", vk
    SetKeyboardState original(0)
    Exit Sub
Failed:
    failure = Err.Description
    SetKeyboardState original(0)
    Err.Raise 513, , failure
End Sub

Public Function Preview(ByVal vk As Long, ByVal modifiers As Long) As String
    On Error GoTo Failed
    Dim ds As Object, label As String, assigned As String
    Set ds = Source()
    Capture ds, vk, modifiers
    label = CStr(ds.GetProperty("NewShortcutKey"))
    assigned = CStr(ds.GetProperty("CurrentlyAssignedTo"))
    If Len(label) = 0 Then Err.Raise 513, , "Corel shortcut service is not ready"
    Preview = "OK" & vbTab & label & vbTab & assigned
    Exit Function
Failed:
    Preview = "ERROR" & vbTab & Err.Description
End Function

Public Function Assign(ByVal vk As Long, ByVal modifiers As Long, ByVal expected As String) As String
    On Error GoTo Failed
    Dim check As String, ds As Object, xml As Object, node As Object, label As String
    check = Preview(vk, modifiers)
    If check <> expected Then Err.Raise 513, , "Assignment changed. Check the shortcut again."
    Set ds = Source()
    Capture ds, vk, modifiers
    label = CStr(ds.GetProperty("NewShortcutKey"))
    If Not CBool(ds.GetProperty("IsAssignButtonEnable")) Then Err.Raise 513, , "Assignment is unavailable"
    ds.SetProperty "GoToConflict", False
    ds.InvokeMethod "OnAssignButton"
    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    If Not xml.LoadXML(CStr(ds.GetProperty("ShortcutKeyList"))) Then Err.Raise 513, , "Cannot verify assignment"
    For Each node In xml.SelectNodes("//itemData")
        If node.getAttribute("text") = label Then
            Assign = "OK" & vbTab & label
            Exit Function
        End If
    Next node
    Err.Raise 513, , "Assignment needs verification in Corel"
Failed:
    Assign = "ERROR" & vbTab & Err.Description
End Function
