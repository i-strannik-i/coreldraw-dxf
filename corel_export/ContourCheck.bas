Attribute VB_Name = "ContourCheck"
Option Explicit

Private Function NeedsClosed(ByVal layerName As String) As Boolean
    Dim n As String
    n = UCase$(Trim$(layerName))
    If Left$(n, 5) = "BACK_" Then n = Mid$(n, 6)
    NeedsClosed = (n = "OUT" Or n = "IN" Or n = "CUT_OUT" Or n = "CUT_IN" Or n = "P" Or n = "D" Or Left$(n, 2) = "P_")
    If Left$(n, 1) = "D" And Len(n) > 1 Then
        If IsNumeric(Replace(Mid$(n, 2), ".", ",")) Or IsNumeric(Mid$(n, 2)) Then NeedsClosed = True
    End If
End Function

Private Sub Scan(ByVal items As Object, ByVal marked As ShapeRange, ByRef total As Long, ByRef opened As Long, ByRef allowed As Long, ByRef zero As Long, ByRef skipped As Long)
    Dim s As Shape, p As SubPath, seg As Segment, bad As Boolean, unitScale As Double
    unitScale = Application.ConvertUnits(1, ActiveDocument.Unit, cdrMillimeter)
    For Each s In items
        If s.Type = cdrGroupShape Then
            Scan s.Shapes, marked, total, opened, allowed, zero, skipped
        ElseIf UCase$(Trim$(s.Layer.Name)) <> "INFO" Then
            bad = False
            If s.Type = cdrCurveShape Then
                For Each p In s.Curve.SubPaths
                    total = total + 1
                    If Not p.Closed Then
                        If NeedsClosed(s.Layer.Name) Then
                            opened = opened + 1
                            bad = True
                        Else
                            allowed = allowed + 1
                        End If
                    End If
                    For Each seg In p.Segments
                        If seg.Length * unitScale <= 0.000000001 Then
                            zero = zero + 1
                            bad = True
                        End If
                    Next seg
                Next p
            ElseIf s.Type = cdrRectangleShape Or s.Type = cdrEllipseShape Then
                total = total + 1
            Else
                skipped = skipped + 1
            End If
            If bad Then marked.Add s
        End If
    Next s
End Sub

Public Function CheckContours(ByVal selectionOnly As Boolean, Optional ByVal showResults As Boolean = True) As String
    On Error GoTo Failed
    If Application.Documents.Count = 0 Then Err.Raise 513, , "Нет открытого документа."
    Dim items As Object, marked As ShapeRange
    Dim total As Long, opened As Long, allowed As Long, zero As Long, skipped As Long
    If selectionOnly Then
        Set items = ActiveSelectionRange
        If items.Count = 0 Then Err.Raise 513, , "Нет выделенных объектов."
    Else
        Set items = ActivePage.Shapes
    End If
    Set marked = Application.CreateShapeRange
    Scan items, marked, total, opened, allowed, zero, skipped
    CheckContours = "Проверено контуров: " & total & vbCrLf & _
        "Незамкнутых для резки/выборки: " & opened & vbCrLf & _
        "Нулевых участков: " & zero & vbCrLf & _
        "Открытых на других слоях (допустимо): " & allowed & vbCrLf & _
        "Пропущено не-кривых объектов/текста: " & skipped & vbCrLf & vbCrLf & _
        "Проблемных объектов: " & marked.Count & ". При наличии они выделяются." & vbCrLf & _
        "Геометрия не изменена." & vbCrLf & _
        "Допуск экспорта 0,1 мм. Автозамыкания нет." & vbCrLf & _
        "Пересечения и наложения этой проверкой не проверяются."
    If showResults Then
        If marked.Count > 0 Then marked.CreateSelection
        MsgBox CheckContours, vbInformation, "Проверка контуров DXF"
    End If
    Exit Function
Failed:
    CheckContours = "Ошибка проверки: " & Err.Description
    If showResults Then MsgBox CheckContours, vbExclamation, "Проверка контуров DXF"
End Function
