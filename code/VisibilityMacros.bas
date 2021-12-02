Attribute VB_Name = "VisibilityMacros"
Sub ToggleShapeNamedRange()
    StopCode
    Dim shp As Shape
    Dim oCaller As Shape
    Set oCaller = ActiveSheet.Shapes(Application.Caller)
    Dim shapeText As String
    shapeText = oCaller.TextFrame2.TextRange.Text
    shapeText = UCase(shapeText)
    If InStr(1, shapeText, vbLf) > 0 Then
        HideAllRanges
        Dim var
        var = Split(shapeText, vbLf)
        Dim element
        For Each element In var
            If RangeExists(CStr(element)) Then
                ToggleNamedRange CStr(element)
            End If
        Next
    ElseIf shapeText = "SHOWALL" Then ShowAllRanges
    ElseIf shapeText = "HIDEALL" Then HideAllRanges
    ElseIf shapeText = "SETTINGS" Then ToggleNamedRange shapeText
    Else
        If RangeExists(shapeText) Then ToggleNamedRange shapeText
    End If
    AdaptShapeFill
    ResetCode
End Sub

Sub HideAllRanges()
    Range([PrintControls].Offset(0, 2), lastCell([PrintControls], True)).EntireColumn.Hidden = True
End Sub

Sub ShowAllRanges()
    Range([PrintControls].Offset(0, 2), Cells([PrintControls].Row, Columns.Count)).EntireColumn.Hidden = False
End Sub

Sub ToggleNamedRange(shapeString As String)
    Dim rng As Range
    Set rng = GetRangeFromShape(shapeString)
    If rng.EntireColumn.Hidden = False Then
        rng.EntireColumn.Hidden = True
    Else
        rng.EntireColumn.Hidden = False
    End If
End Sub

Sub AdaptShapeFill()
    Dim rng As Range
    Dim shp As Shape
    Dim shapeText As String
                
    For Each shp In ActiveSheet.Shapes("visibilityShapeGroup").GroupItems

        shapeText = shp.TextFrame2.TextRange.Text
        If RangeExists(shapeText) Or UCase(shapeText) = "SETTINGS" Then
            Set rng = GetRangeFromShape(shapeText)
        
            If rng.EntireColumn.Hidden = False Then
                shp.Fill.ForeColor.RGB = 32768
            Else
                shp.Fill.ForeColor.RGB = 128
            End If
                    
        End If
    Next
End Sub

