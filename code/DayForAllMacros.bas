Attribute VB_Name = "DayForAllMacros"
Function DayForAllUserRange() As Range
    Dim rng As Range
    Dim cell As Range
    Set cell = [DayForAll].Offset(0, 1)
    Set rng = Range(cell, cell.Resize(1, [DayForAll].CurrentRegion.Columns.Count - 1))
    Set DayForAllUserRange = rng
End Function

Sub LoadDayForAll()
    Dim rng As Range
    Dim cell As Range
    Set cell = [DayForAll].Offset(2, 1)
    Set rng = Range(cell, cell.Resize(1, [DayForAll].CurrentRegion.Columns.Count - 1))
    For Each cell In rng
        LoadDay cell.Offset(-1), cell, cell.Offset(-2)
    Next
    loadWeekDays

End Sub

Sub ModifyDayForAll()
    Dim rng As Range
    Dim cell As Range
    Set cell = [DayForAll].Offset(2, 1)
    Set rng = Range(cell, cell.Resize(1, [DayForAll].CurrentRegion.Columns.Count - 1))
    For Each cell In rng
        ModifyDay cell.Offset(-1), cell, cell.Offset(-2)
        
        AddSummary cell.Offset(-1), cell, [DayForAll].Column, cell.Offset(-2)
        
    Next
    loadWeekDays

End Sub

Function DayForAllRange() As Range
    Dim rng As Range
    Dim cell As Range
    Set cell = [DayForAll].Offset(2, 1)
    Set rng = Range(cell, cell.Resize(1, [DayForAll].CurrentRegion.Columns.Count - 1).Offset([DayForAll].CurrentRegion.Rows.Count - 3))
    Set DayForAllRange = rng
End Function

Sub addUserDayForAll()
    StopCode
    Dim rng As Range
    Set rng = [DayForAll].Offset(0, 1).Resize([DayForAll].CurrentRegion.Rows.Count)
    
    shapeText = UCase(ActiveSheet.Shapes(Application.Caller).Name)
    Select Case shapeText
    Case Is = "ADDUSERVIEW"
        rng.Insert Shift:=xlToRight
        [mainUser].Resize([mainUser].CurrentRegion.Rows.Count).Copy rng.Offset(0, -1)
    Case Is = "REMOVEUSERVIEW"
        
    End Select
    ResetCode
End Sub

