Attribute VB_Name = "WeekMacros"
Sub ModifyWeek()
    Dim cell As Range
    For Each cell In [weekCalendar].Resize(1)
        ModifyDay cell.Offset(-1), cell, [mainUser]
        
        AddSummary cell.Offset(-1), cell, [weekCalendar].Column - 1, [mainUser]
        
    Next
    '    loadDayView
End Sub

Sub weekIncrement()
    Application.ScreenUpdating = False
    Dim shp As Shape
    Set shp = ActiveSheet.Shapes(Application.Caller)
    Select Case UCase(shp.Name)
    Case "NEXTWEEK"
        [weekOffset] = [weekOffset] + 7
    Case "PREVIOUSWEEK"
        [weekOffset] = [weekOffset] - 7
    Case "RESETYEAROFFSET"
        [weekOffset] = 0
    End Select
    loadWeekDays
    Application.ScreenUpdating = True
End Sub

Sub YearIncrement()
    Application.ScreenUpdating = False
    Dim shp As Shape
    Set shp = ActiveSheet.Shapes(Application.Caller)
    Select Case UCase(shp.Name)
    Case "NEXTYEAR"
        [scYear] = [scYear] + 1
    Case "PREVIOUSYEAR"
        [scYear] = [scYear] - 1
    Case "RESETYEAROFFSET"
        [scYear] = Year(Now)
    End Select
    Application.ScreenUpdating = True
End Sub

Sub loadWeekDays()
    Dim cell As Range
    For Each cell In [weekCalendar].Resize(1)
        LoadDay cell.Offset(-1), cell, [mainUser]
    Next
End Sub

