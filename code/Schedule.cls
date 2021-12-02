VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim dbCol As Long
Dim DbSht As String

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet
    
    If Not Intersect(Target, Range("scyear")) Is Nothing Then
        On Error Resume Next
        Set ws = Sheets(CStr([scYear]))
        On Error GoTo 0
        If ws Is Nothing Then Set ws = AddNewSheet(CStr([scYear]))
    End If
    
    If Not Intersect(Target, Range("mainUser")) Is Nothing Then        'Refresh Both Month And Day on Staff Change
        LoadMonth
        loadDayView
        loadWeekDays
        LoadElements [addTask].Offset(1), Sheets("Tasks")
        LoadElements [AddMemo].Offset(1), Sheets("Memo")
    End If
    
    If Not Intersect(Target, Range("dayCalendar")) Is Nothing Then        'And Range("B2").Value = False Then
        If Range("DbCol").Value = "" Then AddNewSheet ([dayCalendar].Cells(1, 1).Offset(-1))

        ModifyDay [selDate], [dayCalendar].Cells(1, 1), [mainUser]
        loadWeekDays
        LoadDayForAll
        
        Range("calSyncRow").Value = Target.Row        'Set Row for Google & Outlook Calendar Sync

        '        If Range("sendOutlook").Value = "Yes" Then SendToOultook        'Send to Outlook when opted in
        '        If Range("sendGoogle").Value = "Yes" Then SendToGoogleCalendar
        AddSummary [dayCalendar].Cells(1, 1).Offset(-1), [dayCalendar].Cells(1, 1), [dayCalendar].Column - 1, [mainUser]
    End If

    If Not Intersect(Target, Range("weekCalendar")) Is Nothing Then
        
        On Error Resume Next
        Set ws = Sheets(CStr(Year(Cells([weekCalendar].Row - 1, Target.Column))))
        On Error GoTo 0
        If ws Is Nothing Then Set ws = AddNewSheet(CStr(Year(Cells([weekCalendar].Row - 1, Target.Column))))
        
        
        ModifyWeek
        loadDayView
        LoadDayForAll
        LoadMonth
    End If
    
   
    If [DayForAll].Offset(0, 1) <> "" Then
        If Not Intersect(Target, DayForAllRange) Is Nothing Then
            ModifyDayForAll
            loadDayView
            loadWeekDays
        End If
        If Not Intersect(Target, DayForAllUserRange) Is Nothing Then
            ModifyDayForAll
            loadDayView
            loadWeekDays
        End If
    End If
    
    StopCode
    'Task Add
    If Not Intersect(Target, [addTask]) Is Nothing Then AddElement [addTask].Offset(1), Sheets("Tasks"), [addTask], True
    'Task Edit
    If Not Intersect(Target, taskRange) Is Nothing Then UpdateElements [addTask].Offset(1), Sheets("Tasks")
    'Memo Add
    If Not Intersect(Target, [AddMemo]) Is Nothing Then AddElement [AddMemo].Offset(1), Sheets("Memo"), [AddMemo], True
    'Memo Edit
    If Not Intersect(Target, memoRange) Is Nothing Then UpdateElements [AddMemo].Offset(1), Sheets("Memo")
    ResetCode
    
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Cells.Count > 6 Then Exit Sub

    If Not Intersect(Target, [yearCalendar]) Is Nothing Then
        If Target <> "" And IsDate(Target) Then
            [selDate] = Target
            [SelMonth] = Month([selDate])
            MonthChange
            LoadDayForAll
        End If
    End If

    If Not Intersect(Target, [MonthCalendar]) Is Nothing Then        'MergedCellsOfRange([monthCalendar])
        If Target.Offset(-1) <> "" Then [selDate].Value = Target.Offset(-1)
        LoadDay [selDate], [dayCalendar].Cells(1, 1), [mainUser]
        [selMonthDay].Value = Target.Address
        LoadDayForAll
    End If
End Sub

