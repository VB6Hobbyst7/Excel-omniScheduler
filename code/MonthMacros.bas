Attribute VB_Name = "MonthMacros"
Sub AddSummary(ofDate As Date, fromRange As Range, timeColumn As Integer, userNameRange As Range)
    'AddSummary(ofDate As Date, fromRange As Range, timeColumn as integer, userNameRange As Range)
    'for day view ->    [dayCalendar].Column - 1

    Set cell = RangeOfValue(CStr(userNameRange), [staff])
    AddRows = 50 * (cell.Row - [staff].Row)
    Dim dbCol As Integer
    Dim element As Range
    For Each element In Sheets(CStr([scYear])).Cells(1, 1).CurrentRegion.Resize(1)
        If element.Text = Format(ofDate, "DD-MM-YY") Then
            Set cell = element
            Exit For
        End If
    Next
    
    'Set cell = RangeOfValue(Format(ofDate, "DD-MM-YY"), Sheets(CStr([ScYear])).Rows(1), , , , xlValues)
    dbCol = cell.Column
    With Sheets("Schedule")
        ScSum = Empty
        For ScRow = fromRange.Row To fromRange.Row + 36
            If .Cells(ScRow, fromRange.Column).Value <> Empty Then
                'todo add variable for time value range instead of [dayCalendar].Column - 1
                ScSum = ScSum & Format(.Cells(ScRow, timeColumn).Value, "h:mma/p") & ": " & .Cells(ScRow, fromRange.Column).Value & vbCrLf
            End If
        Next ScRow
        Sheets(CStr([scYear])).Cells(40 + AddRows, dbCol).Value = ScSum
        .Range([selMonthDay]).Value = ScSum
    End With
    
    
End Sub


Function MergedCellsOfRange(fullRange As Range) As Range
    Dim c As Range
    Dim mergedCells As Range
    For Each c In fullRange
        If c.MergeCells = True Then
            If mergedCells Is Nothing Then
                Set mergedCells = c
            Else
                Set mergedCells = Union(mergedCells, c)
            End If
        End If
    Next
    If Not mergedCells Is Nothing Then
        Set MergedCellsOfRange = mergedCells
    End If
End Function

Sub LoadMonth()
    StopCode
    With Sheets("Schedule")
        .Calculate
        AddRows = 50 * (RangeOfValue([mainUser], [staff]).Row - [staff].Row)        'Set Schedule Add Rows For Staff        If .Range("B5").Value = "" Then AddNewSheet
        .Calculate
        ShtNm = .Range("ScYear").Value
        dbCol = .Range("StartCol").Value        'First Column
        
        MergedCellsOfRange([MonthCalendar]).ClearContents
        '.Range("D6:J10,D12:J16,D18:J22,D24:J28,D30:J34,D36:J40").ClearContents
        'For ScRow = 6 To 36 Step 6
        For ScRow = [MonthCalendar].Row To [MonthCalendar].Row + 30 Step 6
            'For ScCol = 4 To 10
            For ScCol = [MonthCalendar].Column To [MonthCalendar].Column + 6
                If .Cells(ScRow - 1, ScCol).Value = Empty Then GoTo NextCol
                .Cells(ScRow, ScCol).Value = Sheets(CStr([scYear])).Cells(40 + AddRows, dbCol).Value
                dbCol = dbCol + 1
NextCol:
            Next ScCol
        Next ScRow
    End With
    ResetCode
End Sub

Sub MonthSel()

    Dim ws As Worksheet
    Set ws = ActiveSheet.Shapes(Application.Caller).TopLeftCell.Parent.Parent

    ws.Shapes("MonthBtns").ShapeStyle = msoShapeStylePreset41
    ws.Shapes(Application.Caller).ShapeStyle = msoShapeStylePreset27
    ws.Range("selmonth").Value = Sheets("Schedule").Shapes(Application.Caller).TextFrame2.TextRange.Text
    LoadMonth

End Sub

Sub MonthChange()
    Dim ws As Worksheet
    Set ws = Sheets("Schedule")

    ws.Shapes("MonthBtns").ShapeStyle = msoShapeStylePreset41
    Dim shp As Shape
    For Each shp In ws.Shapes("MonthBtns").GroupItems
        If shp.TextFrame2.TextRange.Text = ws.[SelMonth] Then
            shp.ShapeStyle = msoShapeStylePreset27
        End If
    Next
    LoadMonth

End Sub

