Attribute VB_Name = "TasksAndMemos"
Function taskRange() As Range
    Dim rng As Range
    Set rng = Range([addTask].Offset(1, 0), lastCell([addTask], , True))
    If rng.Cells.Count = 1 Then Set rng = Range([addTask], [addTask].Offset(37))
    Set taskRange = rng
End Function

Function memoRange() As Range
    Dim rng As Range
    Set rng = Range([AddMemo].Offset(1, 0), lastCell([AddMemo], , True))
    If rng.Cells.Count = 1 Then Set rng = Range([AddMemo], [AddMemo].Offset(37))
    Set memoRange = rng
End Function

Sub CompleteTasks()
If TypeName(Selection) <> "Range" Then Exit Sub
    Dim cell As Range
    Dim rng As Range
    For Each cell In Selection
        If Not Intersect(cell, taskRange) Is Nothing Then
            If cell <> "" Then
                If rng Is Nothing Then
                    Set rng = cell
                Else
                    Set rng = Union(rng, cell)
                End If
            End If
        End If
    Next
    If Not rng Is Nothing Then
        Dim var(0 To 2)
        For Each cell In rng
            var(0) = Sheets("Schedule").Range("M2")
            var(1) = Format(Now(), "dd-mm-yy HH:NN")
            var(2) = cell
            lastCell(Sheets("Completed").Range("A1"), , True).Resize(, 3).Value = var
        Next
        rng.Delete (xlUp)
        [addTask].Select
    Else
        msbox "Select tasks to mark as completed"
        Exit Sub
    End If
End Sub

Sub LoadElements(fromRange As Range, toSheet As Worksheet)
    Application.EnableEvents = False
    fromRange.Parent.Range(fromRange, Cells(Rows.Count, fromRange.Column)).ClearContents

    Dim cell As Range
    Dim ws As Worksheet
    Set ws = toSheet
    On Error Resume Next
    Set cell = ws.Rows(1).Find(fromRange.Parent.Range("mainUser").Value, , , xlWhole)
    On Error GoTo 0
    If cell Is Nothing Then
        Set cell = lastCell(ws.Range("A1"), True)        ' & ws.Range("A" & Columns.Count).End(xlToLeft).Column + 1)
        If cell.Address = Range("A1").Address Then
            If cell <> "" Then Set cell = cell.Offset(0, 1)
        End If
        cell = fromRange.Parent.Range("mainUser")
    End If
    Dim rng As Range
    Set rng = ws.Cells(2, cell.Column).Resize(10000)
    fromRange.Resize(rng.Rows.Count).Value = rng.Value
    Application.EnableEvents = True
End Sub

Sub AddElement(fromRange As Range, toSheet As Worksheet, inputRange As Range, Optional addToTop As Boolean)
    'fromRange is List's first cell
    Application.EnableEvents = False
    fromRange.Parent.Range(fromRange, Cells(Rows.Count, fromRange.Column)).ClearContents
    Application.EnableEvents = True
    Dim cell As Range
    Dim ws As Worksheet
    Set ws = toSheet
    On Error Resume Next
    Set cell = ws.Rows(1).Find(fromRange.Parent.Range("mainUser").Value, , , xlWhole)
    On Error GoTo 0
    If cell Is Nothing Then
        Set cell = lastCell(ws.Range("A1"), True)        ' & ws.Range("A" & Columns.Count).End(xlToLeft).Column + 1)
        If cell.Address = Range("A1").Address Then
            If cell <> "" Then Set cell = cell.Offset(0, 1)
        End If
        cell = fromRange.Parent.Range("mainUser")
    End If

    If addToTop = True Then
        cell.Offset(1).Insert (xlDown)
        Set cell = cell.Offset(1)
    Else
        Set cell.Value = lastCell(cell).Offset(1).Value
    End If

    cell.Value = inputRange
    Application.EnableEvents = False
    inputRange.ClearContents
    Application.EnableEvents = True
    LoadElements fromRange, toSheet
    inputRange.Select
End Sub

Sub UpdateElements(fromRange As Range, toSheet As Worksheet)

    Dim cell As Range
    Dim ws As Worksheet
    Set ws = toSheet
    On Error Resume Next
    Set cell = ws.Rows(1).Find(fromRange.Parent.Range("mainUser").Value, , , xlWhole)
    On Error GoTo 0
    If cell Is Nothing Then
        Set cell = lastCell(ws.Range("A1"), True)        ' & ws.Range("A" & Columns.Count).End(xlToLeft).Column + 1)
        If cell.Address = Range("A1").Address Then
            If cell <> "" Then Set cell = cell.Offset(0, 1)
        End If
        cell = fromRange.Parent.Range("mainUser")
    End If
    Dim rng As Range
    Set rng = ws.Cells(2, cell.Column).Resize(10000)
    rng.ClearContents

    rng.Value = fromRange.Resize(10000).Value
    DeleteEmptyCells rng
    LoadElements fromRange, toSheet
End Sub

Sub DeleteEmptyCells(rng As Range, Optional DIR As XlDirection = xlUp)
    Application.EnableEvents = False
    On Error Resume Next
    rng.SpecialCells(xlCellTypeBlanks).Delete Shift:=DIR
    Application.EnableEvents = True

End Sub

