Attribute VB_Name = "FunctionMacros"
Dim ShtNm, ScSum As String
Dim lastrow, ScRow, ScCol, StCol, EndCol, dbCol, AddRows As Long

Function AddNewSheet(ShtNm As String) As Worksheet
    Dim ws As Worksheet
    Set ws = Sheets.Add
    ws.Name = ShtNm
    Sheets("2021").Range("A1:NB1").Copy
    ws.Range("A1:NB1").PasteSpecial xlPasteAll
    ws.Range("A1") = "1/1/" & ShtNm
    Application.CutCopyMode = False
    ws.Visible = xlHidden
    Set AddNewSheet = ws
    Sheets("Schedule").Activate
End Function

Sub loadDayView()
    LoadDay [selDate], [dayCalendar].Cells(1, 1), [mainUser]
End Sub

Sub LoadDay(ofDate As Date, fromRange As Range, userNameRange As Range)
    'fromRange is first cell in list to display day items
    'LoadDayDynamic [M4],[m2],[staff]
    Dim ws As Worksheet
    Dim myYear As String
    myYear = CStr(Year(ofDate))

    On Error Resume Next
    Set ws = Sheets(myYear)
    On Error GoTo 0
    If ws Is Nothing Then Set ws = AddNewSheet(myYear)
           
        
    StopCode
    Dim cell As Range
    Set cell = RangeOfValue(userNameRange.Text, [staff])
    AddRows = 50 * (cell.Row - [staff].Row)        'Set Schedule Add Rows For Staff
    For Each cell In Sheets(CStr([scYear])).Rows(1).Cells
        If cell = ofDate Then ScCol = cell.Column: Exit For
    Next
    fromRange.Resize(37).ClearContents
    fromRange.Resize(37).Value = Range(Sheets(CStr([scYear])).Cells(2 + AddRows, ScCol), Sheets(CStr([scYear])).Cells(38 + AddRows, ScCol)).Value
    ResetCode
End Sub

Sub ModifyDay(ofDate As Date, fromRange As Range, userNameRange As Range)
    Dim myYear As String
    myYear = CStr(Year(ofDate))
    '?can remove next part
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Sheets(myYear)
    On Error GoTo 0
    If ws Is Nothing Then Set ws = AddNewSheet(myYear)
       
    'fromRange is first cell in list to display day items
    StopCode
    Dim cell As Range
    Set cell = RangeOfValue(userNameRange.Text, [staff])
    AddRows = 50 * (cell.Row - [staff].Row)        'Set Schedule Add Rows For Staff
    
    For Each cell In ws.Rows(1).Cells
        If cell = ofDate Then ScCol = cell.Column: Exit For
    Next
    Range(ws.Cells(2 + AddRows, ScCol), ws.Cells(38 + AddRows, ScCol)).Value = fromRange.Resize(37).Value
    ResetCode
End Sub

Function GetRangeFromShape(shapeString As String) As Range
    Dim rng As Range
    Select Case UCase(shapeString)
    Case "PRINTCONTROLS"
        Set rng = Range(shapeString)
    Case "YEARCALENDAR", "MONTHCALENDAR", "ADDTASK", "ADDMEMO"
        Set rng = Range(shapeString)
        Set rng = rng.Resize(, rng.Columns.Count + 1)
    Case "DAYCALENDAR", "WEEKCALENDAR"
        Set rng = Range(shapeString)
        Set rng = rng.Offset(-1, -1).CurrentRegion
        Set rng = rng.Resize(, rng.Columns.Count + 1)
    Case "DAYFORALL"
        Set rng = Range(shapeString)
        Set rng = rng.CurrentRegion
    Case "SETTINGS"
        Set rng = Range("A1:B1")
    Case Else
        '
    End Select
    Set GetRangeFromShape = rng
End Function

Function RangeExists(R As String) As Boolean
    Dim Test As Range
    On Error Resume Next
    Set Test = ActiveSheet.Range(R)
    RangeExists = Err.Number = 0
End Function

Function RangeOfValue(findWhat As String, _
                      findWhere As Range, _
                      Optional occurence As Integer, _
                      Optional offsetRow As Long = 0, _
                      Optional offsetCol As Long = 0, _
                      Optional lookIn As XlFindLookIn = xlValues, _
                      Optional lookAt As XlLookAt = xlWhole, _
                      Optional Order As XlSearchOrder = xlByRows, _
                      Optional Direction As XlSearchDirection = xlNext, _
                      Optional CaseSensitive As Boolean) As Range
    Dim cell As Range
    If occurence = 0 Then
        Set cell = findWhere.Find(findWhat, , lookIn, lookAt, Order, Direction, CaseSensitive)
        Set RangeOfValue = cell.MergeArea.Offset(offsetRow, offsetCol)
        '        For Each cell In findWhere
        '            If cell.Value = findWhat Then
        '                Set RangeOfValue = cell.MergeArea.Offset(offsetRow, offsetCol)
        '                Exit Function
        '            End If
        '        Next
    Else
        Set RangeOfValue = RangeFindNth(findWhere, findWhat, occurence)
    End If
End Function

Function RangeFindNth(rng As Range, strText As String, occurence As Integer) As Range
    Dim c As Range
    Dim counter As Integer
    For Each c In rng
        If InStr(1, c, strText) > 0 Then counter = counter + 1
        If counter = occurence Then
            Set RangeFindNth = c
            Exit Function
        End If
    Next c

End Function

Function lastCell(rng As Range, Optional booCol As Boolean, Optional onlyAfterFirstCell As Boolean) As Range
    Dim ws As Worksheet
    Set ws = rng.Parent
    Dim cell As Range
    If booCol = False Then
        Set cell = ws.Cells(Rows.Count, rng.Column).End(xlUp)
        If cell.MergeCells Then Set cell = Cells(cell.Row + cell.Rows.Count - 1, cell.Column)
    Else
        Set cell = ws.Cells(rng.Row, Columns.Count).End(xlToLeft)
        If cell.MergeCells Then Set cell = Cells(cell.Row, cell.Column + cell.Columns.Count - 1)
    End If
    
    If onlyAfterFirstCell = True Then
        If booCol = False Then
            Do While cell.Row <= rng.Row
                Set cell = cell.Offset(1, 0)
            Loop
        Else
            Do While cell.Column <= rng.Column
                Set cell = cell.Offset(0, 1)
            Loop
        End If
    End If
    
    Set lastCell = cell
End Function

Sub FindCF()

    'We have to Dim WhatIsIt as a generic Object instead of declaring as FormatCondition because DataBars screw things up.
    'See http://excelmatters.com/2015/03/04/when-is-a-formatcondition-not-a-formatcondition/

    Dim WhatIsIt As Object
    Dim fc As FormatCondition
    Dim db As Databar
    Dim cs As ColorScale
    Dim ics As IconSetCondition

    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        For Each WhatIsIt In ws.Cells.FormatConditions
            Select Case TypeName(WhatIsIt)
            Case "Databar":
                Set db = WhatIsIt
                Debug.Print "Type: " & "DataBar" & vbTab & "Applies To: " & db.AppliesTo.Address
            Case "FormatCondition"
                Set fc = WhatIsIt
                Debug.Print "Type: " & "FormatCondition" & vbTab & "Applies To: " & fc.AppliesTo.Address & vbTab & "Formula: " & fc.Formula1
            Case "ColorScale"
                Set cs = WhatIsIt
                Debug.Print "Type: " & "ColorScale" & vbTab & "Applies To: " & cs.AppliesTo.Address
            Case "IconSetCondition"
                Set ics = WhatIsIt
                Debug.Print "Type: " & "IconSet" & vbTab & "Applies To: " & ics.AppliesTo.Address
            Case Else
                Stop
            End Select
        Next WhatIsIt
    Next ws
End Sub



Sub FollowLink(folderPath As String)
    Dim oShell As Object
    Dim Wnd As Object
    Set oShell = CreateObject("Shell.Application")
    For Each Wnd In oShell.Windows
        If Wnd.Name = "File Explorer" Then
            If Wnd.Document.Folder.Self.Path = folderPath Then Exit Sub
        End If
    Next Wnd
    Application.ThisWorkbook.FollowHyperlink Address:=folderPath, NewWindow:=True
End Sub

Sub FoldersCreate(folderPath As String)
    'Create all the folders in a folder path
    Dim individualFolders() As String
    Dim tempFolderPath As String
    Dim arrayElement As Variant
    'Split the folder path into individual folder names
    individualFolders = Split(folderPath, "\")
    'Loop though each individual folder name
    For Each arrayElement In individualFolders
        'Build string of folder path
        tempFolderPath = tempFolderPath & arrayElement & "\"
        'If folder does not exist, then create it
        If DIR(tempFolderPath, vbDirectory) = "" Then
            MkDir tempFolderPath
        End If
    Next arrayElement
End Sub


Function CollectionToArray(c As Collection) As Variant
    Dim a() As Variant: ReDim a(0 To c.Count - 1)
    Dim I As Long
    For I = 1 To c.Count
        a(I - 1) = c.Item(I)
    Next
    CollectionToArray = a
End Function

