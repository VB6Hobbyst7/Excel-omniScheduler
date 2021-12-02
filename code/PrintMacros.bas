Attribute VB_Name = "PrintMacros"
Sub printRange(rng As Range, Optional toPDF As Boolean, Optional fileFullName As String)
    SetupPage rng.Parent, xlPaperA4, True, True, rng.Width > rng.Height
    If toPDF = False Then
        rng.PrintPreview
    Else
        rng.ExportAsFixedFormat Type:=xlTypePDF, fileName:=fileFullName
    End If
End Sub

Sub SetupPage(ws As Worksheet, _
              Optional ePaperSize As XlPaperSize = xlPaperA4, _
              Optional booFitWide As Boolean, _
              Optional booFitTall As Boolean, _
              Optional isLandscape As Boolean)
    On Error Resume Next
    With ws.PageSetup
        .LeftMargin = Application.InchesToPoints(0)
        .RightMargin = Application.InchesToPoints(0)
        .TopMargin = Application.InchesToPoints(0)
        .BottomMargin = Application.InchesToPoints(0)
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0)
        If isLandscape = True Then
            .Orientation = xlLandscape
        Else
            .Orientation = xlPortrait
        End If
        .PaperSize = ePaperSize
        .Zoom = False
        If booFitWide = True Then .FitToPagesWide = 1
        If booFitTall = True Then .FitToPagesTall = 1
    End With
End Sub

Function NamesInRange(rng As Range) As Collection
    Dim nm As Name
    Dim nmStr As String
    Dim Collect As New Collection
    For Each nm In ThisWorkbook.Names
        nmStr = nm.Name
        Dim RefersToRange As String
        On Error Resume Next
        RefersToRange = Mid(nm.RefersTo, InStr(1, nm.RefersTo, "$"))
        On Error GoTo 0
        If RangeExists(RefersToRange) Then
            If Not Intersect(rng, Range(RefersToRange)) Is Nothing Then Collect.Add nmStr, nmStr
        End If
    Next nm

    Set NamesInRange = Collect

    Dim element As Variant
    Debug.Print Collect.Count & " Names intersect with range " & rng.Address
    For Each element In Collect
        Debug.Print vbTab & element
    Next
End Function

Sub ToggleShapePrint()
    Dim outputPath As String
    outputPath = ThisWorkbook.Path & "\ScheduleOutput"
    FoldersCreate outputPath
    Dim element
    Dim fileFullName As String

    Dim fileCollection As New Collection
    Dim rangeCollection As New Collection

    Dim oCaller As Shape
    Set oCaller = ActiveSheet.Shapes(Application.Caller)
    Dim shapeText As String
    shapeText = oCaller.TextFrame2.TextRange.Text
    If InStr(1, shapeText, vbLf) > 0 Then
        Dim var
        var = Split(shapeText, vbLf)
        For Each element In var
            If RangeExists(CStr(element)) Then
                fileFullName = outputPath & "\" & Format(Now, "yymmdd hhnnss") & " " & CStr(element) & ".pdf"
                fileCollection.Add fileFullName
                rangeCollection.Add CStr(element)
            End If
        Next
    ElseIf UCase(shapeText) = "VISIBLE" Then
        Dim cell As Range
        Set cell = [PrintControls].Offset(0, 1)
        Dim rng As Range
        
        Set rng = VisibleRange 'Range(cell, lastCell(cell, True)).Resize(10).SpecialCells(xlCellTypeVisible)
        'if SeparateRanges = true then
        
            If WorksheetFunction.CountA(rng) <> 0 Then
                For Each element In NamesInRange(rng)
                    If UCase(element) <> "MAINUSER" Then
                        fileFullName = outputPath & "\" & Format(Now, "yymmdd hhnnss") & " " & CStr(element) & ".pdf"
                        fileCollection.Add fileFullName
                        rangeCollection.Add CStr(element)
                    End If
                Next
            End If
        
        
    Else
        If RangeExists(shapeText) Then
            printRange GetRangeFromShape(shapeText).EntireColumn
        End If
    End If
    
    If fileCollection.Count > 0 Then
        Dim counter As Integer
        For counter = 1 To fileCollection.Count
            printRange GetRangeFromShape(rangeCollection(counter)).EntireColumn, True, fileCollection(counter)
        Next
        
        If [MailSchedule].Value = "YES" Then
            If OutlookExists = True Then
                Dim mailto As String
                mailto = RangeOfValue([emailTo], [staff], , , 1)
                Dim fileArray As Variant: fileArray = CollectionToArray(fileCollection)
                
                If [ZipTheFiles].Value = "YES" Then
                    Dim zipOut As String: zipOut = outputPath & "\" & fileCollection.Count & " ZipedSchedules.zip"
                    ZipFiles fileArray, zipOut
                End If
                
                SendEmail mailto, "Schedules for " & [emailTo], _
                          "Hello " & [emailTo] & "." & vbNewLine & vbNewLine & _
                          "Here are your schedules. Contact us if you have any questions.", _
                          True, , IIf([ZipTheFiles].Value = "YES", zipOut, fileArray)
            Else
                MsgBox "Outlook not found"
                Exit Sub
            End If
        Else
            FollowLink outputPath
        End If
    End If
End Sub

Function VisibleRange() As Range
Dim cell As Range
Set cell = [PrintControls]
Dim rng As Range
Set rng = Range(cell, Cells(Rows.Count, Columns.Count))
Set rng = Range(cell.Offset(-2, 1), Cells(Last(1, rng), Last(2, rng)))
Set rng = rng.SpecialCells(xlCellTypeVisible)
Set VisibleRange = rng

End Function



Function Last(choice As Long, rng As Range)
    'Ron de Bruin, 5 May 2008
    ' 1 = last row
    ' 2 = last column
    ' 3 = last cell
    Dim lrw As Long
    Dim lCol As Long

    Select Case choice

    Case 1:
        On Error Resume Next
        Last = rng.Find(what:="*", _
                        After:=rng.Cells(1), _
                        lookAt:=xlPart, _
                        lookIn:=xlFormulas, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False).Row
        On Error GoTo 0

    Case 2:
        On Error Resume Next
        Last = rng.Find(what:="*", _
                        After:=rng.Cells(1), _
                        lookAt:=xlPart, _
                        lookIn:=xlFormulas, _
                        SearchOrder:=xlByColumns, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False).Column
        On Error GoTo 0

    Case 3:
        On Error Resume Next
        lrw = rng.Find(what:="*", _
                       After:=rng.Cells(1), _
                       lookAt:=xlPart, _
                       lookIn:=xlFormulas, _
                       SearchOrder:=xlByRows, _
                       SearchDirection:=xlPrevious, _
                       MatchCase:=False).Row
        On Error GoTo 0

        On Error Resume Next
        lCol = rng.Find(what:="*", _
                        After:=rng.Cells(1), _
                        lookAt:=xlPart, _
                        lookIn:=xlFormulas, _
                        SearchOrder:=xlByColumns, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False).Column
        On Error GoTo 0

        On Error Resume Next
        Last = rng.Parent.Cells(lrw, lCol).Address(False, False)
        If Err.Number > 0 Then
            Last = rng.Cells(1).Address(False, False)
            Err.Clear
        End If
        On Error GoTo 0

    End Select
End Function



