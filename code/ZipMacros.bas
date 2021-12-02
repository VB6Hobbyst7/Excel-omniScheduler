Attribute VB_Name = "ZipMacros"

Sub ZipFiles(Optional FName As Variant, Optional FileNameZip As String)
'https://www.rondebruin.nl/win/s7/win001.htm
    Dim strDate As String, DefPath As String, sFName As String
    Dim oApp As Object, iCtr As Long, I As Integer
    Dim vArr

    DefPath = Application.DefaultFilePath
    If Right(DefPath, 1) <> "\" Then
        DefPath = DefPath & "\"
    End If

    strDate = Format(Now, " dd-mmm-yy h-mm-ss")
    If FileNameZip = "" Then FileNameZip = DefPath & "MyFilesZip " & strDate & ".zip"
    
    If IsMissing(FName) Then
        'Browse to the file(s), use the Ctrl key to select more files
        FName = Application.GetOpenFilename(filefilter:="Excel Files (*.xl*), *.xl*", _
                    MultiSelect:=True, Title:="Select the files you want to zip")
    End If
    If IsArray(FName) = False Then
        'do nothing
    Else
        'Create empty Zip File
        NewZip (FileNameZip)
        Set oApp = CreateObject("Shell.Application")
        I = 0
        For iCtr = LBound(FName) To UBound(FName)
            vArr = Split(FName(iCtr), "\")
            sFName = vArr(UBound(vArr))
            If bIsBookOpen(sFName) Then
                MsgBox "You can't zip a file that is open!" & vbLf & _
                       "Please close it and try again: " & FName(iCtr)
            Else
                'Copy the file to the compressed folder
                I = I + 1
                
                oApp.Namespace(CVar(FileNameZip)).CopyHere CVar(FName(iCtr))

                'Keep script waiting until Compressing is done
                On Error Resume Next
                Do Until oApp.Namespace(CVar(FileNameZip)).Items.Count = I
                    Application.Wait (Now + TimeValue("0:00:01"))
                Loop
                On Error GoTo 0
            End If
        Next iCtr

        'MsgBox "You find the zipfile here: " & FileNameZip
    End If
End Sub
Sub NewZip(sPath)
'Create empty Zip File
'Changed by keepITcool Dec-12-2005
    If Len(DIR(sPath)) > 0 Then Kill sPath
    Open sPath For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
End Sub


Function bIsBookOpen(ByRef szBookName As String) As Boolean
' Rob Bovey
    On Error Resume Next
    bIsBookOpen = Not (Application.Workbooks(szBookName) Is Nothing)
End Function


