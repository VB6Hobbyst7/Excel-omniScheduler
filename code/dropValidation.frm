VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dropValidation 
   Caption         =   "DropDown Validation"
   ClientHeight    =   2952
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   2904
   OleObjectBlob   =   "dropValidation.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dropValidation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    If Me.Height > 224 Then
        Me.Height = 175
        CommandButton1.Caption = "V"
    Else
        Me.Height = 225
        CommandButton1.Caption = "Ë"
    End If
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim lngValType As Long
    On Error Resume Next
    lngValType = ActiveCell.Validation.Type
    On Error GoTo 0
    If lngValType = 3 Then updateData
    
End Sub

Private Sub ListBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        updateData
    End If
End Sub

Private Sub TextBox1_Change()
    RefreshList
End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        updateData
    End If
End Sub

Private Sub UserForm_initialize()
    RefreshList
End Sub

Sub updateData()
    ActiveCell.Value = Me.ListBox1.Value
    If Me.oClose.Value = True Then
        Me.TextBox1.Value = ""
        Unload Me
    ElseIf Me.oDown.Value = True Then
        ActiveCell.Offset(1, 0).Select
        If Me.TextBox1.Value <> "" Then
            Me.TextBox1.Value = ""
        Else
            RefreshList
        End If
        If Me.ListBox1.ListCount = 0 Then Unload Me
        
    ElseIf Me.oRight.Value = True Then
        ActiveCell.Offset(0, 1).Select
        If Me.TextBox1.Value <> "" Then
            Me.TextBox1.Value = ""
        Else
            RefreshList
        End If
        If Me.ListBox1.ListCount = 0 Then Unload Me
    ElseIf Me.oNone.Value = True Then
        Me.TextBox1.Value = ""
    End If
End Sub

Sub RefreshList()
    Dim arr() As String
    Dim I As Integer
    Me.ListBox1.Clear
    Dim rng As Range
    Dim cell As Range
    If isValidation(ActiveCell) = False Then Exit Sub
    If ValidRange(ActiveCell.Validation.Formula1) = True Then
        Set rng = Range(Replace(ActiveCell.Validation.Formula1, "=", ""))
        For Each cell In rng
            If Me.TextBox1.Value = "" Then
                Me.ListBox1.AddItem cell.Value
            Else
                If InStr(cell.Value, UCase(Me.TextBox1.Value)) > 0 Then
                    Me.ListBox1.AddItem cell.Value
                End If
            End If
        Next
    Else
        Unload Me
    End If
    '    arr = VBA.Split(ActiveCell.Validation.Formula1, ",")
    '
    '    For i = LBound(arr) To UBound(arr)
    '        If Me.TextBox1.Value = "" Then
    '            Me.ListBox1.AddItem arr(i)
    '        Else
    '            If VBA.InStr(arr(i), UCase(Me.TextBox1.Value) > 0) Then
    '                Me.ListBox1.AddItem arr(i)
    '            End If
    '        End If
    '    Next
    On Error Resume Next
    Me.ListBox1.ListIndex = 0
    On Error GoTo 0

End Sub

Function ValidRange(str As String) As Boolean
    Dim rng As Range
    On Error Resume Next
    Set rng = Range(Replace(str, "=", ""))
    On Error GoTo 0
    If rng Is Nothing Then
        ValidRange = False
    Else
        ValidRange = True
    End If
End Function

Function isValidation(rng As Range) As Boolean
    Dim dvtype As Integer
    On Error Resume Next
    dvtype = rng.Validation.Type
    On Error GoTo 0
    If dvtype = 3 Then
        isValidation = True
    Else
        isValidation = False
    End If
End Function

