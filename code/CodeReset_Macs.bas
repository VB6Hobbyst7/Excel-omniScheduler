Attribute VB_Name = "CodeReset_Macs"
Option Explicit

Sub ResetCode()
    With Application
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With
End Sub

Sub StopCode()
    With Application
        .EnableEvents = False
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With
End Sub

