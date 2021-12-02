Attribute VB_Name = "CalGoogleSyncMacros"
Option Explicit

Sub SendToGoogleCalendar()
    Dim objHTTP As Object
    Dim Json, URL, ApName, WebhookID, ApTime As String
    Dim ApStart, ApEnd As Date
    Dim StaffRow, SelRow, ApDur  As Long
    With Sheets("Schedule")
        If .Range("B8").Value = Empty Then        'Check for correct Staff
            MsgBox "Please select a correct Staff"
            Exit Sub
        End If
        StaffRow = .Range("B8").Value        'Staff Row
        WebhookID = Sheet3.Range("F" & StaffRow).Value
        If WebhookID = "" Then
            MsgBox "Please assign a Zapier Webhook ID to this staff"
            Sheet3.Activate
            Sheet3.Range("F" & StaffRow).Select
            Exit Sub
        End If
        SelRow = .Range("B1").Value        'Get Selected Row
        ApDur = .Range("defApDuration").Value        'Set Duration
        ApStart = .Range("M3").Value + .Range("L" & SelRow).Value        'Combine Date & Time
        ApEnd = ApStart + (ApDur * 0.000695)
        ApName = .Range("M" & SelRow).Value

        Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
        URL = "[Add Your Webhook Link Here]" & WebhookID & "/?ApName=" & ApName & "&ApStart=" & ApStart & "&ApEnd=" & ApEnd
        objHTTP.Open "PATCH", URL, False
        objHTTP.setRequestHeader "Content-type", "application/json"
        objHTTP.SEND (Json)        'Send Information
    End With
End Sub

