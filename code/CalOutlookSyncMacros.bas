Attribute VB_Name = "CalOutlookSyncMacros"
Option Explicit

Sub OutSyncAddCal()
    Dim olApp, olfolder As Object        'Dim olApp As Outlook.Application
    'Dim olfolder As Outlook.MAPIFolder
    'Check to make sure a Staff exists on the Selected Line
    With Sheets("Staff")
        If .Range("C" & ActiveCell.Row).Value = Empty Then
            MsgBox "Please select a row containing a staff memeber before assigning a calendar to that staff"
            Exit Sub
        End If

        Set olApp = CreateObject("Outlook.Application")
        Set olfolder = olApp.GetNamespace("MAPI").PickFolder
        If olfolder Is Nothing Then Exit Sub
        '.Range("E" & ActiveCell.Row).Value = (olfolder.Name) 'This adds the Folder Name to anywhere you want
        .Range("E" & ActiveCell.Row).Value = (olfolder.EntryID)        'This adds the Folder ID
        Set olfolder = Nothing
        Set olApp = Nothing
    End With
End Sub

Sub SendToOultook()
    Dim olApp, olApt, olFldr, olObject, olItems, ExistItem, NS As Object        'As Outlook.Application
    Dim StaffRow, SelRow As Long
    Dim ApStart As Date
    Dim ApTime, ApDur, ApName, CalendID, SearchStart As String

    With Sheets("Schedule")
        If .Range("B8").Value = Empty Then
            MsgBox "Please select a correct Staff"
            Exit Sub
        End If

        StaffRow = .Range("B8").Value        'Staff Row
        CalendID = Sheet3.Range("E" & StaffRow).Value
        If CalendID = "" Then
            MsgBox "Please assign a Calendar to this staff"
            Sheet3.Activate
            Sheet3.Range("E" & StaffRow).Select
            OutSyncAddCal
            Exit Sub
        End If

        Set olApp = CreateObject("Outlook.Application")
        Set NS = olApp.GetNamespace("MAPI")
        SelRow = .Range("B1").Value        'Get Selected Row
        ApStart = .Range("M3").Value + .Range("L" & SelRow).Value        'Combine Date & Time
        ApName = .Range("M" & SelRow).Value
        ApDur = .Range("defApDuration").Value        'Set Duration


        Set olFldr = NS.GetFolderFromID(CalendID)
        Set olApt = olApp.CreateItem(olAppointmentItem)
        Set olItems = NS.GetFolderFromID(CalendID).Items
        Set ExistItem = NS.GetFolderFromID(CalendID).Items
        'ApStart = Format(ApStart, "ddddd hh:mm")
        SearchStart = "[Start]='" & Format(ApStart, "ddddd hh:mm") & "'"
        Set ExistItem = olItems.Find(SearchStart)
        If ExistItem Is Nothing Then
            With olApt
                .Subject = ApName
                .Start = ApStart
                .Duration = ApDur
                .ReminderSet = False
                .Categories = "Test Appointment"
                .Save
                .Move olFldr
                .Close olSave
            End With
        Else:
            With ExistItem
                .Subject = ApName
                .Duration = ApDur
                .ReminderSet = False
                .Categories = "Test Appointment"
                .Save
                .Close olSave
            End With
        End If
        Set olApp = Nothing
    End With
End Sub

