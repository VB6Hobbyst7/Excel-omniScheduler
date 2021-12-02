Attribute VB_Name = "EmailSched_Macs"
Option Explicit

Function OutlookExists() As Boolean
    Dim xOLApp As Object
    '    On Error GoTo L1
    Set xOLApp = CreateObject("Outlook.Application")
    If Not xOLApp Is Nothing Then
        OutlookExists = True
        '        MsgBox "Outlook " & xOLApp.Version & " installed", vbExclamation
        Set xOLApp = Nothing
        Exit Function
    End If
    OutlookExists = False
    'L1: MsgBox "Outlook not installed", vbExclamation, "Kutools for Outlook"
End Function
'---------------------------------------------------------------------------------------
' Procedure : SendEmail
' Author    : CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Automate Outlook to send emails with or without attachments
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' strTo         To Recipient email address string (semi-colon separated list)
' strSubject    Text string to be used as the email subject line
' strBody       Text string to be used as the email body (actual message)
' bEdit         True/False whether or not you wish to preview the email before sending
' strBCC        BCC Recipient email address string (semi-colon separated list)
' AttachmentPath    single value or array of attachment (complete file paths with
'                   filename and extensions)
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2007-Nov-16             Initial Release
'---------------------------------------------------------------------------------------
Function SendEmail(strTo As String, strSubject As String, strBody As String, bEdit As Boolean, _
                   Optional strBCC As Variant, Optional AttachmentPath As Variant)
'Send Email using late binding to avoid reference issues
   Dim objOutlook As Object
   Dim objOutlookMsg As Object
   Dim objOutlookRecip As Object
   Dim objOutlookAttach As Object
   Dim I As Integer
   Const olMailItem = 0
 
   On Error GoTo ErrorMsgs
 
   Set objOutlook = CreateObject("Outlook.Application")
 
   Set objOutlookMsg = objOutlook.CreateItem(olMailItem)
   With objOutlookMsg
      Set objOutlookRecip = .Recipients.Add(strTo)
      objOutlookRecip.Type = 1
 
      If Not IsMissing(strBCC) Then
        Set objOutlookRecip = .Recipients.Add(strBCC)
        objOutlookRecip.Type = 3
      End If
 
      .Subject = strSubject
      .Body = strBody
      .Importance = 2  'Importance Level  0=Low,1=Normal,2=High
 
      ' Add attachments to the message.
      If Not IsMissing(AttachmentPath) Then
        If IsArray(AttachmentPath) Then
           For I = LBound(AttachmentPath) To UBound(AttachmentPath) '- 1
              If AttachmentPath(I) <> "" And AttachmentPath(I) <> "False" Then
                Set objOutlookAttach = .Attachments.Add(AttachmentPath(I))
              End If
           Next I
        Else
            If AttachmentPath <> "" Then
                Set objOutlookAttach = .Attachments.Add(AttachmentPath)
            End If
        End If
      End If
 
      For Each objOutlookRecip In .Recipients
         If Not objOutlookRecip.Resolve Then
            objOutlookMsg.Display
         End If
      Next
 
      If bEdit Then 'Choose btw transparent/silent send and preview send
        .Display
      Else
        .SEND
      End If
   End With
 
   Set objOutlookMsg = Nothing
   Set objOutlook = Nothing
   Set objOutlookRecip = Nothing
   Set objOutlookAttach = Nothing
 
ErrorMsgs:
   If Err.Number = "287" Then
      MsgBox "You clicked No to the Outlook security warning. " & _
      "Rerun the procedure and click Yes to access e-mail " & _
      "addresses to send your message. For more information, " & _
      "see the document at http://www.microsoft.com/office" & _
      "/previous/outlook/downloads/security.asp."
      Exit Function
   ElseIf Err.Number <> 0 Then
      MsgBox Err.Number & " - " & Err.Description
      Exit Function
   End If
End Function

