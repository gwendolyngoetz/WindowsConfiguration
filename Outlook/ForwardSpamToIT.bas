' Selected email will be forwarded as an attachment to the specified email account then moved to the Junk Email folder
' Only works on mail items and when a single message is selected.

Attribute VB_Name = "ForwardSpamToIT"
Option Explicit

Const SEND_TO As String = "EMAIL ADDRESS HERE"

Sub ForwardSpamToIT()
    Dim selectedEmail As Outlook.MailItem
    Dim msgboxResult As Integer
    
    msgboxResult = MsgBox("Send this to IT?", vbYesNo)
    
    If msgboxResult <> vbYes Then
        Exit Sub
    End If
    
    Set selectedEmail = GetSelectedEmail()
    
    If selectedEmail Is Nothing Then
        Exit Sub
    End If
        
    ForwardEmail selectedEmail
    MoveToSpamFolder selectedEmail
    
End Sub

Function GetSelectedEmail() As Outlook.MailItem
    Const PR_SENT_REPRESENTING_ENTRYID As String = "http://schemas.microsoft.com/mapi/proptag/0x00410102"
    
    Dim myOlExp As Outlook.Explorer
    Dim myOlSel As Outlook.Selection
    
    Set myOlExp = Application.ActiveExplorer
    Set myOlSel = myOlExp.Selection
    
    ' Only one selection at a time
    If myOlSel.Count <> 1 Then
        Set GetSelectedEmail = Nothing
        Exit Function
    End If
    
    ' Only mail types
    If myOlSel.item(1).Class <> OlObjectClass.olMail Then
        Set GetSelectedEmail = Nothing
        Exit Function
    End If
    
    Set GetSelectedEmail = myOlSel.item(1)
End Function

Sub ForwardEmail(mailToForward As Outlook.MailItem)
    Dim outgoingMail As MailItem
    
    On Error GoTo Release

    If mailToForward.Class <> OlObjectClass.olMail Then
        Exit Sub
    End If
    
    Set outgoingMail = Application.CreateItem(olMailItem)
    outgoingMail.Subject = "[Potential Spam] FW:" & mailToForward.Subject
    outgoingMail.HTMLBody = "Potential spam. Have a nice day." & vbCrLf
    outgoingMail.Attachments.Add mailToForward, olEmbeddeditem
    outgoingMail.To = SEND_TO
    outgoingMail.Save
    outgoingMail.Send
    
Release:
  Set outgoingMail = Nothing
End Sub

Sub MoveToSpamFolder(selectedMail As Outlook.MailItem)
    Dim myNameSpace As Outlook.NameSpace
    Dim myInbox As Outlook.Folder
    Dim destFolder As Outlook.Folder
 
    Set myNameSpace = Application.GetNamespace("MAPI")
    Set myInbox = myNameSpace.GetDefaultFolder(olFolderInbox)
    Set destFolder = myInbox.Parent.Folders("Junk Email")
    
    selectedMail.Move destFolder
End Sub
