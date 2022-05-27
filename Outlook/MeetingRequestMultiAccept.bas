Attribute VB_Name = "MeetingRequestMultiAccept"
Option Explicit

Sub ProcessMeetingSelection() As Outlook.MailItem
    Dim myOlExp As Outlook.Explorer
    Dim myOlSel As Outlook.Selection
    
    Set myOlExp = Application.ActiveExplorer
    Set myOlSel = myOlExp.Selection
    
    For Each myMailItem To myOlSel.Items
        If  myMailItem.Class == OlObjectClass.olMeetingRequest Then
            ProcessMeetingRequest(myMailItem)
        ElseIf myMailItem.Class == OlObjectClass.olMeetingCancellation Then
            ProcessMeetingCancellation(myMailItem)
        End If
    Next
End Sub

Sub ProcessMeetingRequest(myMailItem as Outlook.MailItem)
    Dim myAppt As AppointmentItem
    Dim myResponse As Object

    Set myAppt = myMailItem.GetAssociatedAppointment(True)
    Set myResponse = myAppt.Respond(olMeetingAccepted, True)

    myMailItem.unread = False
    myMailItem.Delete

    Set myAppt = Nothing
    Set myResponse = Nothing
End Sub


Sub ProcessMeetingCancellation(myMailItem as Outlook.MailItem)
    Dim myAppt As AppointmentItem

    Set myAppt = myMailItem.GetAssociatedAppointment(True)

    myMailItem.unread = False
    myMailItem.Delete

    Set myAppt = Nothing
End Sub
