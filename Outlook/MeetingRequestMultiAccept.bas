'Attribute VB_Name = "MeetingRequestMultiAccept"
Option Explicit

Sub ProcessMeetingSelection()
    Dim myOlExp As Outlook.Explorer
    Dim myOlSel As Outlook.Selection
    Dim myMailItem As Object
    Dim i As Integer
    
    Set myOlExp = Application.ActiveExplorer
    Set myOlSel = myOlExp.Selection
    
    For i = 1 To myOlSel.Count
        Set myMailItem = myOlSel.Item(i)
        If myMailItem.Class = OlObjectClass.olMeetingRequest Then
            Call ProcessMeetingRequest(myMailItem)
        ElseIf myMailItem.Class = OlObjectClass.olMeetingCancellation Then
            Call ProcessMeetingCancellation(myMailItem)
        End If
    Next i
End Sub

Sub ProcessMeetingRequest(myMailItem As Object)
    Dim myAppt As AppointmentItem
    Dim myResponse As Object

    Set myAppt = myMailItem.GetAssociatedAppointment(True)
    Set myResponse = myAppt.Respond(olMeetingAccepted, True)

    myMailItem.UnRead = False
    myMailItem.Delete

    Set myAppt = Nothing
    Set myResponse = Nothing
End Sub


Sub ProcessMeetingCancellation(myMailItem As Object)
    Dim myAppt As AppointmentItem

    Set myAppt = myMailItem.GetAssociatedAppointment(True)
    myAppt.Delete
    
    myMailItem.UnRead = False
    myMailItem.Delete

    Set myAppt = Nothing
End Sub
