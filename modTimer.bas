Attribute VB_Name = "modTimer"
Option Explicit

' PROCEDURE FOR CAPTURING TIME EVENT CALLBACKS

Public Function TimerProc(ByVal wTimerID As Long, _
                          ByVal iMsg As Long, _
                          ByVal dwUser As Long, _
                          ByVal dw1 As Long, _
                          ByVal dw2 As Long) As Long
   
    Dim blnPaused       As Boolean
    Dim strTime         As String
    Dim dtmNow          As Date
    Dim lngDayDiff      As Long
    Dim lngHourDiff     As Long
    Dim lngMinDiff      As Long
    Dim lngSecDiff      As Long
    
    With arrClients(dwUser)
        If .IsConnected Then
            dtmNow = Now
            lngSecDiff = DateDiff("s", .LastActivityStartDate, dtmNow)
            lngDayDiff = Int(lngSecDiff / 86400)
            lngSecDiff = lngSecDiff - (lngDayDiff * 86400)
            lngHourDiff = Int(lngSecDiff / 3600)
            lngSecDiff = lngSecDiff - (lngHourDiff * 3600)
            lngMinDiff = Int(lngSecDiff / 60)
            lngSecDiff = lngSecDiff - (lngMinDiff * 60)
            
            strTime = Format(lngDayDiff, "00") & ":" & _
                      Format(lngHourDiff, "00") & ":" & _
                      Format(lngMinDiff, "00") & ":" & _
                      Format(lngSecDiff, "00")
                      
            frmServer.lvClients.ListItems("c" & dwUser).ListSubItems(4) = strTime
            'Call SetTimerEvent(dwUser)
        End If
    End With
           
   
End Function


Function SetTimerEvent(ByVal lngIndex As Long) As Long
    arrClients(lngIndex).TimerHandle = timeSetEvent(1000, 1, AddressOf TimerProc, ByVal lngIndex, TIME_ONESHOT)
End Function
