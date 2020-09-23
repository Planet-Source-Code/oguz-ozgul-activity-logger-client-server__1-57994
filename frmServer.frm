VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServer 
   Caption         =   "Server"
   ClientHeight    =   4080
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10830
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4080
   ScaleWidth      =   10830
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrClient 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   1000
      Left            =   4500
      Top             =   1620
   End
   Begin VB.Timer tmrPing 
      Enabled         =   0   'False
      Index           =   0
      Left            =   2520
      Top             =   1740
   End
   Begin MSWinsockLib.Winsock sckClient 
      Index           =   0
      Left            =   9120
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   9120
   End
   Begin MSComctlLib.ListView lvClients 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   873
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    winOpen.blnActivityServer = True
    Width = mdiMain.Width
    Height = mdiMain.Height - 600
    WindowState = vbMaximized
    ScaleMode = vbPixels
    With lvClients
        .View = lvwReport
        .Left = 0
        .Width = Width
        .Top = 0
        .Height = 450
        .AllowColumnReorder = True
        .Appearance = ccFlat
        .FlatScrollBar = True
        .FullRowSelect = True
        .HideColumnHeaders = False
        .HotTracking = True
        .LabelEdit = lvwManual
        .LabelWrap = False
        .MultiSelect = False
        .HoverSelection = True
    End With
    With lvClients.ColumnHeaders
        .Add 1, "clientName", "Name", (Width / Screen.TwipsPerPixelX) * (15 / 100)
        .Add 2, "clientProject", "Project", (Width / Screen.TwipsPerPixelX) * (16 / 100)
        .Add 3, "clientTaskType", "Task Type", (Width / Screen.TwipsPerPixelX) * (16 / 100)
        .Add 4, "clientTask", "Task", (Width / Screen.TwipsPerPixelX) * (44 / 100)
        .Add 5, "clientDuration", "Duration", (Width / Screen.TwipsPerPixelX) * (9 / 100)
    End With
    
    Call GetActivityTypes
    Call GetProjects
    Call GetPauseCauses
    
    sckClient(0).Listen
    
    NewMSG "Activity Server is listening on port 9120..", COLOR_NORMAL
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Dim lngIdx As Long
    
    On Error Resume Next
    
    If blnCloseCommand Then
    
        If lvClients.ListItems.Count > 0 Then
            NewMSG "There are clients connected to the server. Please disconnect them first", COLOR_WARNING
            Cancel = 1
            blnCloseCommand = False
            Exit Sub
        End If
        
        sckClient(0).Close
        
        For lngIdx = 1 To sckClient.UBound
            sckClient(lngIdx).Close
            Unload sckClient(lngIdx)
            'timeKillEvent arrClients(lngIdx).TimerHandle
            tmrClient(lngIdx).Enabled = False
            Unload tmrClient(lngIdx)
            tmrPing(lngIdx).Enabled = False
            Unload tmrPing(lngIdx)
        Next lngIdx
    
    Else
        Me.WindowState = vbMinimized
        Cancel = 1
    End If
    
End Sub



Private Sub sckClient_Close(Index As Integer)
    On Error Resume Next
    With arrClients(Index)
        'timeKillEvent .TimerHandle
        tmrClient(Index).Enabled = False
        If .IsLoggedIn Then
            If Not .ByeSent Then
                NewMSG .Name & " has disconnected because of a connection failure..", COLOR_EXCEPTION
                ' abnormal login
                If .LoginID > 0 And .UserID > 0 Then
                    Call objData.InsertLogOut(.UserID, .LoginID, False)
                End If
            Else
                NewMSG .Name & " has logged out..", COLOR_NORMAL
            End If
        Else
            NewMSG "The remote IP " & sckClient(Index).RemoteHostIP & " has disconnected", COLOR_NORMAL
        End If
        .IsConnected = False
    End With
    sckClient(Index).Close
    tmrClient(Index).Enabled = False
    Unload tmrClient(Index)
    tmrPing(Index).Enabled = False
    Unload tmrPing(Index)
    lvClients.ListItems.Remove "c" & Index
End Sub

Private Sub sckClient_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Dim lngIP           As Long
    Dim lngIPExists     As Long
    
    On Error Resume Next
    
    NewMSG "Connection Request received from " & sckClient(Index).RemoteHostIP, COLOR_NORMAL
    For lngIPExists = 1 To sckClient.UBound
        If sckClient(lngIPExists).RemoteHostIP = sckClient(0).RemoteHostIP Then
            If Err.Number = 0 Then
                ' If socket is alive and IP's conflict, do not accept request
                NewMSG sckClient(Index).RemoteHostIP & " is already connected. refusing connection", COLOR_EXCEPTION
                'Exit Sub
            Else
                ' Socket is unloaded. Try next
                Err.Clear
            End If
        End If
    Next lngIPExists
    Load tmrPing(sckClient.UBound + 1)
    With tmrPing(sckClient.UBound + 1)
        .Enabled = True
        .Interval = PING_INTERVAL
    End With
    Load tmrClient(sckClient.UBound + 1)
    With tmrClient(sckClient.UBound + 1)
        .Enabled = False
        .Interval = 1000
    End With
    Load sckClient(sckClient.UBound + 1)
    If sckClient.UBound = 1 Then
        ReDim ArrivedData(1)
    Else
        ReDim Preserve ArrivedData(sckClient.UBound)
    End If
    sckClient(sckClient.UBound).Accept requestID
    NewMSG "Connection accepted by socket " & sckClient.UBound, COLOR_NORMAL
    
    If Index = 1 Then
        ReDim arrClients(sckClient.UBound)
    Else
        ReDim Preserve arrClients(sckClient.UBound)
    End If
    
    With arrClients(sckClient.UBound)
        .ArrivedData = ""
        .ConnectedDate = Now
        .IsConnected = True
        .LastActivityStartDate = Now
        .Name = ""
    End With
        
    lvClients.ListItems.Add lvClients.ListItems.Count + 1, "c" & sckClient.UBound, "Connected.."
        'End If
    'Next lngIP
End Sub

Private Sub sckClient_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strData             As String
    Dim lngVBCRLFPos        As Long
    
    On Error GoTo Exc
    sckClient(Index).GetData strData
    With arrClients(Index)
        .ArrivedData = .ArrivedData & strData
        lngVBCRLFPos = InStrB(1, .ArrivedData, vbCrLf, vbBinaryCompare)
        If lngVBCRLFPos > 0 Then
            If lngVBCRLFPos < LenB(.ArrivedData) - 1 Then
                strData = LeftB(.ArrivedData, lngVBCRLFPos - 1)
                .ArrivedData = MidB(.ArrivedData, lngVBCRLFPos + 2)
            Else
                strData = LeftB(.ArrivedData, lngVBCRLFPos - 1)
                .ArrivedData = ""
            End If
            strData = objCrypto.Decrypt(strData)
            ProcessIncomingData Index, strData
        End If
    End With
    Exit Sub
Exc:
    ' Invalid data, possible CRC error.
    ' Omit invalid data, reset stored client data
    arrClients(Index).ArrivedData = ""
    ' log exception
    LogException "DataArrival", Err.Number, Err.Description, "Exception while receiving data from client " & arrClients(Index).Name, True
End Sub

Private Function ProcessIncomingData(ByVal Index As Integer, ByVal strData As String) As Long
    Dim arrDataInfo()       As String
    Dim lngCheckResult      As Long
    Dim strUserName         As String
    
    With arrClients(Index)
        Select Case Left(strData, 4)
            Case "LOGN"
                
                If .IsLoggedIn Then
                    ' Log exceptional data
                    LogException "ProcessIncomingData", 0, "Secondary login attempt", "The user " & .Name & " tried to login while he/she is alreaddy logged in", False
                    .ArrivedData = ""
                    Exit Function
                End If
                
                arrDataInfo = Split(Mid(strData, 5), "|")
                
                ' CheckPassword returns the user id if successful
                ' CheckPassword returns 0 if credentials ar invalid
                ' CheckPassword returns -1 if an exception occurs
                lngCheckResult = objData.CheckPassword(arrDataInfo(0), arrDataInfo(1))
                
                If lngCheckResult < 0 Then ' Exception occured
                    sckClient(Index).SendData objCrypto.Encrypt("LOGNEXCP|An exception occured while checking your credentials. Please try again") & vbCrLf
                    .ArrivedData = ""
                    Exit Function
                ElseIf lngCheckResult = 0 Then ' Login failed
                    NewMSG "Login failed for username " & arrDataInfo(0), COLOR_WARNING
                    sckClient(Index).SendData objCrypto.Encrypt("LOGNEXCP|Invalid user name and/or password") & vbCrLf
                    .ArrivedData = ""
                    ' Log Login
                    Call objData.InsertLogin(arrDataInfo(0), False, sckClient(Index).RemoteHostIP)
                    Exit Function
                Else
                    
                    ' Log Login
                    
                    .IsLoggedIn = True
                    
                    .LoginID = objData.InsertLogin(arrDataInfo(0), True, sckClient(Index).RemoteHostIP)
                    
                    .UserID = lngCheckResult
                    
                    strUserName = objData.GetUser(lngCheckResult)
                    
                    If strUserName <> "" Then
                        NewMSG "User logged in successfully, login ID: " & .LoginID & ", name: " & strUserName, COLOR_SUCCESS
                        sckClient(Index).SendData objCrypto.Encrypt("LOGNOK|" & strUserName) & vbCrLf
                    Else
                        .LoginID = 0
                        .UserID = 0
                        .IsConnected = False
                        sckClient(Index).SendData objCrypto.Encrypt("LOGNEXCP|An exception occured while checking your credentials. Please try again") & vbCrLf
                        .ArrivedData = ""
                        Exit Function
                    End If
                
                End If
                
                With lvClients.ListItems("c" & Index)
                    .Text = strUserName
                    .SubItems(1) = ""
                    .SubItems(2) = "No Activity"
                    .SubItems(3) = ""
                    .SubItems(4) = "00:00:00:00"
                End With
                
                .Name = strUserName
                .LastActivityStartDate = Now
                .IsOnPause = False
                .DoingNothing = True
            
                JoinEncryptAndSendToClient "PRJS", CLng(Index)
                JoinEncryptAndSendToClient "ACTT", CLng(Index)
                JoinEncryptAndSendToClient "PAUS", CLng(Index)
                
                '.TimerHandle = SetTimerEvent(CLng(Index))
                tmrClient(Index).Enabled = False
                tmrClient(Index).Enabled = True
            
            Case "STRT"
                
                If Not .IsLoggedIn Then
                    LogException "ProcessIncomingData", -1, "Activity Start Message while not logged in", "The remote IP " & sckClient(Index).RemoteHostIP & " tried to start an activity while not logged in", False
                    .ArrivedData = ""
                    Exit Function
                End If
                
                ' Check if an activity is going on
                If .ActivityID > 0 Then
                    Call objData.UpdateActivity(.ActivityID, ConvertTimeToSec(lvClients.ListItems("c" & Index).SubItems(4)))
                End If
                
                arrDataInfo = Split(Mid(strData, 5), "|")
                
                ' Insert activity details, get an activity ID
                .ActivityID = objData.InsertActivity(.UserID, arrDataInfo(0), arrDataInfo(1), arrDataInfo(2))
                
                If .ActivityID = 0 Then
                    LogException "ProcessIncomingData", -1, "Activity could not be inserted", "User " & .Name & " sent an activity start message, but the activity could not be logged", True
                    sckClient(Index).SendData objCrypto.Encrypt("STRTEXCP|Could not retrieve activity ID. Please try again") & vbCrLf
                    .ArrivedData = ""
                    Exit Function
                End If
                
                sckClient(Index).SendData objCrypto.Encrypt("STRTOK|" & .ActivityID) & vbCrLf
                
                With lvClients.ListItems("c" & Index)
                    .SubItems(1) = GetProjectName(arrDataInfo(0))
                    .SubItems(2) = GetActivityTypeName(arrDataInfo(1))
                    .SubItems(3) = arrDataInfo(2)
                    .SubItems(4) = "00:00:00:00"
                End With
                
                .LastActivityStartDate = Now
                .PreviousProject = arrDataInfo(0)
                .PreviousActivityType = arrDataInfo(1)
                .PreviousActivity = arrDataInfo(2)
                .PreviousActivityStartDate = Now
                .IsOnPause = False
                .DoingNothing = False
            
                ' Reset timer
                tmrClient(Index).Enabled = False
                tmrClient(Index).Enabled = True
            
            Case "PAUS"
                
                If Not .IsLoggedIn Then
                    LogException "ProcessIncomingData", -1, "Pause message while not logged in", "The remote IP " & sckClient(Index).RemoteHostIP & " sent a pause message while not logged in", False
                    .ArrivedData = ""
                    Exit Function
                End If
                
                If .IsOnPause Or .DoingNothing Then
                    LogException "ProcessIncomingData", -1, "Pause message while on puase or doing nothing", "The User: " & .Name & " sent a pause message while on puase or doing nothing", False
                    .ArrivedData = ""
                    Exit Function
                End If
                
                arrDataInfo = Split(Mid(strData, 5), "|")
                
                ' Insert pause information, get a Pause ID
                .PauseID = objData.InsertPause(.UserID, .ActivityID, arrDataInfo(0), arrDataInfo(1))
                
                If .PauseID = 0 Then
                    LogException "ProcessIncomingData", -1, "Could not insert pause start record", "The user " & .Name & " sent a pause message but an exception occured and the puase record could not be inserted", True
                    sckClient(Index).SendData objCrypto.Encrypt("PRESEXCP|Could not insert pause start record") & vbCrLf
                    .ArrivedData = ""
                    Exit Function
                End If
                
                sckClient(Index).SendData objCrypto.Encrypt("PRESOK|" & .PauseID) & vbCrLf
                
                .PauseStartTime = lvClients.ListItems("c" & Index).SubItems(4)
                
                With lvClients.ListItems("c" & Index)
                    .SubItems(1) = "Paused"
                    .SubItems(2) = GetPauseCauseName(arrDataInfo(0))
                    .SubItems(3) = arrDataInfo(1)
                    .SubItems(4) = "00:00:00:00"
                End With
                
                .DoingNothing = False
                .IsOnPause = True
                .LastActivityStartDate = Now
                
                tmrClient(Index).Enabled = False
                tmrClient(Index).Enabled = True
            
            Case "CONT"
                
                If Not .IsLoggedIn Then
                    LogException "ProcessIncomingData", -1, "Continue message while not logged in", "The remote IP " & sckClient(Index).RemoteHostIP & " sent a continue message while not logged in", False
                    .ArrivedData = ""
                    Exit Function
                End If
                
                If Not .IsOnPause Then
                    LogException "ProcessIncomingData", -1, "Continue message while not in pause", "The user " & .Name & " sent a continue message while not in pause", False
                    .ArrivedData = ""
                    Exit Function
                End If
                
                ' Update the pause information. Set total pause time
                Call objData.UpdatePause(.PauseID, ConvertTimeToSec(lvClients.ListItems("c" & Index).SubItems(4)))
                
                .PauseID = 0
                
                sckClient(Index).SendData objCrypto.Encrypt("CONTOK") & vbCrLf
                
                With lvClients.ListItems("c" & Index)
                    .SubItems(1) = GetProjectName(arrClients(Index).PreviousProject)
                    .SubItems(2) = GetActivityTypeName(arrClients(Index).PreviousActivityType)
                    .SubItems(3) = arrClients(Index).PreviousActivity
                    .SubItems(4) = arrClients(Index).PauseStartTime
                End With
                
                .LastActivityStartDate = Now
                .LastActivityStartDate = DateAdd("s", -CLng(Right(.PauseStartTime, 2)), Now)
                .LastActivityStartDate = DateAdd("n", -CLng(Mid(.PauseStartTime, 7, 2)), .LastActivityStartDate)
                .LastActivityStartDate = DateAdd("h", -CLng(Mid(.PauseStartTime, 4, 2)), .LastActivityStartDate)
                .LastActivityStartDate = DateAdd("d", -CLng(Left(.PauseStartTime, 2)), .LastActivityStartDate)
                .IsOnPause = False
                .DoingNothing = False
                
                tmrClient(Index).Enabled = False
                tmrClient(Index).Enabled = True
                
            Case "STOP"
                
                If Not .IsLoggedIn Then
                    LogException "ProcessIncomingData", -1, "Stop message while not logged in", "The remote IP " & sckClient(Index).RemoteHostIP & " sent a stop message while not logged in", False
                    .ArrivedData = ""
                    Exit Function
                End If
                If .DoingNothing Then
                    LogException "ProcessIncomingData", -1, "Stop message while doing nothing", "The user " & .Name & " sent a stop message while doing nothing", False
                    .ArrivedData = ""
                    Exit Function
                End If
                
                ' First, update activity total time
                
                Call objData.UpdateActivity(.ActivityID, ConvertTimeToSec(lvClients.ListItems("c" & Index).SubItems(4)))
                
                .ActivityID = 0
                
                sckClient(Index).SendData objCrypto.Encrypt("STOPOK") & vbCrLf
                
                With lvClients.ListItems("c" & Index)
                    .SubItems(1) = ""
                    .SubItems(2) = "No Activity"
                    .SubItems(3) = ""
                    .SubItems(4) = "00:00:00:00"
                End With
                
                .PreviousActivity = ""
                .PreviousActivityStartDate = Now
                .PreviousActivityType = ""
                .PreviousProject = ""
                .LastActivityStartDate = Now
                .IsOnPause = False
                .DoingNothing = True
                
                tmrClient(Index).Enabled = False
                tmrClient(Index).Enabled = True
                
            Case "RPRT"
            
                If .IsLoggedIn Then
                    
                    arrDataInfo = Split(Mid(strData, 5), "|")
                    ' 1 date expected
                    If IsDate(arrDataInfo(0)) Then
                        NewMSG "Report request received from client: " & .Name & ", Date: " & arrDataInfo(0), COLOR_NORMAL
                        strData = objData.GetActivityReports(arrDataInfo(0), arrDataInfo(0), .UserID)
                        ' Send the report data back to client
                        sckClient(Index).SendData objCrypto.Encrypt("RPRT" & strData) & vbCrLf
                    End If
                End If
                
            Case "PING"
                
                NewMSG "Connection with remote IP " & sckClient(Index).RemoteHostIP & " is alive", COLOR_SUCCESS
                .PingSent = False
            
            Case "BYE!"
                
                If .IsLoggedIn And (Not .DoingNothing) Then
                    LogException "ProcessIncomingData", -1, "BYE message while in activity or pause", "User " & .Name & " sent BYE message while in activity or pause", False
                    .ArrivedData = ""
                    Exit Function
                End If
                
                If .IsLoggedIn Then
                    Call objData.InsertLogOut(.UserID, .LoginID, True)
                End If
                
                .ByeSent = True
                sckClient(Index).Close
                Call sckClient_Close(Index)
            
            Case Else
                ' Do nothing. Logging here can cause DOS (Denial of service)
            
        End Select
        
        .ArrivedData = ""
    
    End With

End Function

Function JoinEncryptAndSendToClient(ByVal strPrefix As String, ByVal lngIndex As Long) As Long
    Dim strData     As String
    Dim lngIdx      As Long
    On Error Resume Next
    If strPrefix = "PRJS" Then
        strData = "PRJS"
        For lngIdx = 0 To UBound(arrProjects)
            If Err.Number <> 0 Then Exit Function
            With arrProjects(lngIdx)
                If .Visible Then
                    strData = strData & .ProjectID & "|" & .ProjectName & "|"
                End If
            End With
        Next lngIdx
    ElseIf strPrefix = "ACTT" Then
        strData = "ACTT"
        For lngIdx = 0 To UBound(arrActivityTypes)
            If Err.Number <> 0 Then Exit Function
            With arrActivityTypes(lngIdx)
                If .Visible Then
                    strData = strData & .ActivityTypeID & "|" & .ActivityTypeName & "|"
                End If
            End With
        Next lngIdx
    ElseIf strPrefix = "PAUS" Then
        strData = "PAUS"
        For lngIdx = 0 To UBound(arrPauseCauses)
            If Err.Number <> 0 Then Exit Function
            With arrPauseCauses(lngIdx)
                strData = strData & .PauseCauseID & "|" & .strPauseCause & "|"
            End With
        Next lngIdx
    End If
    
    If Right(strData, 1) = "|" Then
        strData = Left(strData, Len(strData) - 1)
    End If
    
    sckClient(lngIndex).SendData objCrypto.Encrypt(strData) & vbCrLf

End Function

Private Sub sckClient_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    ' Log Exception
    With arrClients(Index)
        Call LogException("sclClient_Error", Number, Description, "The user " & .Name & " has been disconnected from the server because of communications failure", True)
        If .UserID > 0 And .LoginID > 0 Then
            Call objData.InsertLogOut(arrClients(Index).UserID, arrClients(Index).LoginID, False)
        End If
    End With
    sckClient(Index).Close
    Call sckClient_Close(Index)
    
End Sub

Private Sub tmrClient_Timer(Index As Integer)
    TimerProc 0, 0, Index, 0, 0
End Sub

Private Sub tmrPing_Timer(Index As Integer)

    'Exit Sub
    With arrClients(Index)
        If .IsConnected Then
            If .PingSent Then
                sckClient(Index).Close
                Call sckClient_Close(Index)
                NewMSG "The connection with user " & .Name & " has timed out. The client has been disconnected", COLOR_EXCEPTION
            Else
                sckClient(Index).SendData objCrypto.Encrypt("PING") & vbCrLf
                .PingSent = True
            End If
        Else
            tmrPing(Index).Enabled = False
        End If
    End With

End Sub

Private Function ConvertTimeToSec(ByVal strDDHHMMSS As String) As Long
    
    Dim lngDays         As Long
    Dim lngHours        As Long
    Dim lngMinutes      As Long
    Dim lngSeconds      As Long
    
    lngDays = Val(Left(strDDHHMMSS, 2))
    lngHours = Val(Mid(strDDHHMMSS, 4, 2))
    lngMinutes = Val(Mid(strDDHHMMSS, 7, 2))
    lngSeconds = Val(Right(strDDHHMMSS, 2))
    
    ConvertTimeToSec = lngDays * 86400 + _
                       lngHours * 3600 + _
                       lngMinutes * 60 + _
                       lngSeconds

End Function


Function SendToAll(ByVal strText As String) As Long
    On Error GoTo excHandler
    Dim lngClient       As Long
    Dim strEncText      As String
    strEncText = objCrypto.Encrypt("AMSG|" & strText)
    For lngClient = 0 To UBound(arrClients)
        With arrClients(lngClient)
            If .IsLoggedIn Then
                sckClient(lngClient).SendData strEncText & vbCrLf
            End If
        End With
    Next lngClient
    Exit Function
excHandler:
    LogException "SendToAll", Err.Number, Err.Description, "Exception while sending data to clients", True
    Resume Next
End Function


