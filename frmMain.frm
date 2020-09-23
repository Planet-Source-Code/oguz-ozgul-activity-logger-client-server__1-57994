VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "00:00:00:00"
   ClientHeight    =   495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2775
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   33
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUpload 
      DisabledPicture =   "frmMain.frx":0742
      Height          =   495
      Left            =   2280
      Picture         =   "frmMain.frx":1468
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Activity Report And I-Loop Upload"
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton cmdWhatAmIDoing 
      DisabledPicture =   "frmMain.frx":218E
      Height          =   495
      Left            =   1785
      Picture         =   "frmMain.frx":2EB4
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Information about the current activity"
      Top             =   0
      Width           =   495
   End
   Begin VB.PictureBox picLedOn 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   30
      Picture         =   "frmMain.frx":3BDA
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox picLedOff 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   30
      Picture         =   "frmMain.frx":3EEC
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   3
      Top             =   120
      Width           =   225
   End
   Begin VB.CommandButton cmdStop 
      DisabledPicture =   "frmMain.frx":41FE
      Height          =   495
      Left            =   1290
      Picture         =   "frmMain.frx":4F24
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Ends the current activity"
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton cmdPause 
      DisabledPicture =   "frmMain.frx":5C4A
      Height          =   495
      Left            =   795
      Picture         =   "frmMain.frx":6970
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Pauses the current activity"
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton cmdStart 
      DisabledPicture =   "frmMain.frx":7696
      Height          =   495
      Left            =   300
      Picture         =   "frmMain.frx":83BC
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Starts a new activity, or continues a paused activity"
      Top             =   0
      Width           =   495
   End
   Begin VB.Timer tmrActivity 
      Interval        =   950
      Left            =   1380
      Top             =   -60
   End
   Begin VB.Timer tmrConnection 
      Interval        =   5000
      Left            =   900
      Top             =   -60
   End
   Begin MSWinsockLib.Winsock sckServer 
      Left            =   420
      Top             =   -60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mPopupSys 
      Caption         =   "&SysTray"
      Visible         =   0   'False
      Begin VB.Menu mPopRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mPopExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private dtmActivityStarted      As Date
Private blnPaused               As Boolean
Private blnDoingNothing         As Boolean
Private strPausedActivity       As String
Private strCurrentActivity      As String
Private strPausedProject        As String
Private strPausedActivityType   As String
Private strPausedTime           As String
Private arrivedData             As String
Private blnInputBoxDisplayed    As Boolean
Private strUserFullName         As String


Private Enum appStates
    NOT_LOGGED_IN = 0
    DOING_NOTHING = 1
    IN_ACTIVITY = 2
    IN_PAUSE = 3
End Enum

Private Type btnEnableState
    Login_Enabled       As Boolean
    Disconnect_Enabled  As Boolean
    Start_Enabled       As Boolean
    Stop_Enabled        As Boolean
    Finish_Enabled      As Boolean
End Type

Private lngAppState     As appStates

Private enableStates    As btnEnableState

Private strServerIP     As String
Private lngServerPort   As Long






Private Sub cmdStart_Click()
    
    Dim strData As String
    
    If enableStates.Start_Enabled Then
        If lngAppState = DOING_NOTHING Then
            strData = "STRT"
            
            blnInputBoxDisplayed = True
            strCurrentActivity = InputBox("Please enter your activity detail", "Activity Client", "")
            If Len(Trim(strCurrentActivity)) = 0 Then
                MsgBox "Please enter your activity detail"
                Exit Sub
            End If
            strData = strData & "1|1|" & strCurrentActivity
            blnInputBoxDisplayed = False
            
            sckServer.SendData objCrypto.Encrypt(strData) & vbCrLf
        ElseIf lngAppState = IN_PAUSE Then
            sckServer.SendData objCrypto.Encrypt("CONT") & vbCrLf
        End If
        
    End If

End Sub

Private Sub cmdStop_Click()
    If enableStates.Finish_Enabled Then
        sckServer.SendData objCrypto.Encrypt("STOP") & vbCrLf
    End If
End Sub

Private Sub cmdUpload_Click()
    blnInputBoxDisplayed = True
    frmReport.Show ' vbModal, Me
End Sub

Private Sub cmdWhatAmIDoing_Click()
    If lngAppState = DOING_NOTHING Then
        MsgBox "No Activity", vbOKOnly, "Activity Client"
    ElseIf lngAppState = NOT_LOGGED_IN Then
        MsgBox "You are not logged in. Make sure The Activity Server is up and running and your domain name is configured as a user by your team lead/supervisor"
    ElseIf lngAppState = IN_ACTIVITY Then
        MsgBox "Your current Activity: " & strCurrentActivity
    ElseIf lngAppState = IN_PAUSE Then
        MsgBox "Your " & vbCrLf & vbCrLf & strPausedActivity & vbCrLf & vbCrLf & "Activity is interrupted. The cause is:" & vbCrLf & vbCrLf & strCurrentActivity, vbOKOnly, "Activity Client"
    End If
End Sub

Private Sub Form_Load()
    
    Dim ff          As Long
    Dim strInput    As String
    
    ChangeAppState NOT_LOGGED_IN
    
    dtmActivityStarted = Now
    
    Caption = "00:00:00:00"
    
    ff = FreeFile
    
    Open App.Path & "\client.cfg" For Input Access Read As ff
    
    Line Input #ff, strInput
    strServerIP = Mid(strInput, Len("ServerIP=") + 1)
    
    Line Input #ff, strInput
    lngServerPort = CLng(Mid(strInput, Len("ServerPort=") + 1))
    
    sckServer.Connect strServerIP, lngServerPort
    
    tmrConnection.Enabled = True
    
    'the form must be fully visible before calling Shell_NotifyIcon
    
    Me.Refresh
    
    With nid
     .cbSize = Len(nid)
     .hwnd = Me.hwnd ' since Top is 0, lvClients gets the events
     .uId = vbNull
     .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
     .uCallBackMessage = WM_MOUSEMOVE
     .hIcon = Me.Icon
     .szTip = "Activity Organizer Client" & vbNullChar
    End With
    
    Shell_NotifyIcon NIM_ADD, nid

    SetTopMostWindow Me.hwnd, True
    
    Dim s As Size
    
    s = GetStartBarSize()
    
    Me.Left = Screen.Width - Me.Width
    Me.Top = Screen.Height - Me.Height - (s.Height - 20) * Screen.TwipsPerPixelY

End Sub


Private Function startActivity(ByVal strActivity As String) As Long
    
    Dim ff          As Long
    Dim strTime     As String
    
    If strActivity = "CONTINUE" Then
        'strActivity = txtActivity.Text
        strTime = strPausedTime
        dtmActivityStarted = DateAdd("s", -CLng(Right(strPausedTime, 2)), Now)
        dtmActivityStarted = DateAdd("n", -CLng(Mid(strPausedTime, 7, 2)), dtmActivityStarted)
        dtmActivityStarted = DateAdd("h", -CLng(Mid(strPausedTime, 4, 2)), dtmActivityStarted)
        dtmActivityStarted = DateAdd("d", -CLng(Left(strPausedTime, 2)), dtmActivityStarted)
    ElseIf strActivity <> "PAUSE" And strActivity <> "STOP" Then
        strTime = "00:00:00:00"
        dtmActivityStarted = Now
    End If
    
    If strActivity <> "PAUSE" And strActivity <> "STOP" Then
        If Len(strActivity) <= 64 Then
            frmMain.Caption = strTime
        Else
            frmMain.Caption = strTime
        End If
        blnPaused = False
    End If
    
End Function


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'    Me.WindowState = vbMinimized
'    Cancel = 1
'    Exit Sub
    
    If lngAppState = DOING_NOTHING Then
        sckServer.SendData ("BYE!")
    ElseIf lngAppState <> NOT_LOGGED_IN Then
        Cancel = 1
        'lblAlert.Caption = "Please finish your activity first"
    End If

End Sub


Private Sub cmdPause_Click()
    
    Dim strData             As String
    Dim strTempActivity     As String
    
    If enableStates.Stop_Enabled Then
        
        strData = "PAUS"
        
        blnInputBoxDisplayed = True
        strTempActivity = InputBox("Please enter the interruption reason", "Activity Client")
        If Len(Trim(strTempActivity)) = 0 Then
            MsgBox "Please enter the interruption reason", vbOKOnly, "Activity Client"
            Exit Sub
        End If
        strPausedActivity = strCurrentActivity
        strCurrentActivity = strTempActivity
        strData = strData & "1|" & strCurrentActivity
        blnInputBoxDisplayed = False
        
        sckServer.SendData objCrypto.Encrypt(strData) & vbCrLf
    
    End If

End Sub

Private Sub sckServer_Close()
    ChangeAppState NOT_LOGGED_IN
    sckServer.Close
    sckServer.Connect strServerIP, lngServerPort
    tmrConnection.Enabled = False
    tmrConnection.Enabled = True
End Sub

Private Sub sckServer_Connect()
    tmrConnection.Enabled = False
    picLedOff.Visible = False
    picLedOn.Visible = True
    sckServer.SendData objCrypto.Encrypt("LOGN" & GetLoggedInUserName & "|NAN") & vbCrLf
End Sub

Private Sub sckServer_DataArrival(ByVal bytesTotal As Long)
    Dim strData             As String
    Dim lngVBCRLFPos        As Long
    
    On Error GoTo Exc
    sckServer.GetData strData
    arrivedData = arrivedData & strData
    
    Do
        
        
        If lngVBCRLFPos > 0 Then
            If lngVBCRLFPos < LenB(arrivedData) - 1 Then
                strData = LeftB(arrivedData, lngVBCRLFPos - 1)
                arrivedData = MidB(arrivedData, lngVBCRLFPos + 4)
            Else
                strData = LeftB(arrivedData, lngVBCRLFPos - 1)
                arrivedData = ""
            End If
            strData = objCrypto.Decrypt(strData)
            ProcessIncomingData strData
        End If
        
        lngVBCRLFPos = InStrB(1, arrivedData, vbCrLf, vbBinaryCompare)
    
    Loop Until lngVBCRLFPos <= 0
    
    Exit Sub
Exc:
    ' Invalid data, possible CRC error.
    ' Omit invalid data, reset stored client data
    arrivedData = ""
    ' log exception
End Sub

Private Function ProcessIncomingData(ByVal strData As String) As Long

    Dim arrDataInfo()       As String

    Select Case Left(strData, 4)
        
        Case "PING"
            sckServer.SendData objCrypto.Encrypt("PING") & vbCrLf
        Case "LOGN"
            arrDataInfo = Split(Mid(strData, 5), "|")
            If arrDataInfo(0) = "EXCP" Then
                'lblAlert.Caption = arrDataInfo(1)
            ElseIf arrDataInfo(0) = "OK" Then
                strUserFullName = arrDataInfo(1)
                'lblAlert.Caption = "Logged in successfully"
                ChangeAppState DOING_NOTHING
            End If
        Case "PRJS"
            CreateCombo "PRJS", Mid(strData, 5)
        Case "ACTT"
            CreateCombo "ACTT", Mid(strData, 5)
        Case "PAUS"
            CreateCombo "PAUS", Mid(strData, 5)
        Case "STRT"
            arrDataInfo = Split(Mid(strData, 5), "|")
            If arrDataInfo(0) = "OK" Then
                startActivity "" 'txtActivity.Text
                'lblAlert.Caption = "Activity Started"
                ChangeAppState IN_ACTIVITY
            Else
                'lblAlert.Caption = arrDataInfo(1)
            End If
        
        Case "PRES"
            arrDataInfo = Split(Mid(strData, 5), "|")
            If arrDataInfo(0) = "OK" Then
                strPausedTime = Right(Caption, 11)
                startActivity "PAUSE"
                dtmActivityStarted = Now
                Caption = "00:00:00:00"
                'lblAlert.Caption = "Activity Paused.."
                ChangeAppState IN_PAUSE
            Else
                'lblAlert.Caption = arrDataInfo(1) ' Exception msg
            End If
        Case "CONT"
            arrDataInfo = Split(Mid(strData, 5), "|")
            If arrDataInfo(0) = "OK" Then
                'lblAlert.Caption = "Continuing Activity.."
                ChangeAppState IN_ACTIVITY
                startActivity "CONTINUE"
                strCurrentActivity = strPausedActivity
            Else
                'lblAlert.Caption = arrDataInfo(1) ' Exception msg
            End If
        Case "STOP"
            arrDataInfo = Split(Mid(strData, 5), "|")
            If arrDataInfo(0) = "OK" Then
                startActivity "STOP"
                Caption = "00:00:00:00"
                dtmActivityStarted = Now
                'lblAlert.Caption = "No Activity.."
                ChangeAppState DOING_NOTHING
            Else
                'lblAlert.Caption = arrDataInfo(1) ' Exception msg
            End If
        Case "AMSG"
            'lblAlert.Caption = Mid(strData, 5)
        Case "RPRT"
            Dim arrActivityInfo()   As String
            Dim lngActivity         As Long
            
            arrDataInfo = Split(Mid(strData, 5), Chr(1))
            
            With frmReport.flexReport
                .Rows = 2
                .TextMatrix(1, 0) = ""
                .TextMatrix(1, 1) = ""
                .TextMatrix(1, 2) = ""
            
                For lngActivity = 0 To UBound(arrDataInfo) - 1
                    If lngActivity > 0 Then
                        .Rows = .Rows + 1
                    End If
                    arrActivityInfo = Split(arrDataInfo(lngActivity), Chr(2))
                    .TextMatrix(.Rows - 1, 0) = arrActivityInfo(1)
                    .TextMatrix(.Rows - 1, 1) = arrActivityInfo(0)
                    .TextMatrix(.Rows - 1, 2) = FormatNumber(arrActivityInfo(2) / 60, 2)
                    
                Next ' lngActivity
                
            End With
            
        
        
        
    End Select

    'arrivedData = ""

End Function

Private Function CreateCombo(ByVal strPrefix As String, strData As String) As Long
    
    Dim arrData()       As String
    Dim lngIdx          As Long
    
    Select Case strPrefix
        Case "PRJS"
            arrData = Split(strData, "|")
            For lngIdx = 0 To UBound(arrData) - 1 Step 2
                'cmbProjects.AddItem arrData(lngIdx + 1)
                'cmbProjectIDs.AddItem arrData(lngIdx)
            Next lngIdx
        Case "ACTT"
            arrData = Split(strData, "|")
            For lngIdx = 0 To UBound(arrData) - 1 Step 2
                'cmbActivityTypes.AddItem arrData(lngIdx + 1)
                'cmbActivityTypeIDs.AddItem arrData(lngIdx)
            Next lngIdx
        Case "PAUS"
            arrData = Split(strData, "|")
            For lngIdx = 0 To UBound(arrData) - 1 Step 2
                'cmbPauseCauses.AddItem arrData(lngIdx + 1)
                'cmbPauseCauseIDs.AddItem arrData(lngIdx)
            Next lngIdx
    End Select
End Function

Private Sub tmrActivity_Timer()
    
    Dim strTime         As String
    Dim dtmNow          As Date
    Dim lngDayDiff      As Long
    Dim lngHourDiff     As Long
    Dim lngMinDiff      As Long
    Dim lngSecDiff      As Long
    
    If Not blnInputBoxDisplayed Then
        SetTopMostWindow Me.hwnd, True
    End If
    
    'If lngAppState =  Then
        dtmNow = Now
        lngSecDiff = DateDiff("s", dtmActivityStarted, dtmNow)
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
                  
        Caption = Left(Caption, Len(Caption) - 11) & strTime
    'End If

End Sub

Private Sub tmrConnection_Timer()
    sckServer.Close
    'lblAlert.Caption = "Could not connect to server. Trying again.."
    sckServer.Connect strServerIP, lngServerPort
End Sub



Private Sub form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'this procedure receives the callbacks from the System Tray icon.
    Dim Result As Long
    Dim msg As Long
 
    If Me.WindowState = vbNormal Then Exit Sub
 
    'the value of X will vary depending upon the scalemode setting
    If ScaleMode = vbPixels Then
        msg = X
    Else
        msg = X / Screen.TwipsPerPixelX
    End If
    
    Select Case msg
        Case WM_LBUTTONUP        '514 restore form window
            Me.WindowState = vbNormal
            Result = SetForegroundWindow(Me.hwnd)
            Me.Show
        Case WM_LBUTTONDBLCLK    '515 restore form window
            Me.WindowState = vbNormal
            Result = SetForegroundWindow(Me.hwnd)
            Me.Show
        Case WM_RBUTTONUP        '517 display popup menu
            Result = SetForegroundWindow(Me.hwnd)
            Me.PopupMenu Me.mPopupSys
    End Select

End Sub

Private Sub Form_Resize()
    'this is necessary to assure that the minimized window is hidden
    If Me.WindowState = vbMinimized Then Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'this removes the icon from the system tray
    Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub mPopExit_Click()
    'called when user clicks the popup menu Exit command
    Unload Me
End Sub

Private Sub mPopRestore_Click()
    'called when the user clicks the popup menu Restore command
    Dim Result As Long
    Me.WindowState = vbNormal
    Result = SetForegroundWindow(Me.hwnd)
    Me.Show
End Sub


Private Function ChangeAppState(ByVal lngState As appStates) As Long
    Select Case lngState
        Case NOT_LOGGED_IN
            'picLogin.Picture = LoadPicture(App.Path & "\img\cBtnLgn.bmp")
            
            cmdStart.Enabled = False
            cmdPause.Enabled = False
            cmdStop.Enabled = False
            'picStart.Picture = LoadPicture(App.Path & "\img\cBtnStrt_d.bmp")
            'picPause.Picture = LoadPicture(App.Path & "\img\cBtnStop_d.bmp")
            'picFinish.Picture = LoadPicture(App.Path & "\img\cBtnFin_d.bmp")
            lngAppState = NOT_LOGGED_IN
            With enableStates
                .Finish_Enabled = False
                .Login_Enabled = True
                .Disconnect_Enabled = False
                .Start_Enabled = False
                .Stop_Enabled = False
            End With
            'txtUserName.Enabled = True
            'txtPassword.Enabled = True
            'cmbProjects.Enabled = False
            'cmbActivityTypes.Enabled = False
            'cmbPauseCauses.Enabled = False
            'txtActivity.Enabled = False
            
        Case DOING_NOTHING
            'picLogin.Picture = LoadPicture(App.Path & "\img\cBtnDscn.bmp")
            cmdStart.Enabled = True
            cmdPause.Enabled = False
            cmdStop.Enabled = False
            'picStart.Picture = LoadPicture(App.Path & "\img\cBtnStrt.bmp")
            'picPause.Picture = LoadPicture(App.Path & "\img\cBtnStop_d.bmp")
            'picFinish.Picture = LoadPicture(App.Path & "\img\cBtnFin_d.bmp")
            lngAppState = DOING_NOTHING
            With enableStates
                .Finish_Enabled = False
                .Login_Enabled = False
                .Disconnect_Enabled = True
                .Start_Enabled = True
                .Stop_Enabled = False
            End With
            'txtUserName.Enabled = False
            'txtPassword.Enabled = False
            'cmbProjects.Enabled = True
            'cmbActivityTypes.Enabled = True
            'cmbPauseCauses.Enabled = False
            'txtActivity.Enabled = True
        
        Case IN_ACTIVITY
            'picLogin.Picture = LoadPicture(App.Path & "\img\cBtnLgn_d.bmp")
            cmdStart.Enabled = False
            cmdPause.Enabled = True
            cmdStop.Enabled = True
            'picStart.Picture = LoadPicture(App.Path & "\img\cBtnStrt_d.bmp")
            'picPause.Picture = LoadPicture(App.Path & "\img\cBtnStop.bmp")
            'picFinish.Picture = LoadPicture(App.Path & "\img\cBtnFin.bmp")
            lngAppState = IN_ACTIVITY
            With enableStates
                .Finish_Enabled = True
                .Login_Enabled = False
                .Disconnect_Enabled = False
                .Start_Enabled = False
                .Stop_Enabled = True
            End With
            'txtUserName.Enabled = False
            'txtPassword.Enabled = False
            'cmbProjects.Enabled = False
            'cmbActivityTypes.Enabled = False
            'cmbPauseCauses.Enabled = True
            'txtActivity.Enabled = True
        
        Case IN_PAUSE
            'picLogin.Picture = LoadPicture(App.Path & "\img\cBtnLgn_d.bmp")
            cmdStart.Enabled = True
            cmdPause.Enabled = False
            cmdStop.Enabled = False
            'picStart.Picture = LoadPicture(App.Path & "\img\cBtnStrt.bmp")
            'picPause.Picture = LoadPicture(App.Path & "\img\cBtnStop_d.bmp")
            'picFinish.Picture = LoadPicture(App.Path & "\img\cBtnFin_d.bmp")
            lngAppState = IN_PAUSE
            With enableStates
                .Finish_Enabled = False
                .Login_Enabled = False
                .Disconnect_Enabled = False
                .Start_Enabled = True
                .Stop_Enabled = False
            End With
            'txtUserName.Enabled = False
            'txtPassword.Enabled = False
            'cmbProjects.Enabled = False
            'cmbActivityTypes.Enabled = False
            'cmbPauseCauses.Enabled = False
            'txtActivity.Enabled = False
    
    End Select

    With enableStates
        If .Login_Enabled Or .Disconnect_Enabled Then
            'picLogin.MousePointer = MousePointerConstants.vbCrosshair
        Else
            'picLogin.MousePointer = MousePointerConstants.vbArrow
        End If
        If .Start_Enabled Then
            'picStart.MousePointer = MousePointerConstants.vbCrosshair
        Else
            'picStart.MousePointer = MousePointerConstants.vbArrow
        End If
        If .Stop_Enabled Then
            'picPause.MousePointer = MousePointerConstants.vbCrosshair
        Else
            'picPause.MousePointer = MousePointerConstants.vbArrow
        End If
        If .Finish_Enabled Then
            'picFinish.MousePointer = MousePointerConstants.vbCrosshair
        Else
            'picFinish.MousePointer = MousePointerConstants.vbArrow
        End If
    End With
End Function
