VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "Team Activity Organizer"
   ClientHeight    =   6525
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8970
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picProgress 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      Height          =   1515
      Left            =   0
      ScaleHeight     =   1455
      ScaleWidth      =   8910
      TabIndex        =   1
      Top             =   3675
      Width           =   8970
   End
   Begin VB.PictureBox picTAS 
      Align           =   2  'Align Bottom
      AutoSize        =   -1  'True
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1275
      ScaleWidth      =   8910
      TabIndex        =   0
      Top             =   5190
      Width           =   8970
   End
   Begin MSComctlLib.ImageList iList 
      Left            =   2400
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":0742
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":099E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":0BFA
            Key             =   ""
         EndProperty
      EndProperty
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
   Begin VB.Menu mnuConsole 
      Caption         =   "&Console   "
      Begin VB.Menu mnuConsoleSQLServer 
         Caption         =   "&SQL Server"
      End
      Begin VB.Menu mnuConsoleExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuAdministration 
      Caption         =   "&Administration   "
      Begin VB.Menu mnuAdministrationUsers 
         Caption         =   "&Users"
      End
      Begin VB.Menu mnuAdministrationProjects 
         Caption         =   "&Projects"
      End
      Begin VB.Menu mnuAdministrationActivityTypes 
         Caption         =   "&Activity Types"
      End
      Begin VB.Menu mnuAdministrationPauseCauses 
         Caption         =   "Pause &Causes"
      End
   End
   Begin VB.Menu mnuMessaging 
      Caption         =   "&Messaging   "
      Begin VB.Menu mnuMessagingSendMessage 
         Caption         =   "Send &Message"
      End
   End
   Begin VB.Menu mnuReportsAndStatistics 
      Caption         =   "&Reports And Statistics   "
      Begin VB.Menu mnuReportsAndStatisticsGenerateDailyActivityReports 
         Caption         =   "Generate &Daily Activity Reports"
      End
   End
   Begin VB.Menu mnuOrganizer 
      Caption         =   "&Organiser   "
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "&Windows    "
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub MDIForm_Load()
    Width = Screen.Width
    Height = Screen.Height - 600
    Me.WindowState = vbMaximized
    
    'the form must be fully visible before calling Shell_NotifyIcon
    Me.Visible = True
        
    SetMenuIcons
    
    With nid
     .cbSize = Len(nid)
     .hwnd = Me.hwnd ' since Top is 0, lvClients gets the events
     .uId = vbNull
     .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
     .uCallBackMessage = WM_MOUSEMOVE
     .hIcon = Me.Icon
     .szTip = "Team Activity Organizer" & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, nid
    
    picTAS.BackColor = RGB(200, 200, 200)
    picTAS.Picture = LoadPicture(App.Path & "\img\tas.bmp")
    
    picProgress.Height = 2340
    
    NewMSG "Loading Database Manager..", COLOR_NORMAL
    
    Load frmSqlServer

End Sub



Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If blnCloseCommand Then
        Unload Me
    Else
        Me.WindowState = vbMinimized
    End If
End Sub

'Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    If blnCloseCommand Then
'        Unload Me
'    Else
'        Cancel = 1
'    End If
'End Sub

Private Sub MDIForm_Resize()
    If Me.WindowState = vbMinimized Then
        'frmServer.WindowState = vbMinimized
        Me.Hide
    End If

End Sub


Private Sub mnuAdministrationActivityTypes_Click()
    Dim frm As frmProjects
    If Not winOpen.blnActivityTypes Then
        Set frm = New frmProjects
        frm.InitializeForm "ACTT"
    End If
End Sub

Private Sub mnuAdministrationPauseCauses_Click()
    Dim frm As frmProjects
    If Not winOpen.blnPauseCauses Then
        Set frm = New frmProjects
        frm.InitializeForm "PAUS"
    End If
End Sub

Private Sub mnuAdministrationProjects_Click()
    'Dim frm As frmProjects
    If Not winOpen.blnProjects Then
        Load frmProjects
        frmProjects.InitializeForm "PRJS"
    End If
End Sub

Private Sub mnuAdministrationUsers_Click()
    With winOpen
        If .blnActivityServer Then
            frmServer.WindowState = vbMinimized
        End If
        If .blnSQLServer Then
            frmSqlServer.WindowState = vbMinimized
        End If
        If .blnUsers Then
            frmUsers.SetFocus
        Else
            Load frmUsers
        End If
    End With
End Sub

'Private Sub MDIForm_Unload(Cancel As Integer)
'    If blnCloseCommand Then
'        Unload Me
'    Else
'        Cancel = 1
'    End If
'End Sub

Private Sub mnuConsoleExit_Click()
    blnCloseCommand = True
    Unload Me
End Sub






Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, x As _
   Single, y As Single)
'this procedure receives the callbacks from the System Tray icon.
Dim Result As Long
Dim msg As Long
 
 'the value of X will vary depending upon the scalemode setting
 'If lvClients.ScaleMode = vbPixels Then
 ' msg = X
 'Else
  
  If Me.WindowState = vbMaximized Then
    Exit Sub
  End If
  
 'If Me.ScaleMode = vbPixels Then
 '   msg = X
 'Else
    msg = x / Screen.TwipsPerPixelX
 'End If
 
 Select Case msg
  Case WM_LBUTTONUP        '514 restore form window
   Me.WindowState = vbMaximized
   Result = SetForegroundWindow(Me.hwnd)
   Me.Show
  Case WM_LBUTTONDBLCLK    '515 restore form window
   Me.WindowState = vbMaximized
   Result = SetForegroundWindow(Me.hwnd)
   Me.Show
  Case WM_RBUTTONUP        '517 display popup menu
   Result = SetForegroundWindow(Me.hwnd)
   Me.PopupMenu Me.mPopupSys
 End Select
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    ' this removes the icon from the system tray
    If blnCloseCommand Then
        Shell_NotifyIcon NIM_DELETE, nid
    Else
        Cancel = 1
    End If
End Sub


Private Sub mnuConsoleSQLServer_Click()
    If Not winOpen.blnSQLServer Then
        Load frmSqlServer
    End If
End Sub


Private Sub mnuMessagingSendMessage_Click()
    frmMessaging.Show
    If Not frmServer.lvClients.SelectedItem Is Nothing Then
        frmMessaging.SetClientIndex Mid(frmServer.lvClients.SelectedItem.Key, 2)
    End If
End Sub

Private Sub mnuReportsAndStatisticsGenerateDailyActivityReports_Click()
    'If winOpen.blnDailyReports Then
    '    frmDailyReport.SetFocus
    'Else
    '    frmDailyReport.Show
    'End If
End Sub

Private Sub mPopExit_Click()
 'called when user clicks the popup menu Exit command
 Unload Me
End Sub

Private Sub mPopRestore_Click()
 'called when the user clicks the popup menu Restore command
 Dim Result As Long
 Me.WindowState = vbMaximized
 Me.Show
 Result = SetForegroundWindow(Me.hwnd)
 ' mdiMain.Show
End Sub




