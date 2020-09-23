VERSION 5.00
Begin VB.Form frmSqlServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sql Server Credentials"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   4245
   Begin VB.Timer tmrEnableConnect 
      Interval        =   100
      Left            =   3360
      Top             =   600
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "?"
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Top             =   2280
      Width           =   315
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   2280
      Width           =   1395
   End
   Begin VB.CheckBox chkAskEachTime 
      Caption         =   "Ask password each time the application starts"
      Enabled         =   0   'False
      Height          =   195
      Left            =   60
      TabIndex        =   8
      Top             =   1980
      Width           =   4035
   End
   Begin VB.TextBox txtPassword 
      Enabled         =   0   'False
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   1560
      Width           =   2955
   End
   Begin VB.TextBox txtUserID 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Top             =   1200
      Width           =   2955
   End
   Begin VB.OptionButton optSQLServerLogin 
      Caption         =   "SQL Server Login"
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   840
      Width           =   2415
   End
   Begin VB.OptionButton optNTAuthentication 
      Caption         =   "NT Authentication"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   540
      Value           =   -1  'True
      Width           =   2415
   End
   Begin VB.TextBox txtSQLServer 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   2955
   End
   Begin VB.Label Label3 
      Caption         =   "Password"
      Height          =   195
      Left            =   60
      TabIndex        =   6
      Top             =   1620
      Width           =   1035
   End
   Begin VB.Label Label2 
      Caption         =   "User ID"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   1260
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "SQL Server"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   180
      Width           =   1035
   End
End
Attribute VB_Name = "frmSqlServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rstDBs As New ADODB.Recordset

Private Sub cmdConnect_Click()
    
    On Error GoTo excHandler
    
    SaveSetting "Team Activity Organizer", "SQL Server", "SQL Server", objCrypto.Encrypt(txtSQLServer.Text)
    
    If optNTAuthentication Then
        SaveSetting "Team Activity Organizer", "SQL Server", "NT Auth", objCrypto.Encrypt("1")
        SaveSetting "Team Activity Organizer", "SQL Server", "User ID", objCrypto.Encrypt("")
        SaveSetting "Team Activity Organizer", "SQL Server", "Password", objCrypto.Encrypt("")
        SaveSetting "Team Activity Organizer", "SQL Server", "AskEachTime", objCrypto.Encrypt("0")
    Else
        SaveSetting "Team Activity Organizer", "SQL Server", "NT Auth", objCrypto.Encrypt("0")
        SaveSetting "Team Activity Organizer", "SQL Server", "User ID", objCrypto.Encrypt(txtUserID.Text)
        If chkAskEachTime.Value = vbChecked Then
            SaveSetting "Team Activity Organizer", "SQL Server", "AskEachTime", objCrypto.Encrypt("1")
            SaveSetting "Team Activity Organizer", "SQL Server", "Password", objCrypto.Encrypt("")
        Else
            SaveSetting "Team Activity Organizer", "SQL Server", "Password", objCrypto.Encrypt(txtPassword.Text)
            SaveSetting "Team Activity Organizer", "SQL Server", "AskEachTime", objCrypto.Encrypt("0")
        End If
    End If
    
    
    Call PutSQLServerProperties
    
    Call SetSQLServerProperties(cnnMain, False)
    
    
    ' Check if dbActivity database exists
    'rstDBs.CursorLocation = adUseClient
    
    Screen.MousePointer = vbHourglass
    
    'rstDBs.Open "sp_databases", cnnMain, adOpenStatic, adLockOptimistic, adCmdText
    
    On Error Resume Next
    cnnMain.Execute "USE dbActivities"
        
    If Err.Number <> 0 Then
        If cnnMain.Errors(0).NativeError = 916 Then
            ' No rights to database
            DisplayInformation "The Application database is found, but the current user" & vbNewLine & "does not have rights to access it. Please login with another account", Err.Number, Err.Source, Err.Description
            Screen.MousePointer = vbNormal
            cnnMain.Close
            Set cnnMain = Nothing
            Exit Sub
        ElseIf cnnMain.Errors(0).NativeError = 911 Then
            ' No such database
            On Error GoTo excHandler
            Screen.MousePointer = vbNormal
        
        'rstDBs.Filter = "DATABASE_NAME='dbActivities'"
        
        'If rstDBs.RecordCount = 0 Then
            NewMSG "The Application database could not be found", COLOR_EXCEPTION
            If MsgBox("The Application database could not be found. Do you want it to be created now?", vbYesNoCancel, "Application Database Creation") = vbYes Then
                Screen.MousePointer = vbHourglass
                If Not CreateApplicationDatabase Then
                    On Error Resume Next
                    cnnMain.Execute "USE master", , adCmdText
                    cnnMain.Execute "DROP DATABASE dbActivities", , adCmdText
                    On Error GoTo excHandler
                    cnnMain.Close
                    Set cnnMain = Nothing
                    Screen.MousePointer = vbNormal
                    Exit Sub
                End If
                Screen.MousePointer = vbNormal
            Else
                Screen.MousePointer = vbNormal
                cnnMain.Close
                Set cnnMain = Nothing
                Exit Sub
            End If
        Else
            DisplayInformation "An exception occured while trying to access Application database", Err.Number, Err.Source, Err.de
            cnnMain.Close
            Set cnnMain = Nothing
            Exit Sub
        End If
    End If
    
    Set cnnMain = Nothing
    
    Screen.MousePointer = vbNormal
    
    NewMSG "Loading Server..", COLOR_NORMAL
    Load frmServer
    frmServer.Visible = True
    frmServer.WindowState = vbMaximized
    
    Unload Me
    
    Exit Sub

excHandler:
    
    Screen.MousePointer = vbNormal
    
    DisplayInformation "An Exception Occured while opening connection to " & txtSQLServer.Text, Err.Number, Err.Source, Err.Description
           
    Set cnnMain = Nothing

End Sub



Private Sub Form_Activate()
    
    Dim strEmpty            As String
    
    On Error GoTo excHandler
    
    strEmpty = objCrypto.Encrypt("")
    
    txtSQLServer.Text = objCrypto.Decrypt(GetSetting("Team Activity Organizer", "SQL Server", "SQL Server", strEmpty))
    
    If objCrypto.Decrypt(GetSetting("Team Activity Organizer", "SQL Server", "NT Auth", strEmpty)) = "1" Then
        optNTAuthentication.Value = True
    Else
        optSQLServerLogin.Value = True
        txtUserID.Text = objCrypto.Decrypt(GetSetting("Team Activity Organizer", "SQL Server", "User ID", strEmpty))
        If objCrypto.Decrypt(GetSetting("Team Activity Organizer", "SQL Server", "AskEachTime", strEmpty)) = "1" Then
            txtPassword.Text = ""
            chkAskEachTime.Value = vbChecked
        Else
            txtPassword.Text = objCrypto.Decrypt(GetSetting("Team Activity Organizer", "SQL Server", "Password", strEmpty))
            chkAskEachTime.Value = vbUnchecked
        End If
    End If
    
    If chkAskEachTime.Value = vbUnchecked And Len(txtSQLServer.Text) > 0 Then
        Call cmdConnect_Click
    End If
    
    Exit Sub

excHandler:

    DisplayInformation "Registry values are corrupt. Please re-enter SQL Server properties", Err.Number, Err.Source, Err.Description

End Sub

Private Sub optNTAuthentication_Click()
    txtUserID.Enabled = False
    txtPassword.Enabled = False
    chkAskEachTime.Enabled = False
End Sub

Private Sub optSQLServerLogin_Click()
    txtUserID.Enabled = True
    txtPassword.Enabled = True
    chkAskEachTime.Enabled = True
End Sub

Private Sub tmrEnableConnect_Timer()
    
    If Len(txtSQLServer.Text) > 0 And _
          ( _
             optNTAuthentication.Value = True Or _
             ( _
                 optSQLServerLogin.Value = True And Len(txtUserID.Text) > 0 And Len(txtPassword.Text) > 0 _
             ) _
          ) _
    Then
        
        cmdConnect.Enabled = True
    
    Else
        
        cmdConnect.Enabled = False
    
    End If

End Sub


Private Function CreateApplicationDatabase() As Boolean
    Dim strSQL              As String
    Dim lngFF               As Long
    Dim strLine             As String
    Dim blnInTrans          As Boolean
    Dim blnDBCreated        As Boolean
    Dim lngBatchCount       As Long
    Dim lngCurrentBatch     As Long
    Dim strMessage          As String
    Const strMessageBase    As String = "Creating Application Database.. #PRG#% Done"
    
    On Error GoTo excHandler
    
    Screen.MousePointer = vbHourglass
    
    lngFF = FreeFile
    
    Open App.Path & "\db.sql" For Input Access Read As lngFF
    
    Line Input #lngFF, strLine
    If Left(strLine, 14) = "--BATCH_COUNT=" Then
        If IsNumeric(Mid(strLine, 15)) Then
            lngBatchCount = Mid(strLine, 15)
        End If
    End If
    
    Load frmDBCreateProgress
    
    With frmDBCreateProgress
        .Left = 150
        .Top = 150
        .Show
        .prgDB.Max = lngBatchCount
        .prgDB.Min = 0
        .lblMessage = "Please Wait.."
        DoEvents: DoEvents: DoEvents: DoEvents
    End With
    
    lngCurrentBatch = 0
    
    cnnMain.BeginTrans
    
    blnInTrans = True
    blnDBCreated = True
    
    Do While Not EOF(lngFF)
        Line Input #lngFF, strLine
        If Left(strLine, 2) = "--" Then
            strMessage = Mid(strLine, 3)
            frmDBCreateProgress.lblMessage = strMessage
            DoEvents: DoEvents: DoEvents: DoEvents
        End If
        If Trim(Replace(strLine, Chr(9), "")) = "GO" Then
            cnnMain.Execute strSQL, , adCmdText
            If lngBatchCount > 0 Then
                lngCurrentBatch = lngCurrentBatch + 1
                NewMSG "STEP " & lngCurrentBatch & "/" & lngBatchCount & " : " & strMessage, COLOR_NORMAL
                frmDBCreateProgress.Caption = Replace(strMessageBase, "#PRG#", CLng((lngCurrentBatch / lngBatchCount) * 100))
                frmDBCreateProgress.prgDB.Value = lngCurrentBatch
                DoEvents: DoEvents: DoEvents: DoEvents
            End If
            blnDBCreated = True
            NewMSG "Step has been completed..", COLOR_SUCCESS
            strSQL = ""
        Else
            strSQL = strSQL & strLine & vbNewLine
        End If
        
    Loop

    cnnMain.CommitTrans
    
    NewMSG "Application Database has been created successfully", COLOR_SUCCESS

    Close #lngFF
    
    CreateApplicationDatabase = True
    
    Screen.MousePointer = vbNormal
    
    Unload frmDBCreateProgress
    
    Exit Function

excHandler:
    
    Unload frmDBCreateProgress
    
    If blnInTrans Then
        cnnMain.RollbackTrans
    End If
    
    Close #lngFF
    
    Screen.MousePointer = vbNormal
    
    DisplayInformation "An Exception Occured while creating Application Database", Err.Number, Err.Source, Err.Description
            
    CreateApplicationDatabase = False
    
End Function

