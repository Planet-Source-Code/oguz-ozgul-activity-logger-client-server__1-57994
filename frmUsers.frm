VERSION 5.00
Begin VB.Form frmUsers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Administration - Users"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6675
   Icon            =   "frmUsers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "No Action"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   5460
      Width           =   2115
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   25
      Top             =   5460
      Width           =   2115
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Update / Insert"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4440
      TabIndex        =   24
      Top             =   5460
      Width           =   2115
   End
   Begin VB.TextBox txtEMail 
      Enabled         =   0   'False
      Height          =   345
      Left            =   1740
      MaxLength       =   128
      TabIndex        =   23
      Top             =   4980
      Width           =   4815
   End
   Begin VB.TextBox txtCellPhone 
      Enabled         =   0   'False
      Height          =   345
      Left            =   1740
      MaxLength       =   32
      TabIndex        =   21
      Top             =   4560
      Width           =   1815
   End
   Begin VB.TextBox txtExtension 
      Enabled         =   0   'False
      Height          =   345
      Left            =   1740
      MaxLength       =   8
      TabIndex        =   19
      Top             =   4140
      Width           =   615
   End
   Begin VB.TextBox txtPhone 
      Enabled         =   0   'False
      Height          =   345
      Left            =   1740
      MaxLength       =   32
      TabIndex        =   17
      Top             =   3720
      Width           =   1815
   End
   Begin VB.TextBox txtTitle 
      Enabled         =   0   'False
      Height          =   345
      Left            =   1740
      MaxLength       =   128
      TabIndex        =   15
      Top             =   3300
      Width           =   4815
   End
   Begin VB.TextBox txtLastName 
      Enabled         =   0   'False
      Height          =   345
      Left            =   1740
      MaxLength       =   64
      TabIndex        =   13
      Top             =   2880
      Width           =   2775
   End
   Begin VB.TextBox txtFirstName 
      Enabled         =   0   'False
      Height          =   345
      Left            =   1740
      MaxLength       =   64
      TabIndex        =   11
      Top             =   2460
      Width           =   2775
   End
   Begin VB.TextBox txtPassword2 
      Enabled         =   0   'False
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1740
      MaxLength       =   64
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   2040
      Width           =   3855
   End
   Begin VB.TextBox txtPassword1 
      Enabled         =   0   'False
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1740
      MaxLength       =   64
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   1620
      Width           =   3855
   End
   Begin VB.TextBox txtUserName 
      Enabled         =   0   'False
      Height          =   345
      Left            =   1740
      MaxLength       =   64
      TabIndex        =   5
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CommandButton cmdNewUser 
      Caption         =   "Add New User"
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   780
      Width           =   1575
   End
   Begin VB.ComboBox cmbUserIDs 
      Height          =   315
      Left            =   4320
      TabIndex        =   2
      Top             =   60
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.ComboBox cmbUsers 
      Height          =   315
      ItemData        =   "frmUsers.frx":0442
      Left            =   120
      List            =   "frmUsers.frx":0444
      TabIndex        =   0
      Top             =   360
      Width           =   6495
   End
   Begin VB.Label lblUserID 
      Caption         =   "N/A"
      Height          =   255
      Left            =   1740
      TabIndex        =   27
      Top             =   840
      Width           =   1515
   End
   Begin VB.Label Label12 
      Caption         =   "User ID"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   840
      Width           =   1515
   End
   Begin VB.Label Label11 
      Caption         =   "E-Mail"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   5040
      Width           =   1515
   End
   Begin VB.Label Label10 
      Caption         =   "Cell Phone"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   4620
      Width           =   1515
   End
   Begin VB.Label Label9 
      Caption         =   "Extension Number"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   4200
      Width           =   1515
   End
   Begin VB.Label Label8 
      Caption         =   "Phone"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3780
      Width           =   1515
   End
   Begin VB.Label Label7 
      Caption         =   "Title"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3360
      Width           =   1515
   End
   Begin VB.Label Label6 
      Caption         =   "Last Name"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2940
      Width           =   1515
   End
   Begin VB.Label Label5 
      Caption         =   "First Name"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   1515
   End
   Begin VB.Label Label4 
      Caption         =   "Password (re-entry)"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2100
      Width           =   1515
   End
   Begin VB.Label Label3 
      Caption         =   "Password"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1515
   End
   Begin VB.Label Label2 
      Caption         =   "User Name"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1260
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "Select User"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4035
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rstUsers        As New ADODB.Recordset
Dim cmdUsers        As New ADODB.Command

Private lngDisabledColor        As Long
Private lngEnabledColor         As Long



Private Sub cmbUsers_Click()
    
    On Error GoTo excHandler
    
    If cmbUsers.ListIndex > -1 Then
    
        SetSQLServerProperties cnnMain, True
        
        With cmdUsers
            .CommandType = adCmdStoredProc
            .CommandText = "prGetUser"
            .ActiveConnection = cnnMain
            .Parameters.Refresh
            .Parameters("@intUserID").Value = cmbUserIDs.List(cmbUsers.ListIndex)
            Set rstUsers = .Execute
        End With
        
        If rstUsers.RecordCount = 1 Then
            
            NewMSG "User Information ready..", COLOR_NORMAL
            FillInputAreas
            EnableInputFields False
        
        Else
        
            NewMSG "User Does Not Exist", COLOR_EXCEPTION
            Err.Raise -1, "prGetUser", "User does not exist. User ID: " & cmbUserIDs.List(cmbUsers.ListIndex)
        
        End If

    Else
    
        NewMSG "Please select a user", COLOR_WARNING
        
    End If
    
    
    
ExitSub:
    
    Set cmdUsers = Nothing
    Set rstUsers = Nothing
    Set cnnMain = Nothing

    Exit Sub

excHandler:

    DisplayInformation "User Details could not be retrieved", Err.Number, Err.Source, Err.Description
    Resume ExitSub

End Sub

Private Sub cmdCancel_Click()
    EmptyInputFields
    DisableInputFields
    cmbUsers.Enabled = True
    cmbUsers.Text = ""
    cmdNewUser.Enabled = True
    cmdInsert.Enabled = False
    cmdCancel.Enabled = False
    NewMSG "User Manager Ready..", COLOR_NORMAL
End Sub

Private Sub cmdInsert_Click()
    On Error GoTo excHandler
    Dim blnCheckOK      As Boolean
    If IsNumeric("" & lblUserID.Caption) Then
        blnCheckOK = CheckForm(True)
    Else
        blnCheckOK = CheckForm(False)
    End If
    
    If blnCheckOK Then
        Screen.MousePointer = vbHourglass
        SetSQLServerProperties cnnMain, True
        If Not IsNumeric("" & lblUserID.Caption) Then ' INSERT NEW USER
            With cmdUsers
                .CommandType = adCmdStoredProc
                .CommandText = "prAddUser"
                .ActiveConnection = cnnMain
                .Parameters.Refresh
                .Parameters("@strUserName").Value = txtUserName.Text
                .Parameters("@strUserPassword").Value = txtPassword1.Text
                .Parameters("@strUserFirstName").Value = txtFirstName.Text
                .Parameters("@strUserLastName").Value = txtLastName.Text
                .Parameters("@strUserTitle").Value = txtTitle.Text
                .Parameters("@strUserPhone").Value = txtPhone.Text
                .Parameters("@strUserExtension").Value = txtExtension.Text
                .Parameters("@strUserCellPhone").Value = txtCellPhone.Text
                .Parameters("@strUserEMail").Value = txtEMail.Text
                .Execute
                If IsNull(.Parameters("@strException").Value) Then
                    lblUserID.Caption = .Parameters("@intUserID").Value
                    cmbUserIDs.AddItem .Parameters("@intUserID").Value
                    If txtTitle.Text = "N/A" Then
                        cmbUsers.AddItem txtFirstName.Text & " " & txtLastName.Text
                    Else
                        cmbUsers.AddItem txtFirstName.Text & " " & txtLastName.Text & " (" & txtTitle.Text & ")"
                    End If
                    NewMSG "User has been added successfully", COLOR_SUCCESS
                Else
                    NewMSG "Failed to add user", COLOR_EXCEPTION
                    Err.Raise -1, "prAddUser", "" & .Parameters("@strException").Value
                End If
            End With
        Else
            With cmdUsers
                .CommandType = adCmdStoredProc
                .CommandText = "prUpdateUser"
                .ActiveConnection = cnnMain
                .Parameters.Refresh
                .Parameters("@intUserID").Value = lblUserID.Caption
                .Parameters("@strUserFirstName").Value = txtFirstName.Text
                .Parameters("@strUserLastName").Value = txtLastName.Text
                .Parameters("@strUserTitle").Value = txtTitle.Text
                .Parameters("@strUserPhone").Value = txtPhone.Text
                .Parameters("@strUserExtension").Value = txtExtension.Text
                .Parameters("@strUserCellPhone").Value = txtCellPhone.Text
                .Parameters("@strUserEMail").Value = txtEMail.Text
                .Execute
                If IsNull(.Parameters("@strException").Value) Then
                    If txtTitle.Text = "N/A" Then
                        cmbUsers.Text = txtFirstName.Text & " " & txtLastName.Text
                    Else
                        cmbUsers.Text = txtFirstName.Text & " " & txtLastName.Text & " (" & txtTitle.Text & ")"
                    End If
                    NewMSG "User information has been updated", COLOR_SUCCESS
                Else
                    NewMSG "User information could not be updated", COLOR_EXCEPTION
                    Err.Raise -1, "prAddUser", "" & .Parameters("@strException").Value
                End If
            End With
        End If
        
    Else
    
        NewMSG "Invalid or insufficient input", COLOR_WARNING
        Exit Sub
    
    End If
    
    EnableInputFields False
    
ExitSub:
    
    Screen.MousePointer = vbNormal
    Set cmdUsers = Nothing
    Set cnnMain = Nothing
    
    Exit Sub
    
excHandler:

    DisplayInformation "The User could not be added", Err.Number, Err.Source, Err.Description
    Resume ExitSub

End Sub

Private Sub cmdNewUser_Click()
    EmptyInputFields
    EnableInputFields True
End Sub

Private Sub cmdRemove_Click()
    On Error GoTo excHandler
    Dim lngRecsAffected     As Long
    Dim lngUserIndex        As Long
        
    If cmbUsers.ListIndex > -1 Then
        
        SetSQLServerProperties cnnMain, True
    
        With cmdUsers
            .CommandType = adCmdStoredProc
            .CommandText = "prRemoveUser"
            .ActiveConnection = cnnMain
            .Parameters.Refresh
        
            .Parameters("@intUserID") = cmbUserIDs.List(cmbUsers.ListIndex)
            
            .Execute lngRecsAffected
        
        End With
        
        If lngRecsAffected = 0 Then
            NewMSG "User could not be removed", COLOR_EXCEPTION
            Err.Raise -1, "prRemoveUser", "Update Statement has affected no rows, user could not be removed"
        End If
        
        NewMSG "User has been removed successfully", COLOR_SUCCESS
        
        lngUserIndex = cmbUsers.ListIndex
        cmbUsers.RemoveItem lngUserIndex
        cmbUserIDs.RemoveItem lngUserIndex
        EmptyInputFields
        DisableInputFields
        
    Else
    
        NewMSG "Please select a user to remove", COLOR_WARNING
        
        Exit Sub
    
    End If

ExitSub:
    Set cmdUsers = Nothing
    Set cnnMain = Nothing
    
    Exit Sub

excHandler:

    DisplayInformation "Could not remove user", Err.Number, Err.Source, Err.Description
    Resume ExitSub


End Sub

Private Sub Form_Load()
    On Error GoTo excHandler
    
    lngEnabledColor = RGB(255, 255, 255)
    lngDisabledColor = RGB(192, 192, 192)
    
    DisableInputFields
    
    Call objData.GetApplicationData(rstUsers, APP_DATA_TYPE_USERS)

    If rstUsers.RecordCount > 0 Then
        Do While Not rstUsers.EOF
            If rstUsers("strUserName") <> "administrator" Then
                cmbUserIDs.AddItem rstUsers("intUserID")
                If rstUsers("strUserTitle") = "N/A" Then
                    cmbUsers.AddItem rstUsers("strUserFirstName") & " " & rstUsers("strUserLastName")
                Else
                    cmbUsers.AddItem rstUsers("strUserFirstName") & " " & rstUsers("strUserLastName") & " (" & rstUsers("strUserTitle") & ")"
                End If
            End If
            rstUsers.MoveNext
        Loop
    End If
    
    NewMSG rstUsers.RecordCount & " Users found", COLOR_NORMAL
    
    rstUsers.Close
    
ExitSub:
    
    Set rstUsers = Nothing
    Set cnnMain = Nothing
    
    Exit Sub

excHandler:
    
    DisplayInformation "User Information could not be loaded", Err.Number, Err.Source, Err.Description
    Resume ExitSub

End Sub

Private Function EmptyInputFields()
    lblUserID.Caption = "N/A"
    txtUserName.Text = ""
    txtPassword1.Text = ""
    txtPassword2.Text = ""
    txtFirstName.Text = ""
    txtLastName.Text = ""
    txtTitle.Text = ""
    txtPhone.Text = ""
    txtExtension.Text = ""
    txtCellPhone.Text = ""
    txtEMail.Text = ""
End Function


Private Function FillInputAreas()
    lblUserID.Caption = rstUsers("intUserID").Value
    txtUserName.Text = rstUsers("strUserName").Value
    txtPassword1.Text = "****************"
    txtPassword2.Text = "****************"
    txtFirstName.Text = rstUsers("strUserFirstName").Value
    txtLastName.Text = rstUsers("strUserLastName").Value
    txtTitle.Text = rstUsers("strUserTitle").Value
    txtPhone.Text = rstUsers("strUserPhone").Value
    txtExtension.Text = rstUsers("strUserExtension").Value
    txtCellPhone.Text = rstUsers("strUserCellPhone").Value
    txtEMail.Text = rstUsers("strUserEMail").Value
End Function





Private Function EnableInputFields(ByVal blnWithCredentials As Boolean)

    If blnWithCredentials Then
        lblUserID.Caption = "N/A"
        cmbUsers.Enabled = False
        cmdNewUser.Enabled = False
        cmdInsert.Enabled = True
        cmdRemove.Enabled = False
        cmdCancel.Enabled = True
        txtUserName.Enabled = True
        txtUserName.BackColor = lngEnabledColor
        txtPassword1.Enabled = True
        txtPassword1.BackColor = lngEnabledColor
        txtPassword2.Enabled = True
        txtPassword2.BackColor = lngEnabledColor
    Else
        cmbUsers.Enabled = False
        cmdNewUser.Enabled = True
        cmdInsert.Enabled = True
        cmdRemove.Enabled = True
        cmdCancel.Enabled = True
        txtUserName.Enabled = False
        txtUserName.BackColor = lngDisabledColor
        txtPassword1.Enabled = False
        txtPassword1.BackColor = lngDisabledColor
        txtPassword2.Enabled = False
        txtPassword2.BackColor = lngDisabledColor
    End If
    txtFirstName.Enabled = True
    txtFirstName.BackColor = lngEnabledColor
    txtLastName.Enabled = True
    txtLastName.BackColor = lngEnabledColor
    txtTitle.Enabled = True
    txtTitle.BackColor = lngEnabledColor
    txtPhone.Enabled = True
    txtPhone.BackColor = lngEnabledColor
    txtExtension.Enabled = True
    txtExtension.BackColor = lngEnabledColor
    txtCellPhone.Enabled = True
    txtCellPhone.BackColor = lngEnabledColor
    txtEMail.Enabled = True
    txtEMail.BackColor = lngEnabledColor

End Function

Private Function DisableInputFields()

    cmdInsert.Enabled = False
    cmdRemove.Enabled = False
    cmdCancel.Enabled = False
    cmbUsers.Enabled = True
    txtUserName.Enabled = False
    txtUserName.BackColor = lngDisabledColor
    txtPassword1.Enabled = False
    txtPassword1.BackColor = lngDisabledColor
    txtPassword2.Enabled = False
    txtPassword2.BackColor = lngDisabledColor
    txtFirstName.Enabled = False
    txtFirstName.BackColor = lngDisabledColor
    txtLastName.Enabled = False
    txtLastName.BackColor = lngDisabledColor
    txtTitle.Enabled = False
    txtTitle.BackColor = lngDisabledColor
    txtPhone.Enabled = False
    txtPhone.BackColor = lngDisabledColor
    txtExtension.Enabled = False
    txtExtension.BackColor = lngDisabledColor
    txtCellPhone.Enabled = False
    txtCellPhone.BackColor = lngDisabledColor
    txtEMail.Enabled = False
    txtEMail.BackColor = lngDisabledColor

End Function


Private Function CheckForm(ByVal blnForInsert As Boolean) As Boolean

    CheckForm = False
    
    If blnForInsert Then
        If Not CheckUserName(txtUserName.Text) Then
            NewMSG "The user name specified is invalid. Please use letters, digits and underscore only, do not begin with a digit or underscore", COLOR_WARNING
            Exit Function
        End If
        If Len(txtPassword1.Text) = 0 Or Len(txtPassword2.Text) = 0 Then
            NewMSG "Please fill both password areas", COLOR_WARNING
            Exit Function
        End If
        If txtPassword1.Text <> txtPassword2.Text Then
            NewMSG "Password areas do not match. The two password and the password re-entry should be the same", COLOR_WARNING
            Exit Function
        End If
    End If
    If Len(txtFirstName) = 0 Then
        NewMSG "User First Name should be specified", COLOR_WARNING
        Exit Function
    End If
    If Len(txtLastName) = 0 Then
        NewMSG "User Last Name should be specified", COLOR_WARNING
        Exit Function
    End If
    If Len(txtTitle.Text) = 0 Then
        txtTitle.Text = "N/A"
    End If
    If Len(txtPhone.Text) = 0 Then
        txtPhone.Text = "N/A"
    End If
    If Len(txtExtension.Text) = 0 Then
        txtExtension.Text = "N/A"
    End If
    If Len(txtCellPhone.Text) = 0 Then
        txtCellPhone.Text = "N/A"
    End If
    If Len(txtEMail.Text) = 0 Then
        txtEMail.Text = "N/A"
    End If
    CheckForm = True
End Function



Private Function CheckUserName(ByVal strText As String) As Boolean
    
    Const VALID_CHARS = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_"
    
    Const NOT_AT_THE_BEGINNING = "0123456789_"

    Dim lngPos      As Long
    
    If Len(strText) = 0 Then
        Exit Function
    End If
    
    If InStr(1, NOT_AT_THE_BEGINNING, Left(strText, 1)) > 0 Then
        Exit Function
    End If

    For lngPos = 1 To Len(strText)
        If InStr(1, VALID_CHARS, Mid(strText, lngPos, 1)) <= 0 Then
            Exit Function
        End If
    Next
    
    CheckUserName = True
    
End Function

