VERSION 5.00
Begin VB.Form frmProjects 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Projects"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6405
   Icon            =   "frmProjects.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   6405
   Begin VB.Timer tmrAnim 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   2100
      Top             =   540
   End
   Begin VB.ComboBox cmbProjectIDs 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   540
      Width           =   1575
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox txtProjectName 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   1260
      Width           =   6135
   End
   Begin VB.ComboBox cmbProjects 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6195
   End
   Begin VB.Label Label1 
      Caption         =   "Messages:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1860
      Width           =   2475
   End
   Begin VB.Label lblResult 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   2100
      Width           =   6135
   End
   Begin VB.Label lblFormType 
      Caption         =   "Add new project"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1020
      Width           =   3855
   End
End
Attribute VB_Name = "frmProjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cnnGeneral                  As New ADODB.Connection
Private cmdGeneral                  As New ADODB.Command
Private rstGeneral                  As New ADODB.Recordset

Private Enum FormType
    FORM_TYPE_PROJECTS = 1
    FORM_TYPE_ACTIVITY_TYPES = 2
    FORM_TYPE_PAUSE_CAUSES = 3
End Enum

Private strGetStoredProcedureName   As String
Private strSetStoredProcedureName   As String
Private strDelStoredProcedureName   As String
Private strIdentityField            As String
Private strNameField                As String
Private lngFormType                 As FormType
Private strResultText               As String


Function InitializeForm(ByVal strType As String) As Long
    
    On Error GoTo excHandler
    
    Select Case strType
        Case "PRJS"
            strGetStoredProcedureName = "prGetProjects"
            strSetStoredProcedureName = "prAddProject"
            strDelStoredProcedureName = "prRemoveProject"
            strIdentityField = "intProjectID"
            strNameField = "strProjectName"
            Caption = "Projects"
            lblFormType.Caption = "Add New Project"
            winOpen.blnProjects = True
            lngFormType = FORM_TYPE_PROJECTS
        Case "ACTT"
            strGetStoredProcedureName = "prGetActivityTypes"
            strSetStoredProcedureName = "prAddActivityType"
            strDelStoredProcedureName = "prRemoveActivityType"
            strIdentityField = "intActivityTypeID"
            strNameField = "strActivityTypeName"
            Caption = "Activity Types"
            lblFormType.Caption = "Add New Activity Type"
            winOpen.blnActivityTypes = True
            lngFormType = FORM_TYPE_ACTIVITY_TYPES
        Case "PAUS"
            strGetStoredProcedureName = "prGetPauseCauses"
            strSetStoredProcedureName = "prAddPauseCause"
            strDelStoredProcedureName = "prRemovePauseCause"
            strIdentityField = "intPauseCauseID"
            strNameField = "strPauseCause"
            Caption = "Pause Causes"
            lblFormType.Caption = "Add New Pause Cause"
            winOpen.blnPauseCauses = True
            lngFormType = FORM_TYPE_PAUSE_CAUSES
        Case Else
            Unload Me
    End Select
    
    SetSQLServerProperties cnnGeneral, True
    
    With cmdGeneral
        .CommandType = adCmdStoredProc
        .CommandText = strGetStoredProcedureName
        .ActiveConnection = cnnGeneral
    End With
    
    rstGeneral.Open cmdGeneral
    
    cmbProjectIDs.Clear
    cmbProjects.Clear
    
    Do While Not rstGeneral.EOF
        cmbProjectIDs.AddItem rstGeneral(strIdentityField).Value
        cmbProjects.AddItem rstGeneral(strNameField).Value
        rstGeneral.MoveNext
    Loop
        
    Set rstGeneral = Nothing
    Set cmdGeneral = Nothing
    Set cnnGeneral = Nothing
    
    Me.Show

    Exit Function

excHandler:
    
    DisplayInformation "Database Connection Error", Err.Number, Err.Source, Err.Description
    
    Unload Me

End Function




Private Sub cmbProjects_Click()
    If cmbProjects.ListIndex > -1 Then
        cmdRemove.Enabled = True
    Else
        cmdRemove.Enabled = False
    End If
End Sub

Private Sub cmdAdd_Click()
    On Error GoTo excHandler

    SetSQLServerProperties cnnGeneral, True
    
    With cmdGeneral
        .CommandType = adCmdStoredProc
        .CommandText = strSetStoredProcedureName
        .ActiveConnection = cnnGeneral
        .Parameters.Refresh
        .Parameters("@" & strNameField).Value = txtProjectName.Text
        .Execute
        If IsNull(.Parameters("@" & strIdentityField).Value) Then
            Err.Raise -1, "Insert Item", "The item could not be added"
        Else
            cmbProjectIDs.AddItem .Parameters("@" & strIdentityField).Value
            cmbProjects.AddItem txtProjectName.Text
            Select Case lngFormType
                Case FORM_TYPE_PROJECTS
                    ReDim Preserve arrProjects(UBound(arrProjects) + 1)
                    With arrProjects(UBound(arrProjects))
                        .ProjectID = cmdGeneral.Parameters("@" & strIdentityField).Value
                        .ProjectName = txtProjectName.Text
                        .Visible = True
                    End With
                Case FORM_TYPE_ACTIVITY_TYPES
                    ReDim Preserve arrActivityTypes(UBound(arrActivityTypes) + 1)
                    With arrActivityTypes(UBound(arrActivityTypes))
                        .ActivityTypeID = cmdGeneral.Parameters("@" & strIdentityField).Value
                        .ActivityTypeName = txtProjectName.Text
                        .Visible = True
                    End With
                Case FORM_TYPE_PAUSE_CAUSES
                    ReDim Preserve arrPauseCauses(UBound(arrPauseCauses) + 1)
                    With arrPauseCauses(UBound(arrPauseCauses))
                        .PauseCauseID = cmdGeneral.Parameters("@" & strIdentityField).Value
                        .strPauseCause = txtProjectName.Text
                        .Visible = True
                    End With
            End Select
        End If
    End With
    
    NewMSG "Record has been added successfully", COLOR_SUCCESS
    txtProjectName.Text = ""
    
ExitSub:
    
    Set rstGeneral = Nothing
    Set cmdGeneral = Nothing
    Set cnnGeneral = Nothing
    
    Exit Sub

excHandler:
    
    Dim lngError    As Long
    Dim strError    As String
    Dim strSource   As String
    
    lngError = Err.Number
    strError = Err.Description
    strSource = Err.Source
    
    LogException "frmProjects.cmdAdd_Click()", lngError, strError, "Could not add item", True
    DisplayInformation "The operation has failed", lngError, strSource, strError
    
    Resume ExitSub

End Sub

Private Sub cmdRemove_Click()

    On Error GoTo excHandler

    SetSQLServerProperties cnnGeneral, True
    
    With cmdGeneral
        .CommandType = adCmdStoredProc
        .CommandText = strDelStoredProcedureName
        .ActiveConnection = cnnGeneral
        .Parameters.Refresh
        .Parameters("@" & strIdentityField).Value = cmbProjectIDs.List(cmbProjects.ListIndex)
        .Execute
    End With
    
    cmbProjectIDs.RemoveItem cmbProjects.ListIndex
    cmbProjects.RemoveItem cmbProjects.ListIndex

ExitSub:
    
    Set rstGeneral = Nothing
    Set cmdGeneral = Nothing
    Set cnnGeneral = Nothing
    
    Exit Sub

excHandler:
    
    LogException "frmProjects.cmdRemove_Click()", Err.Number, Err.Description, "Could not remove item", True
    DisplayInformation "The operation has failed", Err.Number, Err.Source, Err.Description
    
    Resume ExitSub

End Sub

Private Sub Form_Load()
    Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Select Case lngFormType
        Case FORM_TYPE_PROJECTS
            winOpen.blnProjects = False
        Case FORM_TYPE_ACTIVITY_TYPES
            winOpen.blnActivityTypes = False
        Case FORM_TYPE_PAUSE_CAUSES
            winOpen.blnPauseCauses = False
    End Select
End Sub

Private Sub txtProjectName_Change()
    If Len(txtProjectName.Text) > 0 Then
        cmdAdd.Enabled = True
    Else
        cmdAdd.Enabled = False
    End If
End Sub
