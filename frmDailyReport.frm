VERSION 5.00
Begin VB.Form frmDailyReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daily Activity Reports"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   4440
   Begin VB.CommandButton cmdGenerateReport 
      Caption         =   "Generate Report"
      Height          =   315
      Left            =   60
      TabIndex        =   11
      Top             =   4740
      Width           =   4275
   End
   Begin VB.Frame Frame2 
      Caption         =   " Output "
      Height          =   1635
      Left            =   60
      TabIndex        =   10
      Top             =   3000
      Width           =   4275
      Begin VB.OptionButton optXMLFile 
         Caption         =   "XML File"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   960
         Width           =   3855
      End
      Begin VB.OptionButton optPrinter 
         Caption         =   "Printer"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   1260
         Width           =   3855
      End
      Begin VB.OptionButton optTextFile 
         Caption         =   "Text File"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   660
         Width           =   3855
      End
      Begin VB.OptionButton optScreen 
         Caption         =   "Screen"
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   360
         Value           =   -1  'True
         Width           =   3855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Daily Activity Report Parameters "
      Height          =   2535
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   4275
      Begin VB.ComboBox cmbUserIDs 
         Height          =   315
         Left            =   60
         TabIndex        =   16
         Top             =   1680
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.CheckBox chkForAll 
         Caption         =   "Generate report for all users"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2160
         Width           =   3855
      End
      Begin VB.ComboBox cmbUsers 
         Height          =   315
         Left            =   180
         TabIndex        =   7
         Top             =   1680
         Width           =   3975
      End
      Begin VB.ComboBox cmbYear 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3240
         TabIndex        =   6
         Top             =   960
         Width           =   915
      End
      Begin VB.ComboBox cmbMonth 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2460
         TabIndex        =   5
         Top             =   960
         Width           =   735
      End
      Begin VB.ComboBox cmbDay 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmDailyReport.frx":0000
         Left            =   1680
         List            =   "frmDailyReport.frx":0002
         TabIndex        =   4
         Top             =   960
         Width           =   735
      End
      Begin VB.OptionButton optSpecificDate 
         Caption         =   "Specific Date"
         Height          =   255
         Left            =   180
         TabIndex        =   3
         Top             =   1020
         Width           =   1935
      End
      Begin VB.OptionButton optYesterday 
         Caption         =   "Yesterday"
         Height          =   255
         Left            =   180
         TabIndex        =   2
         Top             =   660
         Width           =   1755
      End
      Begin VB.OptionButton optToday 
         Caption         =   "Today"
         Height          =   255
         Left            =   180
         TabIndex        =   1
         Top             =   300
         Value           =   -1  'True
         Width           =   1755
      End
      Begin VB.Label Label1 
         Caption         =   "Select user"
         Height          =   255
         Left            =   180
         TabIndex        =   8
         Top             =   1440
         Width           =   2115
      End
   End
   Begin VB.Label lblAlert 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   60
      TabIndex        =   17
      Top             =   2700
      Width           =   4275
   End
End
Attribute VB_Name = "frmDailyReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rstUsers As New ADODB.Recordset

Private Sub chkForAll_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If chkForAll.Value = vbChecked Then
        cmbUsers.Enabled = False
    Else
        cmbUsers.Enabled = True
    End If
End Sub

Private Sub cmdGenerateReport_Click()

    Dim dtmReportDate       As Date
    Dim intReportUserID     As Long
    Dim lngReportOutputType As ReportTypes
    Dim rstReport           As New ADODB.Recordset
    
    lblAlert.Caption = ""
    
    If optToday Then
        dtmReportDate = Now
    ElseIf optYesterday Then
        dtmReportDate = DateAdd("d", -1, Now)
    Else
        If Not IsDate(cmbYear & "-" & cmbMonth & "-" & cmbDay) Then
            lblAlert.Caption = "Please specify a valid date"
            Exit Sub
        Else
            dtmReportDate = CDate(cmbYear & "-" & cmbMonth & "-" & cmbDay)
        End If
    End If
    
    If chkForAll.Value = vbChecked Then
        intReportUserID = 0
    Else
        If cmbUsers.ListIndex <= 0 Then
            lblAlert.Caption = "Please select a user"
            Exit Sub
        Else
            intReportUserID = cmbUserIDs.List(cmbUsers.ListIndex)
        End If
    End If

    If optScreen Then
        lngReportOutputType = REPORT_TYPE_SCREEN
    ElseIf optTextFile Then
        lngReportOutputType = REPORT_TYPE_TEXT_FILE
    ElseIf optXMLFile Then
        lngReportOutputType = REPORT_TYPE_XML_FILE
    ElseIf optPrinter Then
        lngReportOutputType = REPORT_TYPE_PRINTER
    End If
    
    'Call objData.GetActivityReports(rstReport, dtmReportDate, intReportUserID)
    
    If Not rstReport.EOF Then
        Screen.MousePointer = vbHourglass
        If CreatePrintReport(rstReport) Then
            lblAlert.Caption = "Your report has been sent to printer"
        Else
            lblAlert.Caption = "Failed to create report"
        End If
        Screen.MousePointer = vbNormal
        
    End If
    
End Sub

Private Sub Form_Load()
    Dim lngIdx As Long
    
    winOpen.blnDailyReports = True
    
    cmbDay.Clear
    For lngIdx = 1 To 31
        cmbDay.AddItem lngIdx
    Next lngIdx
    cmbDay.Text = Day(Now)
    
    cmbMonth.Clear
    For lngIdx = 1 To 12
        cmbMonth.AddItem lngIdx
    Next lngIdx
    cmbMonth.Text = Month(Now)
    
    cmbYear.Clear
    For lngIdx = Year(Now) - 10 To Year(Now)
        cmbYear.AddItem lngIdx
    Next lngIdx
    cmbYear.Text = Year(Now)
    
    Call objData.GetApplicationData(rstUsers, APP_DATA_TYPE_USERS)
    
    cmbUsers.Clear
    cmbUserIDs.Clear
    cmbUsers.AddItem "Please Select.."
    cmbUserIDs.AddItem ""
    
    If rstUsers.RecordCount > 0 Then
        Do While Not rstUsers.EOF
            If rstUsers("bitUserIsActive") Then
                cmbUsers.AddItem rstUsers("strUserFirstName") & " " & rstUsers("strUserLastName")
                cmbUserIDs.AddItem rstUsers("intUserID")
            End If
            rstUsers.MoveNext
        Loop
    End If
    
    cmbUsers.Text = "Please Select.."
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    winOpen.blnDailyReports = False
End Sub

Private Sub optSpecificDate_Click()
    cmbDay.Enabled = True
    cmbMonth.Enabled = True
    cmbYear.Enabled = True
End Sub

Private Sub optToday_Click()
    cmbDay.Enabled = False
    cmbMonth.Enabled = False
    cmbYear.Enabled = False
End Sub

Private Sub optYesterday_Click()
    cmbDay.Enabled = False
    cmbMonth.Enabled = False
    cmbYear.Enabled = False
End Sub


