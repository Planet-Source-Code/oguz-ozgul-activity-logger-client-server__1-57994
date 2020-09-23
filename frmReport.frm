VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Activity Reports"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   11685
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy as Tab Delimited Text"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4440
      TabIndex        =   6
      Top             =   360
      Width           =   3255
   End
   Begin MSFlexGridLib.MSFlexGrid flexReport 
      Height          =   3735
      Left            =   60
      TabIndex        =   5
      Top             =   780
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   6588
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox cmbDay1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   675
   End
   Begin VB.ComboBox cmbMonth1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   780
      TabIndex        =   2
      Top             =   360
      Width           =   675
   End
   Begin VB.ComboBox cmbYear1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton cmdGenerateReport 
      Caption         =   "Query"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Tarih Aralýðý"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   60
      Width           =   4455
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCopy_Click()
    Dim strData
    Dim lngDay
    
    With flexReport
        For lngDay = 0 To .Rows - 1 ' Include headers
            strData = strData & .TextMatrix(lngDay, 0) & Chr(9) & .TextMatrix(lngDay, 1) & Chr(9) & .TextMatrix(lngDay, 2) & vbCrLf
        Next lngDay
    End With
    
    With Clipboard
        .Clear
        .SetText strData
    End With
End Sub

Private Sub cmdGenerateReport_Click()
    Dim dtmDate1 As Date
    Dim dtmDate2 As Date
    Dim dtmDate3 As Date
        
    If IsDate(cmbYear1.Text & "-" & cmbMonth1.Text & "-" & cmbDay1.Text) Then
        dtmDate1 = CDate(cmbYear1.Text & "-" & cmbMonth1.Text & "-" & cmbDay1.Text)
        ' Check if date2 is specified
'        If cmbYear2.Text <> "" Or cmbMonth2.Text <> "" Or cmbDay2.Text <> "" Then
'            If IsDate(cmbYear2.Text & "-" & cmbMonth2.Text & "-" & cmbDay2.Text) Then
'                dtmDate2 = CDate(cmbYear2.Text & "-" & cmbMonth2.Text & "-" & cmbDay2.Text)
'            Else
'                MsgBox "The end date is not a valid date..", vbExclamation
'                Exit Sub
'            End If
'        Else
'            dtmDate2 = dtmDate1
'        End If
'        ' Replace dates if necessary
        'If DateDiff("d", dtmDate1, dtmDate2) < 0 Then
        '    dtmDate3 = dtmDate2
        '    dtmDate2 = dtmDate1
        '    dtmDate1 = dtmDate3
        'End If
    
                
        ' Send the report request to the server
        frmMain.sckServer.SendData objCrypto.Encrypt("RPRT" & Format(dtmDate1, "YYYY-MM-DD")) & vbCrLf
        
        'Call generateReport(dtmDate1, dtmDate2, txtSearchCriteria.Text, cmbProjects.Text, cmbActivityTypes.Text)
    
    Else
        MsgBox "Date should be specified as a valid date..", vbExclamation
    End If
    
End Sub

Private Sub Form_Load()
    Dim lngIdx As Long
    Dim ff          As Long
    Dim strInput    As String
    
    For lngIdx = 1 To 31
        cmbDay1.AddItem Format(lngIdx, "00")
        'cmbDay2.AddItem Format(lngIdx, "00")
    Next lngIdx

    For lngIdx = 1 To 12
        cmbMonth1.AddItem Format(lngIdx, "00")
        'cmbMonth2.AddItem Format(lngIdx, "00")
    Next lngIdx

    For lngIdx = 2004 To Year(Now)
        cmbYear1.AddItem Format(lngIdx, "0000")
        'cmbYear2.AddItem Format(lngIdx, "0000")
    Next lngIdx

    cmbDay1.Text = Format(Day(Now), "00")
    cmbMonth1.Text = Format(Month(Now), "00")
    cmbYear1.Text = Format(Year(Now), "0000")

    'cmbDay2.Text = Format(Day(Now), "00")
    'cmbMonth2.Text = Format(Month(Now), "00")
    'cmbYear2.Text = Format(Year(Now), "0000")
    
    With flexReport
        .Rows = 2
        .FixedRows = 1
        .Cols = 3
        .FixedCols = 0
        .ColWidth(0) = .Width * 0.2
        .ColWidth(1) = .Width * 0.65
        .ColWidth(2) = .Width * 0.15
        .ScrollBars = flexScrollBarVertical
        .TextMatrix(0, 0) = "Date Started"
        .TextMatrix(0, 1) = "Activity"
        .TextMatrix(0, 2) = "Duration(min)"
    End With

End Sub


'Private Function generateReport(ByVal dtmDateStart As Date, _
'                                ByVal dtmDateEnd As Date, _
'                                ByVal strSearchCriteria As String, _
'                                ByVal strProject As String, _
'                                ByVal strActivityType As String)
    
'    Dim ffIn                    As Long
'    Dim ffOut                   As Long
'    Dim strLine                 As String
'    Dim arrTemp()               As String
'    Dim blnFitsCriteria         As Boolean
'    Dim blnExactDate            As Boolean
'    Dim dtmActivityDate         As Date
'    Dim blnLastActivityValid    As Boolean
    
'    Dim strLastActivity         As String
'    Dim strLastProject          As String
'    Dim strLastActivityType     As String
'    Dim dtmLastStartDate        As Date
'    Dim lngTotalSeconds         As Long
    
'    Dim lngHoursWorked          As Long
'    Dim lngMinutesWorked        As Long
'    Dim lngSecondsWorked        As Long
    
'    ffIn = FreeFile
    
'    If dtmDateStart = dtmDateEnd Then
'        blnExactDate = True
'    Else
'        blnExactDate = False
'    End If
    
'    If Dir(App.Path & "\activity.txt", vbNormal) <> "" Then
'        Open App.Path & "\activity.txt" For Input Access Read As ffIn
'        ffOut = FreeFile
'        Open App.Path & "\ActivityReport.htm" For Output Access Write As ffOut
'    Else
'        MsgBox "Activity File Not Found..", vbExclamation
'        Exit Function
'    End If
    
'    Print #ffOut, "<html>" & _
'                  "<head>" & _
'                  "<title>Activity Report</title>" & _
'                  "<meta http-equiv=""Content-Type"" content=""text/html; charset=windows-1254"">" & _
'                  "</head>" & _
'                  "<style>" & _
'                  "body{font-family:tahoma,verdana,arial;font-size:11px;}" & _
'                  "table{font-family:tahoma,verdana,arial;font-size:11px;}" & _
'                  "td{font-family:tahoma,verdana,arial;font-size:11px;}" & _
'                  "</style>" & _
'                  "<body>" & _
'                  "<table width=600 border=1 bordercolordark=#f0f0f0 bordercolorlight=#808080>" & _
'                  "<tr><td width=""25%""><b>PROJECT</b></td><td width=""25%""><b>ACTIVITY TYPE</b></td><td width=""35%""><b>ACTIVITY</b></td><td width=""15%""><b>TOTAL TIME</b></td></tr>"

'    Do While Not EOF(ffIn)
'        blnFitsCriteria = True
'        Line Input #ffIn, strLine
'        arrTemp = Split(strLine, "|")
'        If IsDate(arrTemp(0)) Then
'            dtmActivityDate = CDate(arrTemp(0))
            
'            ' Check if date is between the report date interval
'            If blnExactDate Then
'                If DateDiff("d", dtmActivityDate, dtmDateStart) <> 0 Then
'                    blnFitsCriteria = False
'                    If Not blnLastActivityValid Then
'                        GoTo ReadNext
'                    End If
'                End If
'            Else
'                If dtmActivityDate < dtmDateStart Or dtmActivityDate > dtmDateEnd Then
'                    blnFitsCriteria = False
'                    If Not blnLastActivityValid Then
'                        GoTo ReadNext
'                    End If
'                End If
'            End If
        
'            If UBound(arrTemp) = 3 Then
                
'                ' Check if project criteria is met
'                If strProject <> "" Then
'                    If arrTemp(1) <> strProject Then
'                        blnFitsCriteria = False
'                        If Not blnLastActivityValid Then
'                            GoTo ReadNext
'                        End If
'                    End If
'                End If
                
'                ' Check if activity type criteria is met
'                If strActivityType <> "" Then
'                    If arrTemp(2) <> strActivityType Then
'                        blnFitsCriteria = False
'                        If Not blnLastActivityValid Then
'                            GoTo ReadNext
'                        End If
'                    End If
'                End If
                
'                ' Check if search criteria is met
'                If strSearchCriteria <> "" Then
'                    If InStr(1, arrTemp(3), strSearchCriteria) <= 0 Then
'                        blnFitsCriteria = False
'                        If Not blnLastActivityValid Then
'                            GoTo ReadNext
'                        End If
'                    End If
'                End If
                
'                ' activity fits search criterias
'                If blnFitsCriteria Then
'                    If blnLastActivityValid Then
'                        ' Insert record for last activity
'                        lngTotalSeconds = DateDiff("s", dtmLastStartDate, dtmActivityDate)
'                        lngHoursWorked = Int(lngTotalSeconds / 3600)
'                        lngTotalSeconds = lngTotalSeconds - (lngHoursWorked * 3600)
'                        lngMinutesWorked = Int(lngTotalSeconds / 60)
'                        lngTotalSeconds = lngTotalSeconds - (lngMinutesWorked * 60)
'                        lngSecondsWorked = lngTotalSeconds
'                        Print #ffOut, "<tr><td>" & strLastProject & "</td><td>" & strLastActivityType & "</td><td>" & strLastActivity & "</td><td>" & Format(lngHoursWorked, "00") & ":" & Format(lngMinutesWorked, "00") & ":" & Format(lngSecondsWorked, "00") & "</td></tr>"
                
'                    End If
                    
'                    ' Set activity properties
'                    dtmLastStartDate = dtmActivityDate
'                    strLastProject = arrTemp(1)
'                    strLastActivityType = arrTemp(2)
'                    strLastActivity = arrTemp(3)
                    
'                    blnLastActivityValid = True
                
'                ElseIf blnLastActivityValid Then
'                    ' Insert record for last activity
'                    lngTotalSeconds = DateDiff("s", dtmLastStartDate, dtmActivityDate)
'                    lngHoursWorked = Int(lngTotalSeconds / 3600)
'                    lngTotalSeconds = lngTotalSeconds - (lngHoursWorked * 3600)
'                    lngMinutesWorked = Int(lngTotalSeconds / 60)
'                    lngTotalSeconds = lngTotalSeconds - (lngMinutesWorked * 60)
'                    lngSecondsWorked = lngTotalSeconds
'                    Print #ffOut, "<tr><td>" & strLastProject & "</td><td>" & strLastActivityType & "</td><td>" & strLastActivity & "</td><td>" & Format(lngHoursWorked, "00") & ":" & Format(lngMinutesWorked, "00") & ":" & Format(lngSecondsWorked, "00") & "</td></tr>"
                    
'                    blnLastActivityValid = False
                
'                End If
                
'            ElseIf blnLastActivityValid Then
'                If arrTemp(1) = "PAUSE" Then
'                    ' Store the total seconds worked before pause
'                    lngTotalSeconds = DateDiff("s", dtmLastStartDate, dtmActivityDate)
'                ElseIf arrTemp(1) = "STOP" Then
'                    ' Insert record for last activity
'                    lngTotalSeconds = DateDiff("s", dtmLastStartDate, dtmActivityDate)
'                    lngHoursWorked = Int(lngTotalSeconds / 3600)
'                    lngTotalSeconds = lngTotalSeconds - (lngHoursWorked * 3600)
'                    lngMinutesWorked = Int(lngTotalSeconds / 60)
'                    lngTotalSeconds = lngTotalSeconds - (lngMinutesWorked * 60)
'                    lngSecondsWorked = lngTotalSeconds
'                    Print #ffOut, "<tr><td>" & strLastProject & "</td><td>" & strLastActivityType & "</td><td>" & strLastActivity & "</td><td>" & Format(lngHoursWorked, "00") & ":" & Format(lngMinutesWorked, "00") & ":" & Format(lngSecondsWorked, "00") & "</td></tr>"
                
'                    blnLastActivityValid = False
                
'                ElseIf arrTemp(1) = "CONTINUE" Then
'                    ' Update the last activity start date
'                    dtmLastStartDate = DateAdd("s", -lngTotalSeconds, dtmActivityDate)
'                End If
            
'            End If
        
'        End If

'ReadNext:
    
'    Loop
    
'    Print #ffOut, "</table></body></html>"
    
'    Close

'End Function
