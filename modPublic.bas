Attribute VB_Name = "modPublic"
Option Explicit

Declare Function timeSetEvent Lib "winmm.dll" (ByVal uDelay As Long, ByVal uResolution As Long, ByVal lpFunction As Long, ByVal dwUser As Long, ByVal uFlags As Long) As Long
Declare Function timeKillEvent Lib "winmm.dll" (ByVal uId As Long) As Long

'user defined type required by Shell_NotifyIcon API call
Public Type NOTIFYICONDATA
    cbSize              As Long
    hwnd                As Long
    uId                 As Long
    uFlags              As Long
    uCallBackMessage    As Long
    hIcon               As Long
    szTip               As String * 64
End Type

'constants required by Shell_NotifyIcon API call:
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201     'Button down
Public Const WM_LBUTTONUP = &H202       'Button up
Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
Public Const WM_RBUTTONDOWN = &H204     'Button down
Public Const WM_RBUTTONUP = &H205       'Button up
Public Const WM_RBUTTONDBLCLK = &H206   'Double-click

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Public nid As NOTIFYICONDATA

'  flags for wFlags parameter of timeSetEvent() function
Public Const TIME_ONESHOT = 0  '  program timer for single event
Public Const TIME_PERIODIC = 1  '  program for continuous periodic event

Public Type Client
    UserID                      As Long
    ArrivedData                 As String
    ConnectedDate               As Date
    LastActivityStartDate       As Date
    Name                        As String
    IsConnected                 As Boolean
    TimerHandle                 As Long
    PreviousProject             As String
    PreviousActivityType        As String
    PreviousActivity            As String
    PreviousActivityStartDate   As Date
    PauseStartTime              As String
    IsOnPause                   As Boolean
    DoingNothing                As Boolean
    PingSent                    As Boolean
    ActivityID                  As Long
    PauseID                     As Long
    IsLoggedIn                  As Boolean
    LoginID                     As Long
    ByeSent                     As Boolean
End Type

Public Type Project
    ProjectID                   As Long
    ProjectName                 As String
    Visible                     As Boolean
End Type


Public Type ActivityType
    ActivityTypeID              As Long
    ActivityTypeName            As String
    Visible                     As Boolean
End Type

Public Type PauseCause
    PauseCauseID                As Long
    strPauseCause               As String
    Visible                     As Boolean
End Type
    

Public arrProjects()            As Project
Public arrActivityTypes()       As ActivityType
Public arrPauseCauses()         As PauseCause
Public arrClients()             As Client

Public blnCloseCommand          As Boolean

Public objCrypto                As New clsEncrypt
Public objData                  As New clsData

Public cnnMain                  As New ADODB.Connection
Public cnnClient                As New ADODB.Connection

Public Type OpenedWindows
    blnSQLServer            As Boolean
    blnActivityServer       As Boolean
    blnUsers                As Boolean
    blnActivityTypes        As Boolean
    blnProjects             As Boolean
    blnPauseCauses          As Boolean
    blnDailyReports         As Boolean
    


End Type


Public Type SQLServerConnectionProperties
    ServerName          As String
    UserID              As String
    Password            As String
    NTAuth              As Boolean
End Type


Public winOpen                  As OpenedWindows

Public SQLServerProperties      As SQLServerConnectionProperties

Public arrMSGLinesWritten(10)    As String
Public arrMSGColors(10)          As Long
Public lngMSGLinesWritten       As Long

Public Const MAX_LINE_COUNT = 11

Public Const COLOR_WARNING = &H104040
Public Const COLOR_EXCEPTION = &H101080
Public Const COLOR_SUCCESS = &H108010
Public Const COLOR_NORMAL = &H101010

Public Const PING_INTERVAL = 30000 ' per 30 second

Public Enum ApplicationDataTypes
    APP_DATA_TYPE_USERS = 0
    APP_DATA_TYPE_PROJECTS = 1
    APP_DATA_TYPE_ACTIVITY_TYPES = 2
    APP_DATA_TYPE_PAUSE_CAUSES = 3
End Enum

Public Enum ReportTypes
    REPORT_TYPE_SCREEN = 0
    REPORT_TYPE_TEXT_FILE = 1
    REPORT_TYPE_XML_FILE = 2
    REPORT_TYPE_PRINTER = 3
End Enum


Public Function NewMSG(ByVal strMessage As String, ByVal lngColor As Long) As Long
    
    Dim lngMsg      As Long
    
    If lngMSGLinesWritten >= MAX_LINE_COUNT Then
        ' Move Messages Up
        For lngMsg = 1 To MAX_LINE_COUNT - 1
            arrMSGLinesWritten(lngMsg - 1) = arrMSGLinesWritten(lngMsg)
            arrMSGColors(lngMsg - 1) = arrMSGColors(lngMsg)
        Next
        arrMSGLinesWritten(MAX_LINE_COUNT - 1) = strMessage
        arrMSGColors(MAX_LINE_COUNT - 1) = lngColor
    Else
        arrMSGLinesWritten(lngMSGLinesWritten) = strMessage
        arrMSGColors(lngMSGLinesWritten) = lngColor
        lngMSGLinesWritten = lngMSGLinesWritten + 1
    End If
    
    Call DisplayMessages
    
End Function

Public Function DisplayMessages()
    Dim lngMsg As Long
    With mdiMain.picProgress
        .Cls
        For lngMsg = 0 To lngMSGLinesWritten - 1
            .ForeColor = arrMSGColors(lngMsg)
            mdiMain.picProgress.Print " " & arrMSGLinesWritten(lngMsg)
        Next lngMsg
    End With
End Function

Public Function GetProjects()
    
    On Error GoTo excHandler
    
    Dim cnnProjects     As New ADODB.Connection
    Dim cmdProjects     As New ADODB.Command
    Dim rstProjects     As New ADODB.Recordset
    Dim lngCount        As Long
    
    SetSQLServerProperties cnnProjects, True
    
    With cmdProjects
        .CommandType = adCmdStoredProc
        .CommandText = "prGetProjects"
        .ActiveConnection = cnnProjects
        rstProjects.Open cmdProjects
    End With
    
    Do While Not rstProjects.EOF
        If lngCount = 0 Then
            ReDim arrProjects(lngCount)
        Else
            ReDim Preserve arrProjects(lngCount)
        End If
        arrProjects(lngCount).ProjectID = rstProjects("intProjectID")
        arrProjects(lngCount).ProjectName = rstProjects("strProjectName")
        arrProjects(lngCount).Visible = rstProjects("bitVisible")
        lngCount = lngCount + 1
        rstProjects.MoveNext
    Loop
    
ExitFunction:

    Set rstProjects = Nothing
    Set cmdProjects = Nothing
    Set cnnProjects = Nothing
    
    Exit Function

excHandler:

    LogException "GetProjects", Err.Number, Err.Description, "Could not get projects", True

    Resume ExitFunction
    
End Function

Public Function GetActivityTypes()
    
    On Error GoTo excHandler
    
    Dim cnnActivityTypes        As New ADODB.Connection
    Dim cmdActivityTypes        As New ADODB.Command
    Dim rstActivityTypes        As New ADODB.Recordset
    Dim lngCount                As Long
    
    SetSQLServerProperties cnnActivityTypes, True
    
    With cmdActivityTypes
        .CommandType = adCmdStoredProc
        .CommandText = "prGetActivityTypes"
        .ActiveConnection = cnnActivityTypes
        rstActivityTypes.Open cmdActivityTypes
    End With
    
    Do While Not rstActivityTypes.EOF
        If lngCount = 0 Then
            ReDim arrActivityTypes(lngCount)
        Else
            ReDim Preserve arrActivityTypes(lngCount)
        End If
        arrActivityTypes(lngCount).ActivityTypeID = rstActivityTypes("intActivityTypeID")
        arrActivityTypes(lngCount).ActivityTypeName = rstActivityTypes("strActivityTypeName")
        arrActivityTypes(lngCount).Visible = rstActivityTypes("bitVisible")
        lngCount = lngCount + 1
        rstActivityTypes.MoveNext
    Loop
    
ExitFunction:

    Set rstActivityTypes = Nothing
    Set cmdActivityTypes = Nothing
    Set cnnActivityTypes = Nothing
    
    Exit Function

excHandler:

    LogException "GetActivityTypes", Err.Number, Err.Description, "Could not get activity types", True

    Resume ExitFunction

End Function

Public Function GetPauseCauses()
    
    On Error GoTo excHandler
    
    Dim cnnPauseCauses          As New ADODB.Connection
    Dim cmdPauseCauses          As New ADODB.Command
    Dim rstPauseCauses          As New ADODB.Recordset
    Dim lngCount                As Long
    
    SetSQLServerProperties cnnPauseCauses, True
    
    With cmdPauseCauses
        .CommandType = adCmdStoredProc
        .CommandText = "prGetPauseCauses"
        .ActiveConnection = cnnPauseCauses
        rstPauseCauses.Open cmdPauseCauses
    End With
    
    Do While Not rstPauseCauses.EOF
        If lngCount = 0 Then
            ReDim arrPauseCauses(lngCount)
        Else
            ReDim Preserve arrPauseCauses(lngCount)
        End If
        arrPauseCauses(lngCount).PauseCauseID = rstPauseCauses("intPauseCauseID")
        arrPauseCauses(lngCount).strPauseCause = rstPauseCauses("strPauseCause")
        arrPauseCauses(lngCount).Visible = rstPauseCauses("bitVisible")
        lngCount = lngCount + 1
        rstPauseCauses.MoveNext
    Loop
    
ExitFunction:

    Set rstPauseCauses = Nothing
    Set cmdPauseCauses = Nothing
    Set cnnPauseCauses = Nothing
    
    Exit Function

excHandler:

    LogException "GetPauseCauses", Err.Number, Err.Description, "Could not get pause causes", True

    Resume ExitFunction

End Function

Public Function DisplayInformation(ByVal strCustomText As String, ByVal lngError As Long, ByVal strSource As String, ByVal strMessage As String)
    
    NewMSG strCustomText & ", " & _
           "Error Number:  " & lngError & ", " & _
           "Error Source:  " & strSource & ", " & _
           "Error Message: " & strMessage, _
           COLOR_EXCEPTION

End Function

Public Function LogException(ByVal strProcedure As String, ByVal lngExceptionCode As Long, ByVal strExceptionMessage As String, ByVal strExtraInfo As String, ByVal blnIsRunTimeException As Boolean) As Long
    Dim cnnException    As New ADODB.Connection
    Dim cmdException    As New ADODB.Command
    
    On Error Resume Next
    
    NewMSG "Exception occured in procedure [" & strProcedure & "], number [" & lngExceptionCode & "], message [" & strExceptionMessage & "], extra info [" & strExtraInfo & "]", COLOR_EXCEPTION
    
    SetSQLServerProperties cnnException, True
    
    With cmdException
        .CommandType = adCmdStoredProc
        .CommandText = "prLogException"
        .ActiveConnection = cnnException
        .Parameters.Refresh
        
        .Parameters("@strProcedure").Value = strProcedure
        .Parameters("@intExceptionCode").Value = lngExceptionCode
        .Parameters("@strExceptionMessage").Value = strExceptionMessage
        .Parameters("@strExtraInfo").Value = strExtraInfo
        If blnIsRunTimeException Then
            .Parameters("@bitIsRunTimeException").Value = 1
        Else
            .Parameters("@bitIsRunTimeException").Value = 0
        End If
    
        .Execute

    End With

    Set cmdException = Nothing
    Set cnnException = Nothing

End Function


Public Function PutSQLServerProperties()
    
    With SQLServerProperties
        .ServerName = objCrypto.Encrypt(frmSqlServer.txtSQLServer.Text)
        .NTAuth = frmSqlServer.optNTAuthentication.Value = True
        If .NTAuth Then
            .UserID = ""
            .Password = ""
        Else
            .UserID = objCrypto.Encrypt(frmSqlServer.txtUserID.Text)
            .Password = objCrypto.Encrypt(frmSqlServer.txtPassword.Text)
        End If
    End With

End Function

Public Function SetSQLServerProperties(ByRef cnnDB As ADODB.Connection, ByVal blnUseAPPDB As Boolean) As Long

    With cnnDB
        If .State = adStateClosed Then
            .Provider = "SQLOLEDB.1"
            .Properties.Item("Data Source").Value = objCrypto.Decrypt(SQLServerProperties.ServerName)
            .Properties.Item("Initial catalog").Value = "master"
            If SQLServerProperties.NTAuth Then
                .Properties.Item("Integrated Security").Value = "SSPI"
            Else
                .Properties.Item("User ID").Value = objCrypto.Decrypt(SQLServerProperties.UserID)
                .Properties.Item("Password").Value = objCrypto.Decrypt(SQLServerProperties.Password)
            End If
            .CursorLocation = adUseClient
        End If
    End With
    
    Screen.MousePointer = vbHourglass
    
    If cnnDB.State = adStateClosed Then
        cnnDB.Open
    End If
    
    If blnUseAPPDB Then
        cnnDB.Execute "USE dbActivities"
    End If
    
    Screen.MousePointer = vbNormal
    
End Function
    

Public Function GetProjectName(ByVal lngProjectID As Long) As String
    Dim lngIdx As Long
    For lngIdx = 0 To UBound(arrProjects)
        If arrProjects(lngIdx).ProjectID = lngProjectID Then
            GetProjectName = arrProjects(lngIdx).ProjectName
            Exit Function
        End If
    Next lngIdx
End Function

Public Function GetActivityTypeName(ByVal lngActivityTypeID As Long) As String
    Dim lngIdx As Long
    For lngIdx = 0 To UBound(arrActivityTypes)
        If arrActivityTypes(lngIdx).ActivityTypeID = lngActivityTypeID Then
            GetActivityTypeName = arrActivityTypes(lngIdx).ActivityTypeName
            Exit Function
        End If
    Next lngIdx
End Function

Public Function GetPauseCauseName(ByVal lngPauseCauseID As Long) As String
    Dim lngIdx As Long
    For lngIdx = 0 To UBound(arrPauseCauses)
        If arrPauseCauses(lngIdx).PauseCauseID = lngPauseCauseID Then
            GetPauseCauseName = arrPauseCauses(lngIdx).strPauseCause
            Exit Function
        End If
    Next lngIdx
End Function


Public Function CreatePrintReport(ByRef rstReportData As ADODB.Recordset) As Boolean

    Dim rstUsers            As New ADODB.Recordset
    Dim rstProjects         As New ADODB.Recordset
    Dim rstActivityTypes    As New ADODB.Recordset
    
    If rstReportData.EOF Then
        Exit Function
    End If
    
    Call objData.GetApplicationData(rstUsers, APP_DATA_TYPE_USERS)
    Call objData.GetApplicationData(rstActivityTypes, APP_DATA_TYPE_ACTIVITY_TYPES)
    Call objData.GetApplicationData(rstProjects, APP_DATA_TYPE_PROJECTS)

    rstUsers.Filter = "intUserID = " & rstReportData("intUserID").Value
    
    Dim lngPrinterW     As Long
    Dim lngPrinterH     As Long
    
    lngPrinterW = Printer.ScaleWidth
    lngPrinterH = Printer.ScaleHeight
    
    Printer.DrawWidth = 2
    
    ' Main borders
    Printer.Line (lngPrinterW * 0.05, lngPrinterH * 0.05)-(lngPrinterW * 0.9, lngPrinterH * 0.05), vbBlack
    Printer.Line (lngPrinterW * 0.05, lngPrinterH * 0.05)-(lngPrinterW * 0.05, lngPrinterH * 0.9), vbBlack
    Printer.Line (lngPrinterW * 0.9, lngPrinterH * 0.05)-(lngPrinterW * 0.9, lngPrinterH * 0.9), vbBlack
    
    ' Header bottom borders
    Printer.Line (lngPrinterW * 0.05, lngPrinterH * 0.075)-(lngPrinterW * 0.9, lngPrinterH * 0.075), vbBlack
    Printer.Line (lngPrinterW * 0.05, lngPrinterH * 0.1)-(lngPrinterW * 0.9, lngPrinterH * 0.1), vbBlack
    
    ' Add Main Info
    Printer.FontBold = True
    Printer.FontSize = 8
    Printer.CurrentX = lngPrinterW * 0.06
    Printer.CurrentY = lngPrinterH * 0.05 + (lngPrinterH * 0.025 - Printer.TextHeight("A")) / 2
    Printer.Print "User: " & rstUsers("strUserFirstName") & " " & rstUsers("strUserLastName") & ", Date: " & Format(rstReportData("dtmActivityStartDate"), "dd.mm.yyyy")
    
    Printer.FontBold = True
    Printer.CurrentX = lngPrinterW * 0.06
    Printer.CurrentY = lngPrinterH * 0.075 + (lngPrinterH * 0.025 - Printer.TextHeight("P")) / 2
    Printer.Print "Project"
    
    ' Vertical Borders
    Printer.Line (lngPrinterW * 0.2, lngPrinterH * 0.075)-(lngPrinterW * 0.2, lngPrinterH * 0.9), vbBlack
    
    Printer.FontBold = True
    Printer.CurrentX = lngPrinterW * 0.21
    Printer.CurrentY = lngPrinterH * 0.075 + (lngPrinterH * 0.025 - Printer.TextHeight("P")) / 2
    Printer.Print "Activity Type"
    
    Printer.Line (lngPrinterW * 0.35, lngPrinterH * 0.075)-(lngPrinterW * 0.35, lngPrinterH * 0.9), vbBlack
    
    Printer.FontBold = True
    Printer.CurrentX = lngPrinterW * 0.36
    Printer.CurrentY = lngPrinterH * 0.075 + (lngPrinterH * 0.025 - Printer.TextHeight("P")) / 2
    Printer.Print "Activity"
    
    Printer.Line (lngPrinterW * 0.7, lngPrinterH * 0.075)-(lngPrinterW * 0.7, lngPrinterH * 0.9), vbBlack
    
    Printer.FontBold = True
    Printer.CurrentX = lngPrinterW * 0.71
    Printer.CurrentY = lngPrinterH * 0.075 + (lngPrinterH * 0.025 - Printer.TextHeight("P")) / 2
    Printer.Print "Start Time"
    
    Printer.Line (lngPrinterW * 0.8, lngPrinterH * 0.075)-(lngPrinterW * 0.8, lngPrinterH * 0.9), vbBlack
    
    Printer.FontBold = True
    Printer.CurrentX = lngPrinterW * 0.81
    Printer.CurrentY = lngPrinterH * 0.075 + (lngPrinterH * 0.025 - Printer.TextHeight("P")) / 2
    Printer.Print "Duration"
    
    
    Dim lngID       As Long
    Dim lngY        As Long
    Dim lngHour     As Long
    Dim lngMin      As Long
    Dim lngSec      As Long
    Dim arrText()   As String
    Dim lngIdx      As Long
    
    lngID = 0
    
    Do While Not rstReportData.EOF
        
        Printer.FontBold = False
        
        lngY = lngPrinterH * (0.1 + lngID * 0.035) + (lngPrinterH * 0.035 - Printer.TextHeight("P")) / 2
        
        rstProjects.Filter = "intProjectID = " & rstReportData("intProjectID")
        
        arrText = ParseStringToPrinterWidth(lngPrinterW * 0.15, rstProjects("strProjectName"))
        
        For lngIdx = 0 To UBound(arrText)
            
            Printer.CurrentX = lngPrinterW * 0.06
            
            Printer.CurrentY = lngY + lngIdx * (Printer.TextHeight("A"))
            Printer.Print arrText(lngIdx)
            
        Next lngIdx
        
        'Printer.Print rstProjects("strProjectName")
    
        rstActivityTypes.Filter = "intActivityTypeID = " & rstReportData("intActivityTypeID")
        
        arrText = ParseStringToPrinterWidth(lngPrinterW * 0.15, rstActivityTypes("strActivityTypeName"))
        
        For lngIdx = 0 To UBound(arrText)
            Printer.CurrentX = lngPrinterW * 0.21
            Printer.CurrentY = lngY + lngIdx * (Printer.TextHeight("A"))
            Printer.Print arrText(lngIdx)
        Next lngIdx
    
        arrText = ParseStringToPrinterWidth(lngPrinterW * 0.35, rstReportData("strActivity"))
    
        For lngIdx = 0 To UBound(arrText)
            Printer.CurrentX = lngPrinterW * 0.36
            Printer.CurrentY = lngY + lngIdx * (Printer.TextHeight("A"))
            Printer.Print arrText(lngIdx)
        Next lngIdx
        
        Printer.CurrentX = lngPrinterW * 0.71
        Printer.CurrentY = lngY
        Printer.Print Format(rstReportData("dtmActivityStartDate"), "hh:nn:ss")
    
        Printer.CurrentX = lngPrinterW * 0.81
        Printer.CurrentY = lngY
        
        lngSec = rstReportData("intActivityTotalTimeSec")
        lngHour = Int(lngSec / 3600)
        lngSec = lngSec - lngHour * 3600
        lngMin = Int(lngSec / 60)
        lngSec = lngSec - lngMin * 60
        
        Printer.Print Right("00" & lngHour, 2) & " hr. " & Right("00" & lngMin, 2) & " min."
    
        lngID = lngID + 1
        
        rstReportData.MoveNext
            
    Loop
    
    
    Printer.EndDoc
    
    
    
    
    ' For each page, draw the vertical lines first
    
    






    CreatePrintReport = True

End Function



Private Function ParseStringToPrinterWidth(ByVal lngPrintAreaWidth As Long, ByVal strtext As String) As String()

    Dim strLine         As String
    Dim arrResult()     As String
    Dim lngIdx          As Long
    Dim lngSpacePos     As Long
    Dim strtemp         As String
    
    ReDim arrResult(0)
    lngIdx = 0
    
    strtext = Trim(strtext)
    
    lngSpacePos = InStr(1, strtext, " ")
    
    If lngSpacePos = 0 Then
        arrResult(0) = strtext
        ParseStringToPrinterWidth = arrResult
        Exit Function
    End If
    
    strLine = ""
    
    Do
        
        strLine = ""
        
        lngSpacePos = InStr(1, strtext, " ")
        If lngSpacePos > 0 Then
            strtemp = strLine & Left$(strtext, lngSpacePos - 1)
        Else
            strtemp = strtext
        End If
        
        Do While Printer.TextWidth(strtemp) < lngPrintAreaWidth And lngSpacePos > 0
            
            strLine = strLine & Left$(strtext, lngSpacePos) ' Include space
            strtext = Mid(strtext, lngSpacePos + 1)
            lngSpacePos = InStr(1, strtext, " ")
            
            If lngSpacePos > 0 Then
                strtemp = strLine & Left$(strtext, lngSpacePos - 1)
            Else
                strtemp = strLine & strtext
            End If
        
        Loop
            
        If lngIdx > 0 Then
            ReDim Preserve arrResult(lngIdx)
        End If
        
        If strLine <> "" Then
            arrResult(lngIdx) = Trim(strLine)
        Else
            arrResult(lngIdx) = strtemp
            strtext = ""
        End If
        
        lngIdx = lngIdx + 1
    
    Loop Until Len(strtext) = 0
    
    ParseStringToPrinterWidth = arrResult

End Function

Public Function CreateXMLReport(ByRef rstReportData As ADODB.Recordset) As Long
    
    Dim strResultXML    As String
    
    Const XML_VERSION As String = "<?xml version=""1.0"" encoding=""ISO-8859-9""?>"
    
    Const XML_TAG_REPORT        As String = "<REPORT DATE_CREATED=""#DATE#"">" & _
                                                "#DATA#" & _
                                            "</REPORT>"
    
    Const XML_TAG_USER          As String = "<USER ID=""#USER_ID#"" NAME=""#USER_NAME#"" TITLE=""#TITLE#"">" & _
                                                "#DATA#" & _
                                            "</USER>"
    
    Const XML_ACTIVITY_TEMPLATE As String = "<ACTIVITY ID=""#ACTIVITY_ID#"">" & _
                                                "<START_TIME>#START_TIME#</START_TIME>" & _
                                                "<PROJECT_NAME>#PROJECT_NAME#</PROJECT_NAME>" & _
                                                "<ACTIVITY_TYPE>#ACTIVITY_TYPE#</ACTIVITY_TYPE>" & _
                                                "<ACTIVITY_DETAIL>#ACTIVITY_DETAIL#</ACTIVITY_DETAIL>" & _
                                            "</ACTIVITY>"
    
    Const XML_PAUSE_TEMPLATE    As String = "<PAUSE ID=""#PAUSE_ID#"" ACTIVITY_ID=""#ACTIVITY_ID#"">" & _
                                                "<START_TIME>#START_TIME#</START_TIME>" & _
                                                "<PAUSE_TYPE>#PAUSE_TYPE#</PAUSE_TYPE>" & _
                                                "<PAUSE_DETAIL>#PAUSE_DETAIL#</PAUSE_DETAIL>" & _
                                            "</PAUSE>"
    



















End Function


