VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim cmdData     As New ADODB.Command
Dim cnnData     As New ADODB.Connection

Function CheckPassword(ByVal strUserName As String, ByVal strPassword As String) As Long

    On Error GoTo excHandler
    
    SetSQLServerProperties cnnData, True
    
    With cmdData
        .CommandType = adCmdStoredProc
        .CommandText = "prCheckPassword"
        .ActiveConnection = cnnData
        .Parameters.Refresh
    
        .Parameters("@strUserName").Value = strUserName
        .Parameters("@strPassword").Value = strPassword
        
        .Execute
        
        If Not IsNull(.Parameters("@intUserID")) Then
            CheckPassword = .Parameters("@intUserID")
        Else
            CheckPassword = 0
        End If
    End With

ExitFunction:

    Set cmdData = Nothing
    Set cnnData = Nothing

    Exit Function
    
excHandler:

    LogException "CheckPassword", Err.Number, Err.Description, "User name: " & strUserName, True
    
    CheckPassword = -1
    
    Resume ExitFunction

End Function


Public Function GetUser(ByVal lngUserID As Long) As String
    
    Dim cnnUser     As New ADODB.Connection
    Dim cmdUser     As New ADODB.Command
    Dim rstUser     As New ADODB.Recordset
    
    On Error GoTo excHandler
    
    SetSQLServerProperties cnnUser, True
        
    With cmdUser
        .CommandType = adCmdStoredProc
        .CommandText = "prGetUser"
        .ActiveConnection = cnnUser
        .Parameters.Refresh
        .Parameters("@intUserID").Value = lngUserID
        Set rstUser = .Execute
    End With
    
    If rstUser.RecordCount = 1 Then
        GetUser = rstUser("strUserFirstName") & " " & rstUser("strUserLastName")
    Else
        GetUser = ""
    End If

ExitFunction:

    Set cmdUser = Nothing
    Set rstUser = Nothing
    Set cnnUser = Nothing

    Exit Function

excHandler:

    LogException "GetUser", Err.Number, Err.Description, "User ID: " & lngUserID, True
    Resume ExitFunction

End Function


Public Function InsertActivity(ByVal intUserID As Long, _
                               ByVal intProjectID As Long, _
                               ByVal intActivityTypeID As Long, _
                               ByVal strActivity As String) As Long

    On Error GoTo excHandler
    
    Dim cnnActivity As New ADODB.Connection
    Dim cmdActivity As New ADODB.Command
    
    SetSQLServerProperties cnnActivity, True
    
    With cmdActivity
        .CommandType = adCmdStoredProc
        .CommandText = "prInsertActivity"
        .ActiveConnection = cnnActivity
        .Parameters.Refresh
        .Parameters("@intUserID").Value = intUserID
        .Parameters("@intProjectID").Value = intProjectID
        .Parameters("@intActivityTypeID").Value = intActivityTypeID
        .Parameters("@strActivity").Value = strActivity
        .Execute
        If IsNull(.Parameters("@intActivityID").Value) Then
            InsertActivity = 0
        Else
            InsertActivity = .Parameters("@intActivityID").Value
        End If
    End With
        
ExitFunction:
    
    Set cmdActivity = Nothing
    Set cnnActivity = Nothing

    Exit Function

excHandler:
    
    LogException "InsertActivity", Err.Number, Err.Description, "User ID: " & intUserID & ", Project ID: " & intProjectID & ", Activity Type ID: " & intActivityTypeID & ", Date: " & Now & ", The activity is: " & vbNewLine & strActivity, True

    Resume ExitFunction

End Function
        

Public Function UpdateActivity(ByVal intActivityID As Long, _
                               ByVal intActivityTotalTimeSec As Long) As Long

    On Error GoTo excHandler
    
    Dim cnnActivity As New ADODB.Connection
    Dim cmdActivity As New ADODB.Command
    Dim lngRecs  As Long
    
    SetSQLServerProperties cnnActivity, True
    
    With cmdActivity
        .CommandType = adCmdStoredProc
        .CommandText = "prUpdateActivity"
        .ActiveConnection = cnnActivity
        .Parameters.Refresh
        .Parameters("@intActivityID").Value = intActivityID
        .Parameters("@intActivityTotalTimeSec").Value = intActivityTotalTimeSec
        .Execute lngRecs
    End With
        
ExitFunction:
    
    Set cmdActivity = Nothing
    Set cnnActivity = Nothing

    Exit Function

excHandler:
    
    LogException "UpdateActivity", Err.Number, Err.Description, "Activity ID: " & intActivityID & ", Date: " & Now & ", The activity time total in seconds: " & intActivityTotalTimeSec, True

    Resume ExitFunction

End Function



Public Function InsertPause(ByVal intUserID As Long, _
                            ByVal intActivityID As Long, _
                            ByVal intPauseCauseID As Long, _
                            ByVal strPauseCauseDetail As String) As Long

    On Error GoTo excHandler
    
    Dim cnnPause As New ADODB.Connection
    Dim cmdPause As New ADODB.Command
    
    SetSQLServerProperties cnnPause, True
    
    With cmdPause
        .CommandType = adCmdStoredProc
        .CommandText = "prInsertPause"
        .ActiveConnection = cnnPause
        .Parameters.Refresh
        .Parameters("@intUserID").Value = intUserID
        .Parameters("@intActivityID").Value = intActivityID
        .Parameters("@intPauseCauseID").Value = intPauseCauseID
        .Parameters("@strPauseCauseDetail").Value = strPauseCauseDetail
        .Execute
        If IsNull(.Parameters("@intPauseID").Value) Then
            InsertPause = 0
        Else
            InsertPause = .Parameters("@intPauseID").Value
        End If
    End With
        
ExitFunction:
    
    Set cmdPause = Nothing
    Set cnnPause = Nothing

    Exit Function

excHandler:
    
    LogException "InsertPause", Err.Number, Err.Description, "User ID: " & intUserID & ", Activity ID: " & intActivityID & ", Date: " & Now & ", The pause cause is: " & vbNewLine & strPauseCauseDetail, True

    Resume ExitFunction

End Function

Public Function UpdatePause(ByVal intPauseID As Long, _
                            ByVal intPauseTotalTimeSec As Long) As Long

    On Error GoTo excHandler
    
    Dim cnnPause As New ADODB.Connection
    Dim cmdPause As New ADODB.Command
    Dim lngRecs  As Long
    
    SetSQLServerProperties cnnPause, True
    
    With cmdPause
        .CommandType = adCmdStoredProc
        .CommandText = "prUpdatePause"
        .ActiveConnection = cnnPause
        .Parameters.Refresh
        .Parameters("@intPauseID").Value = intPauseID
        .Parameters("@intPauseTotalTimeSec").Value = intPauseTotalTimeSec
        .Execute lngRecs
    End With
        
ExitFunction:
    
    Set cmdPause = Nothing
    Set cnnPause = Nothing

    Exit Function

excHandler:
    
    LogException "UpdatePause", Err.Number, Err.Description, "Pause ID: " & intPauseID & ", Date: " & Now & ", The pause time total in seconds: " & intPauseTotalTimeSec, True

    Resume ExitFunction

End Function

Public Function InsertLogin(ByVal strUserName As String, _
                            ByVal blnPassCorrect As Boolean, _
                            ByVal strRemoteIP As String) As Long

    On Error GoTo excHandler
    
    Dim cnnLogin As New ADODB.Connection
    Dim cmdLogin As New ADODB.Command
    
    SetSQLServerProperties cnnLogin, True
    
    With cmdLogin
        .CommandType = adCmdStoredProc
        .CommandText = "prInsertLogin"
        .ActiveConnection = cnnLogin
        .Parameters.Refresh
        .Parameters("@strUserName").Value = strUserName
        .Parameters("@bitPassCorrect").Value = CLng(IIf(blnPassCorrect, 1, 0))
        .Parameters("@strRemoteIP").Value = strRemoteIP
        .Execute
        If IsNull(.Parameters("@intLoginID").Value) Then
            InsertLogin = 0
        Else
            InsertLogin = .Parameters("@intLoginID").Value
        End If
    End With
        
ExitFunction:
    
    Set cmdLogin = Nothing
    Set cnnLogin = Nothing

    Exit Function

excHandler:
    
    LogException "InsertLogin", Err.Number, Err.Description, "User Name: " & strUserName & ", Pass is Correct: " & blnPassCorrect & ", Date: " & Now & ", The Remote IP is: " & strRemoteIP, True

    Resume ExitFunction

End Function


Public Function InsertLogOut(ByVal intUserID As Long, _
                             ByVal intLoginID As Long, _
                             ByVal blnNormalLogOut As Boolean) As Long

    On Error GoTo excHandler
    
    Dim cnnLogOut As New ADODB.Connection
    Dim cmdLogOut As New ADODB.Command
    
    SetSQLServerProperties cnnLogOut, True
    
    With cmdLogOut
        .CommandType = adCmdStoredProc
        .CommandText = "prInsertLogOut"
        .ActiveConnection = cnnLogOut
        .Parameters.Refresh
        .Parameters("@intUserID").Value = intUserID
        .Parameters("@intLoginID").Value = intLoginID
        .Parameters("@bitNormalLogOut").Value = CLng(IIf(blnNormalLogOut, 1, 0))
        .Execute
    End With
        
ExitFunction:
    
    Set cmdLogOut = Nothing
    Set cnnLogOut = Nothing

    Exit Function

excHandler:
    
    LogException "InsertLogOut", Err.Number, Err.Description, "User ID: " & intUserID & ", Login ID: " & intLoginID & ", LogOut was normal: " & blnNormalLogOut & ", Date: " & Now, True

    Resume ExitFunction

End Function

Public Function GetApplicationData(ByRef rstResult As ADODB.Recordset, ByVal lngDataType As ApplicationDataTypes) As Long
    
    On Error GoTo excHandler
    
    Dim cnnAppData As New ADODB.Connection
    Dim cmdAppData As New ADODB.Command
    
    SetSQLServerProperties cnnAppData, True
    
    With cmdAppData
        .CommandType = adCmdStoredProc
        .ActiveConnection = cnnAppData
        Select Case lngDataType
            Case APP_DATA_TYPE_USERS
                .CommandText = "prGetUsers"
            Case APP_DATA_TYPE_PROJECTS
                .CommandText = "prGetProjects"
            Case APP_DATA_TYPE_ACTIVITY_TYPES
                .CommandText = "prGetActivityTypes"
            Case APP_DATA_TYPE_PAUSE_CAUSES
                .CommandText = "prGetPauseCauses"
        End Select
        Set rstResult = .Execute
        Set rstResult.ActiveConnection = Nothing
    End With
        
ExitFunction:
    
    Set cmdAppData = Nothing
    Set cnnAppData = Nothing

    Exit Function

excHandler:
    
    LogException "GetApplicationData", Err.Number, Err.Description, "", True

    Resume ExitFunction

End Function



Public Function GetActivityReports(ByVal dtmAcitivityDayStart As Date, _
                                   ByVal dtmAcitivityDayEnd As Date, _
                                   Optional ByVal intUserID As Long = 0) As String

    On Error GoTo excHandler
    
    Dim cnnDailyReport  As New ADODB.Connection
    Dim cmdDailyReport  As New ADODB.Command
    Dim rstResult       As New ADODB.Recordset
    Dim lngRecs         As Long
    
    SetSQLServerProperties cnnDailyReport, True
    
    With cmdDailyReport
        .CommandType = adCmdStoredProc
        .ActiveConnection = cnnDailyReport
        If intUserID > 0 Then
            .CommandText = "prReportsGetUserDailyActivityReport"
        Else
            .CommandText = "prReportsGetAllDailyActivityReports"
        End If
        .Parameters.Refresh
        .Parameters("@dtmDayStart").Value = Format(dtmAcitivityDayStart, "yyyymmdd") & " 00:00:01"
        .Parameters("@dtmDayEnd").Value = Format(dtmAcitivityDayEnd, "yyyymmdd") & " 23:59:59"
        If intUserID > 0 Then
            .Parameters("@intUserID").Value = intUserID
        End If
        Set rstResult = .Execute(lngRecs)
        Set rstResult.ActiveConnection = Nothing
        
        'If lngRecs > 0 Then
        If Not (rstResult.EOF And rstResult.BOF) Then
            GetActivityReports = rstResult.GetString(StringFormatEnum.adClipString, , Chr(2), Chr(1))
        End If
        'End If
        
    End With
        
ExitFunction:
    
    Set cmdDailyReport = Nothing
    Set cnnDailyReport = Nothing

    Exit Function

excHandler:
    
    LogException "GetActivityReports", Err.Number, Err.Description, "Activity Day: " & dtmAcitivityDayStart & " - " & dtmAcitivityDayEnd & ", User ID: " & intUserID, True

    Resume ExitFunction

End Function


Public Function GetPauseReports(ByRef rstResult As ADODB.Recordset, _
                                ByVal dtmPauseDay As Date, _
                                Optional ByVal intUserID As Long = 0) As Long

    On Error GoTo excHandler
    
    Dim cnnDailyReport As New ADODB.Connection
    Dim cmdDailyReport As New ADODB.Command
    
    SetSQLServerProperties cnnDailyReport, True
    
    With cmdDailyReport
        .CommandType = adCmdStoredProc
        .ActiveConnection = cnnDailyReport
        If intUserID > 0 Then
            .CommandText = "prReportsGetUserDailyPauseReport"
        Else
            .CommandText = "prReportsGetAllDailyPauseReports"
        End If
        .Parameters.Refresh
        .Parameters("@dtmDay").Value = Format(dtmPauseDay, "yyyymmdd")
        If intUserID > 0 Then
            .Parameters("@intUserID").Value = intUserID
        End If
        Set rstResult = .Execute
        Set rstResult.ActiveConnection = Nothing
    End With
        
ExitFunction:
    
    Set cmdDailyReport = Nothing
    Set cnnDailyReport = Nothing

    Exit Function

excHandler:
    
    LogException "GetPauseReports", Err.Number, Err.Description, "Activity Day: " & dtmPauseDay & ", User ID: " & intUserID, True

    Resume ExitFunction

End Function



