VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEncrypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Const SALT_LENGTH   As Long = 4
Private Const SALT_STRING   As String = "0123456789ABCDEF"

Function Encrypt(ByVal strText As String) As String

    Dim strKey          As String
    Dim strResult       As String
    Dim lngBitCount     As Long
    Dim strSalt         As String
    
    strText = strText & "OK"
    
    strSalt = CreateSalt()
    
    strKey = CreateKey(Len(strSalt) + Len(strText))
    
    strText = ConvertToBits(strSalt & strText)

    Encrypt = strSalt & MixTextAndKey(strText, strKey)
    
End Function

Function Decrypt(ByVal strText As String) As String
    
    Dim lngBitCount     As Long
    Dim strBits         As String
    Dim strTextLocal    As String
    Dim strKey          As String
    Dim strSalt         As String
    Dim strResult       As String
    
    strSalt = Left(strText, SALT_LENGTH)
    
    strTextLocal = Mid(strText, SALT_LENGTH + 1)
    
    lngBitCount = GetBitCount(strTextLocal)
    
    strBits = ConvertHexToBits(strTextLocal)
    
    strKey = CreateKey(Len(strBits) / 8)
    
    strResult = RemixTextAndKey(strBits, strKey, lngBitCount)
    
    If Left(strResult, SALT_LENGTH) = strSalt Then
        Decrypt = Mid(strResult, SALT_LENGTH + 1)
    Else
        Err.Raise 1, "Decrypt", "CRC Error"
    End If
        
End Function

Function CreateSalt() As String
    Dim lngIdx              As Long
    Dim strResult           As String
    Randomize Timer
    For lngIdx = 1 To SALT_LENGTH
        strResult = strResult & Mid(SALT_STRING, 1 + Int(Rnd() * (Len(SALT_STRING) - 0.0001)), 1)
    Next lngIdx
    CreateSalt = strResult
End Function


Private Function ConvertToBits(ByVal strText As String) As String

    Dim lngPos      As Long
    Dim strResult   As String
    Dim bytChar     As Byte
    
    For lngPos = 1 To Len(strText)
        bytChar = Asc(Mid(strText, lngPos, 1))
        strResult = strResult & (bytChar And 128) / 128
        strResult = strResult & (bytChar And 64) / 64
        strResult = strResult & (bytChar And 32) / 32
        strResult = strResult & (bytChar And 16) / 16
        strResult = strResult & (bytChar And 8) / 8
        strResult = strResult & (bytChar And 4) / 4
        strResult = strResult & (bytChar And 2) / 2
        strResult = strResult & (bytChar And 1)
    Next lngPos
    
    ConvertToBits = strResult

End Function


Private Function ConvertToString(ByVal strBits As String) As String
    
    Dim lngPos      As Long
    Dim strResult   As String
    Dim strChar1    As String * 8
    Dim strchar2    As String * 8
    Dim byt1        As Byte
    Dim byt2        As Byte

    For lngPos = 1 To Len(strBits) Step 16
        strChar1 = Mid(strBits, lngPos, 8)
        strchar2 = Mid(strBits, lngPos + 8, 8)
        byt1 = ConvertToByte(strChar1)
        byt2 = ConvertToByte(strchar2)
        strResult = strResult & ChrW(byt1 + byt2 * 256)
    Next lngPos

    ConvertToString = strResult


End Function


Private Function ConvertToByte(ByVal strChar As String) As Byte

    Dim bytResult As Byte

    bytResult = bytResult + Mid(strChar, 1, 1) * 128
    bytResult = bytResult + Mid(strChar, 2, 1) * 64
    bytResult = bytResult + Mid(strChar, 3, 1) * 32
    bytResult = bytResult + Mid(strChar, 4, 1) * 16
    bytResult = bytResult + Mid(strChar, 5, 1) * 8
    bytResult = bytResult + Mid(strChar, 6, 1) * 4
    bytResult = bytResult + Mid(strChar, 7, 1) * 2
    bytResult = bytResult + Mid(strChar, 8, 1)

    ConvertToByte = bytResult
    
End Function

Private Function CreateKey(ByVal lngLength As Long) As String

    Dim strResult   As String
    Dim lngPos      As Long
    Dim strKeys     As String * 128
    Dim strKey      As String
    
    Randomize Timer
    
    'Call GetComputerName(strKeys, 128)
    'strKey = Left(strKeys, InStr(1, strKeys, Chr(0)) - 1)
    'strKeys = String(128, Chr(0))
    'Call GetUserName(strKeys, 128)
    'strKey = strKey & Left(strKeys, InStr(1, strKeys, Chr(0)) - 1)
    'strKeys = String(128, Chr(0))
    'Call GetSystemDirectory(strKeys, 128)
    'strKey = strKey & Left(strKeys, InStr(1, strKeys, Chr(0)) - 1)
    'strKeys = String(128, Chr(0))
    'Call GetTempPath(128, strKeys)
    'strKey = strKey & Left(strKeys, InStr(1, strKeys, Chr(0)) - 1)
    
    For lngPos = 1 To 255 Step 2
        strKey = strKey & ChrB(lngPos)
    Next
    
    For lngPos = 0 To 255 Step 2
        strKey = strKey & ChrB(lngPos)
    Next
    
    strKey = StrReverse(strKey) & strKey & StrReverse(strKey) & strKey & StrReverse(strKey) & strKey & StrReverse(strKey) & strKey
    
    For lngPos = 1 To lngLength
        strResult = strResult & Left(ConvertToBits(Mid(strKey, lngPos, 1)), 8)
    Next lngPos
    
    CreateKey = strResult

End Function


Private Function GetTrueFlagCount(ByVal strBits As String) As Long

    Dim lngPos      As Long
    Dim lngResult   As Long
    
    For lngPos = 1 To Len(strBits)
        lngResult = lngResult + CLng(Mid(strBits, lngPos, 1))
    Next lngPos
    
    GetTrueFlagCount = lngResult

End Function

Private Function MixTextAndKey(ByVal strText As String, ByVal strKey As String) As String

    Dim lngPos              As Long
    Dim lngBitCount         As Long
    Dim strNewText          As String
    Dim strResult           As String
    
    lngBitCount = GetTrueFlagCount(strText)
    
    strNewText = Mid(strText, lngBitCount + 1) & Left(strText, lngBitCount)
    
    For lngPos = 1 To Len(strText)
        If Mid(strNewText, lngPos, 1) = "1" And Mid(strKey, lngPos, 1) = "0" Then
            Mid(strKey, lngPos, 1) = "1"
        ElseIf Mid(strNewText, lngPos, 1) = "1" And Mid(strKey, lngPos, 1) = "1" Then
            Mid(strKey, lngPos, 1) = "0"
        End If
    Next lngPos
    
    For lngPos = 1 To Len(strKey) Step 8
        strResult = strResult & Right("00" & Hex(ConvertToByte(Mid(strKey, lngPos, 8))), 2)
    Next lngPos
    
    MixTextAndKey = strResult & Right("0000" & Hex(lngBitCount), 4)

End Function
 

Private Function GetBitCount(ByRef strHex As String) As Long
    
    Dim strBitCount         As String
    
    strBitCount = Right(strHex, 4)
    
    strHex = Left(strHex, Len(strHex) - 4)
    
    GetBitCount = ConvertHexToByte(Left(strBitCount, 2)) * 256 + ConvertHexToByte(Right(strBitCount, 2))

End Function

Private Function ConvertHexToBits(ByVal strHex As String) As String

    Dim lngPos      As Long
    Dim strResult   As String
    
    For lngPos = 1 To Len(strHex) Step 2
        strResult = strResult & ConvertToBits(Chr(ConvertHexToByte(Mid(strHex, lngPos, 2))))
    Next lngPos

    ConvertHexToBits = strResult
    
End Function

Private Function ConvertHexToByte(ByVal strHex As String) As Byte

    Const Hexs = "0123456789ABCDEF"
    
    ConvertHexToByte = (InStr(1, Hexs, Mid(strHex, 1, 1)) - 1) * 16 + (InStr(1, Hexs, Mid(strHex, 2, 1)) - 1)

End Function

Private Function RemixTextAndKey(ByVal strText As String, ByVal strKey As String, ByVal lngBitCount As Long) As String

    Dim lngPos              As Long
    Dim strNewText          As String
    Dim strResult           As String
    
    strNewText = String(Len(strText), "0")
    
    For lngPos = 1 To Len(strText)
        If Mid(strText, lngPos, 1) = "1" And Mid(strKey, lngPos, 1) = "0" Then
            Mid(strNewText, lngPos, 1) = "1"
        ElseIf Mid(strText, lngPos, 1) = "0" And Mid(strKey, lngPos, 1) = "1" Then
            Mid(strNewText, lngPos, 1) = "1"
        End If
    Next lngPos
    
    strNewText = Right(strNewText, lngBitCount) & Left(strNewText, Len(strNewText) - lngBitCount)
    
    For lngPos = 1 To Len(strKey) Step 8
        strResult = strResult & Chr(ConvertToByte(Mid(strNewText, lngPos, 8)))
    Next lngPos
    
    If Right(strResult, 2) <> "OK" Then
        RemixTextAndKey = ""
        Err.Raise 1, "Decrypt", "CRC Error"
    Else
        RemixTextAndKey = Left(strResult, Len(strResult) - 2)
    End If

End Function

