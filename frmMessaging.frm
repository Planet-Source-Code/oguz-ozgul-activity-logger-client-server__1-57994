VERSION 5.00
Begin VB.Form frmMessaging 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Messaging"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4695
   Begin VB.CheckBox chkSendToAll 
      Caption         =   "Send to all clients"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   2820
      Width           =   4035
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   435
      Left            =   1380
      TabIndex        =   1
      Top             =   3120
      Width           =   1755
   End
   Begin VB.TextBox txtMessage 
      Height          =   2295
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label lblReceiver 
      Height          =   255
      Left            =   1020
      TabIndex        =   3
      Top             =   2400
      Width           =   3075
   End
   Begin VB.Label Label1 
      Caption         =   "Receiver:"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   2400
      Width           =   855
   End
End
Attribute VB_Name = "frmMessaging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lngClientIndex      As Long

Private Sub cmdSend_Click()
    
    Dim lngIndex As Long
    
    If lngClientIndex <> 0 And chkSendToAll.Value = vbUnchecked Then
        If arrClients(lngClientIndex).IsConnected And arrClients(lngClientIndex).IsLoggedIn Then
            frmServer.sckClient(lngClientIndex).SendData objCrypto.Encrypt("AMSG" & Replace(txtMessage.Text, vbCrLf, "~")) & vbCrLf
        End If
    Else
        With frmServer.lvClients
            For lngIndex = 1 To .ListItems.Count
                If arrClients(Mid(.ListItems(lngIndex).Key, 2)).IsConnected And arrClients(Mid(.ListItems(lngIndex).Key, 2)).IsLoggedIn Then
                    frmServer.sckClient(Mid(.ListItems(lngIndex).Key, 2)).SendData objCrypto.Encrypt("AMSG" & Replace(txtMessage.Text, vbCrLf, "~")) & vbCrLf
                End If
            Next lngIndex
        End With
    End If

End Sub

Function SetClientIndex(ByVal lngIndex As Long) As Long
    lngClientIndex = lngIndex
    If arrClients(lngIndex).IsLoggedIn And arrClients(lngIndex).IsConnected Then
        lblReceiver.Caption = arrClients(lngIndex).Name
        chkSendToAll.Value = vbUnchecked
    Else
        lblReceiver.Caption = "All logged in clients"
        chkSendToAll.Value = vbChecked
    End If
End Function
