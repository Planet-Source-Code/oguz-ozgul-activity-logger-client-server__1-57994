VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDBCreateProgress 
   Caption         =   "Creating Application Database.. 0% Done"
   ClientHeight    =   810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   810
   ScaleWidth      =   7575
   Begin MSComctlLib.ProgressBar prgDB 
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   540
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblMessage 
      Caption         =   "Label1"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "frmDBCreateProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Height = 1215
    Width = 7695
End Sub
