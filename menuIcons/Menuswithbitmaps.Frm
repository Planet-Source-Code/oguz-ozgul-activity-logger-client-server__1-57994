VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   2280
   ClientLeft      =   4515
   ClientTop       =   2430
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2280
   ScaleWidth      =   6390
   Begin VB.CommandButton Command3 
      Caption         =   "Show SM height"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Show SM width"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add bitmaps to Menus"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Regards, vijaybond"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "The ""Test"" item shows what happens to a colour item when using this approach. See the images in the image list."
      Height          =   1215
      Left            =   3840
      TabIndex        =   4
      Top             =   960
      Width           =   2415
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Click anywhere on form with right mouse button to view popup version of File menu"
      Height          =   855
      Left            =   3840
      TabIndex        =   3
      Top             =   0
      Width           =   2415
      WordWrap        =   -1  'True
   End
   Begin ComctlLib.ImageList img 
      Left            =   3000
      Top             =   1020
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   16777215
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menuswithbitmaps.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menuswithbitmaps.frx":025A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menuswithbitmaps.frx":04B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menuswithbitmaps.frx":070E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menuswithbitmaps.frx":0968
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menuswithbitmaps.frx":0BC2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuMCR 
      Caption         =   "File"
      Begin VB.Menu mnuOne 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuTwo 
         Caption         =   "&Design"
      End
      Begin VB.Menu mnuS1 
         Caption         =   "D&elete"
      End
      Begin VB.Menu mnuMoeder 
         Caption         =   "&Properties"
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTest 
         Caption         =   "Test"
      End
   End
   Begin VB.Menu mnuFerdi 
      Caption         =   "Other"
      Begin VB.Menu mnuF1 
         Caption         =   "One"
      End
      Begin VB.Menu mnuF2 
         Caption         =   "Two"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As String) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long

Private Declare Function GetMenuCheckMarkDimensions Lib "user32" () As Long

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Sub Command1_Click()
  Dim i%
  Dim hMenu, hSubMenu, menuID, x
  hMenu = GetMenu(hwnd)
  hSubMenu = GetSubMenu(hMenu, 0) '1 for "Other" menu etcetera
  For i = 1 To 4
    menuID = GetMenuItemID(hSubMenu, i - 1)
    x = SetMenuItemBitmaps(hMenu, menuID, &H4, img.ListImages(i).Picture, img.ListImages(i).Picture)
  Next
  menuID = GetMenuItemID(hSubMenu, 5)
  x = SetMenuItemBitmaps(hMenu, menuID, 0, img.ListImages(5).Picture, 0&)

End Sub


Private Sub Command2_Click()
  MsgBox CStr(GetSystemMetrics(71))
End Sub

Private Sub Command3_Click()
  MsgBox CStr(GetSystemMetrics(72))
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 2 Then
    PopupMenu mnuMCR, 2, x
  End If
End Sub



