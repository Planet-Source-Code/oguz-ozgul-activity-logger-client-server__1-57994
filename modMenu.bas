Attribute VB_Name = "modMenu"
Option Explicit

Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long

Public Function SetMenuIcons()
  
  Dim i         As Long
  Dim hMenu     As Long
  Dim hSubMenu  As Long
  Dim menuID    As Long
  Dim x         As Long
  
  hMenu = GetMenu(mdiMain.hwnd)
  hSubMenu = GetSubMenu(hMenu, 0) '1 for "Other" menu etcetera
  menuID = GetMenuItemID(hSubMenu, 0)
  x = SetMenuItemBitmaps(hMenu, menuID, &H4, mdiMain.iList.ListImages(3).Picture, mdiMain.iList.ListImages(3).Picture)
  hSubMenu = GetSubMenu(hMenu, 1) '1 for "Other" menu etcetera
  menuID = GetMenuItemID(hSubMenu, 0)
  x = SetMenuItemBitmaps(hMenu, menuID, &H4, mdiMain.iList.ListImages(1).Picture, mdiMain.iList.ListImages(1).Picture)
  menuID = GetMenuItemID(hSubMenu, 1)
  x = SetMenuItemBitmaps(hMenu, menuID, &H4, mdiMain.iList.ListImages(2).Picture, mdiMain.iList.ListImages(2).Picture)
End Function
