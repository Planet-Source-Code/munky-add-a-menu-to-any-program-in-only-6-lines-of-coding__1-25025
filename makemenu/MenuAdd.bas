Attribute VB_Name = "MenuAdd"
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Const MF_ENABLED = &H0&
Public Const MF_STRING = &H0&
Public Const MF_POPUP = &H10&
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Public Const WM_SYSCOMMAND = &H112
Public Const WM_MOVE = &HF012
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE
Public Declare Function CreatePopupMenu Lib "user32" () As Long

Public Sub CreateMenu(wnd As String, MnuCap As String, Mnuitem As String)
'Thanks for downloading my example, please vote!
'e-mail me at itzmunky@hotmail.com for more
Dim NewMenu As Long, FndWindow As Long, FndWindowMenu As Long
NewMenu = CreatePopupMenu
FndWindow = wnd ' Menu to Add to
'this is where you can add sub menus to the menu you create
'you can add as many as you like
Call AppendMenu(NewMenu, MF_ENABLED Or MF_STRING, 0, Mnuitem) 'You can also add as many items as you want
'this is the min menu
FndWindowMenu = GetMenu(FndWindow)
Call AppendMenu(FndWindowMenu, MF_POPUP, NewMenu, MnuCap)
End Sub
Public Sub WindowDrag(TheHwNd As String)
    Call ReleaseCapture
    Call SendMessage(TheHwNd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub
Public Sub windowontop(TheHwNd As String)
    Call SetWindowPos(TheHwNd, HWND_TOPMOST, 0&, 0&, 0&, 0&, Flags)

End Sub

