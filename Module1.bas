Attribute VB_Name = "Module1"
'This function releases a captured mouse click
Public Declare Function ReleaseCapture Lib "user32" () As Long

'This function sends a message to the queue- used to drag the ruler
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'Variables used by SendMessage to move the ruler
Global Const HTCAPTION = 2
Global Const WM_NCLBUTTONDOWN = &HA1

'This function allows the window to be placed on top of all others
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'Variables used by SetWindowPos
Global Const SWP_NOMOVE = &H2
Global Const SWP_NOSIZE = &H1
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2

'This function returns the system position of the mouse pointer
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

'Variables used by GetCursorPos
Public Type POINTAPI
    X As Long
    Y As Long
End Type
Global MousePos As POINTAPI

'*******************************************************************************
' AlwaysOnTop (SUB)
'
' PARAMETERS:
' (In/Out) - FrmID - Form    - Form that will be positioned
' (In/Out) - OnTop - Integer - Should the form be always on top?
'
' DESCRIPTION:
' Set the form to be always on top (if true) or to zorder normally (if false)
'*******************************************************************************
Sub AlwaysOnTop(FrmID As Form, OnTop As Integer)
    Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
    If OnTop = -1 Then
        OnTop = SetWindowPos(FrmID.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    Else
        OnTop = SetWindowPos(FrmID.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
    End If
End Sub






