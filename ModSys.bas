Attribute VB_Name = "ModSys"
'Tray Module
'Simple traymodule for Anim Icons
'by Scythe
'scythe@cablenet
'www.scythe-tools.de

Option Explicit

'Tray
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'Subclass
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'Allways on top
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_TOPMOST = -1

'Tray data & events
Private Type NOTIFYICONDATA
 cbSize As Long
 hwnd As Long
 uID As Long
 uFlags As Long
 uCallbackMessage As Long
 hIcon As Long
 szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const TRAY_CALLBACK = (&H7E9)
Private Const GWL_WNDPROC = (-4)

Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205


Dim OldWindowProc As Long
Dim TrayDat As NOTIFYICONDATA
Dim TrayForm As Form
Dim TrayMenu As Menu

Public Sub AddToTray(frm As Form, mnu As Menu)
 'Add form to tray
 
 Set TrayMenu = mnu
 Set TrayForm = frm

 'Subclass
 OldWindowProc = SetWindowLong(frm.hwnd, GWL_WNDPROC, AddressOf NewWindowProc)

 'Set the Tray Icon
 With TrayDat
 .uID = 0
 .hwnd = frm.hwnd
 .cbSize = Len(TrayDat)
 'We need a picture on the form to get the Icon from it
 .hIcon = frm.Pic.Picture
 .uFlags = NIF_ICON
 .uCallbackMessage = TRAY_CALLBACK
 .uFlags = .uFlags Or NIF_MESSAGE
 .cbSize = Len(TrayDat)
 End With
 
 'DO it
 Shell_NotifyIcon NIM_ADD, TrayDat

End Sub

'Subclass function
Public Function NewWindowProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
 
 'We pressed any Button ?
 If Msg = TRAY_CALLBACK Then
  If lParam = WM_RBUTTONUP Or lParam = WM_LBUTTONUP Then
   'Show the hidden Menu from form
   TrayForm.PopupMenu TrayMenu
   Exit Function
  End If
 End If
 
 'Go back to old routine
 NewWindowProc = CallWindowProc(OldWindowProc, hwnd, Msg, wParam, lParam)
End Function

'Delete from tray
Public Sub RemoveFromTray()
 'remove TrayIcon
 TrayDat.uFlags = 0
 Shell_NotifyIcon NIM_DELETE, TrayDat
 'End Subclassing
 SetWindowLong TrayForm.hwnd, GWL_WNDPROC, OldWindowProc
End Sub

'Show the new TrayIcon
Public Sub UpdateIcon()
 TrayDat.hIcon = TrayForm.Pic.Picture
 TrayDat.uFlags = NIF_ICON
 Shell_NotifyIcon NIM_MODIFY, TrayDat
End Sub
'Show the New Tooltip
Public Sub SetTrayTip(tip As String)
 TrayDat.szTip = tip & vbNullChar
 TrayDat.uFlags = NIF_TIP
 Shell_NotifyIcon NIM_MODIFY, TrayDat
End Sub
