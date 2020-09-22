Attribute VB_Name = "modSystemTray"
' Tim Harpur - Logicon Enterprises

' These three routines allow a VB program to place a notification icon in the system tray
' - this is usually the lower right of the desktop next to the clock

' The displayed icon is taken from the passed Form and displays the text 'Tip' when the
' user passes the mouse over it.

' When a user presses either left, centre, or right mouse buttons over the icon the Form's
' mouse down event is triggered (you could actually change it so that it used a different
' callback event on the form if you needed the Form's mouse down event for actaul mouse
' down events). When the Form's mouse down event is triggered the first thing to do is
' modify the contents of the passed X value as this is actually the passed mouse button, but
' Visual Basic thinks it is still the Form's mouse down event and so modifies this value by
' Screen.TwipsPerPixelX - this is not a problem though as all you have to do is divide the value
' X by Screen.TwipsPerPixelX to get the button state (one of the STI_... constants below)

' Usually you would display a popup menu for the right click and the trigger the default action
' with a double left click, as is standard with Windows based programs

Option Explicit

Private Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 64
End Type

Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

Private Const NIM_ADD = 0
Private Const NIM_MODIFY = 1
Private Const NIM_DELETE = 2
Private Const NIF_MESSAGE = 1
Private Const NIF_ICON = 2
Private Const NIF_TIP = 4

Private Const STI_CALLBACKEVENT = &H203 'form's MouseDown event

' the are the possible returns that need to be addressed by the calling Form's callback event, which
' is currently set to the calling Form's mouse down event.
Public Const STI_LBUTTONDOWN = &H201
Public Const STI_LBUTTONUP = &H202
Public Const STI_LBUTTONDBCLK = &H203
Public Const STI_RBUTTONDOWN = &H204
Public Const STI_RBUTTONUP = &H205
Public Const STI_RBUTTONDBCLK = &H206

' call this routine to initially create the system tray icon for the program
Public Sub CreateSystemTrayIcon(parentForm As Form, Tip As String)
  Dim notIcon As NOTIFYICONDATA
  
  With notIcon
    .cbSize = Len(notIcon)
    .hwnd = parentForm.hwnd
    .uID = vbNull
    .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
    .uCallbackMessage = STI_CALLBACKEVENT
    .hIcon = parentForm.Icon
    .szTip = Tip & vbNullChar
  End With
  
  Shell_NotifyIconA NIM_ADD, notIcon
End Sub

' call this routine to modify the displayed icon for the program's system tray icon
Public Sub ModifySystemTrayIcon(parentForm As Form, Tip As String)
  Dim notIcon As NOTIFYICONDATA
  
  With notIcon
    .cbSize = Len(notIcon)
    .hwnd = parentForm.hwnd
    .uID = vbNull
    .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
    .uCallbackMessage = STI_CALLBACKEVENT
    .hIcon = parentForm.Icon
    .szTip = Tip & vbNullChar
  End With
  
  Shell_NotifyIconA NIM_MODIFY, notIcon
End Sub

' call this routine to remove the program's system tray icon (if you don't do this step before
' closing the program the icon will still sit there, but won't cause any problems)
Public Sub DeleteSystemTrayIcon(parentForm As Form)
  Dim notIcon As NOTIFYICONDATA
  
  With notIcon
    .cbSize = Len(notIcon)
    .hwnd = parentForm.hwnd
    .uID = vbNull
    .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
    .uCallbackMessage = vbNull
    .hIcon = vbNull
    .szTip = "" & vbNullChar
  End With
  
  Shell_NotifyIconA NIM_DELETE, notIcon
End Sub
