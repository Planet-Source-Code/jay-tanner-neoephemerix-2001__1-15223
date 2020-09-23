Attribute VB_Name = "Lock_Form_On_Top"
  Option Explicit

' =========================================================================

' These functions may be called to lock a form on top of all others
' and set it back to normal again.

  Public Declare Function SetWindowPos Lib "user32" _
 (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
  ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
  ByVal cy As Long, ByVal wFlags As Long) As Long

' USAGE:
' SetWindowPos [FORM_NAME].hwnd, -1,0,0,0,0,3 ' Lock on top
' SetWindowPos [FORM_NAME].hwnd, -2,0,0,0,0,3 ' Unlock on top
'
' Substitute a [FORM_NAME] without brackets (default = Form1)

  Public Sub LOCK_ON_TOP()

  SetWindowPos NeoEphemerix_2001_Interface.hwnd, -1, 0, 0, 0, 0, 3

  End Sub

  Public Sub UNLOCK_ON_TOP()

  SetWindowPos NeoEphemerix_2001_Interface.hwnd, -2, 0, 0, 0, 0, 3

  End Sub

' =========================================================================



