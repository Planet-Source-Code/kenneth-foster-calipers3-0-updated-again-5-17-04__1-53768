Attribute VB_Name = "modNoFocusRect"
Option Explicit
   Private Const GWL_WNDPROC         As Long = (-4)
   Private Const WM_SETFOCUS         As Long = &H7
   Private StandardButtonProc        As Long
   Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
   Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
   Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
   Global Const conHwndTopmost = -1
   Global Const conSwpNoActivate = &H10
   Global Const conSwpShowWindow = &H40


Private Function ButtonProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   
   'The procedure that gets all windows messages for the subclassed
   'button
   
   On Error Resume Next
   Select Case uMsg&
      Case WM_SETFOCUS 'The button is going to get the focus
         'Exit the procedure -> The message doesn´t reach the button
         Exit Function
   End Select
   'Call the standard Button Procedure
   ButtonProc = CallWindowProc(StandardButtonProc, hwnd&, uMsg&, wParam&, lParam&)
   
End Function

Public Sub NoFocusRect(Button As Object, vValue As Boolean)
   
   'Focus rect off
   
   If vValue Then
      'Save the adress of the standard button procedure
      StandardButtonProc = GetWindowLong(Button.hwnd, GWL_WNDPROC)
      'Subclass the button to control its Windows Messages
      SetWindowLong Button.hwnd, GWL_WNDPROC, AddressOf ButtonProc
   Else 'Focus rect on
      'Remove the subclassing from the button
      SetWindowLong Button.hwnd, GWL_WNDPROC, StandardButtonProc
   End If
   
End Sub
