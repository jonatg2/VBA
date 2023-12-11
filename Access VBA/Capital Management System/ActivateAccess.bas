Attribute VB_Name = "ActivateAccess"
Option Compare Database
Option Explicit

 
Public Declare PtrSafe Function SetForegroundWindow _
    Lib "user32" _
    (ByVal hwnd As LongPtr) _
    As LongPtr
 
Public Function ActivateAccessApp() As Boolean
    'Brings the DB to the front of all open windows
    Dim appTarget As Access.Application
 
    Set appTarget = GetObject(CurrentDb.Name)
    ActivateAccessApp = _
        Not (SetForegroundWindow(appTarget.hWndAccessApp) = 0)
    Set appTarget = Nothing
 
End Function
