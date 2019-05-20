Attribute VB_Name = "Clipboard"
Option Explicit

Public Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function EmptyClipboard Lib "user32" () As Long
Public Declare Function CloseClipboard Lib "user32" () As Long

Public Function Clear()
    OpenClipboard (0&)
    EmptyClipboard
    CloseClipboard
End Function
 
