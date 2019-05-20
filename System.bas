Attribute VB_Name = "System"
Option Explicit

Public Function Desktop()
    Desktop = CreateObject("WScript.Shell").specialfolders("Desktop")
End Function

