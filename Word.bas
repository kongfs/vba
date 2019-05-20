Attribute VB_Name = "Word"
Option Explicit
 
Public Function App(Optional ByVal Visible As Boolean = True)
    On Error Resume Next
    Dim wordApp
    Set wordApp = GetObject(, "Word.Application")
    If wordApp Is Nothing Then
        Set wordApp = CreateObject("Word.Application")
    End If
    wordApp.Visible = Visible
    Set App = wordApp
End Function

Public Function QuitApp(wordApp)
    If Not wordApp Is Nothing And wordApp.Documents.Count = 0 Then
        wordApp.Quit
    End If
End Function

