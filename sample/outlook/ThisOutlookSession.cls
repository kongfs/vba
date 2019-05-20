VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'Please refer to https://www.upwork.com/o/jobs/browse/saved/details/~012f8f240d0200cb97/?sort=recency

Const wordName As String = "Outlook.docx"

Sub Test()
 
    Dim wordFile As String
    wordFile = System.Desktop & "\" & wordName
    
    Dim wordApp
    Set wordApp = Word.App
 
    Dim wordDoc
    Set wordDoc = wordApp.Documents.Open(wordFile)
    
    Dim selectedItems As Selection
    Set selectedItems = Application.ActiveExplorer.Selection
  
    Dim selectedItem
    For Each selectedItem In selectedItems
        If selectedItem.Class = olMail Then
            selectedItem.GetInspector().WordEditor.Range.FormattedText.Copy
            Dim oRange
            Set oRange = wordDoc.Content
            oRange.Collapse Direction:=0  'Word.WdCollapseDirection.wdCollapseEnd
            oRange.Paste
        End If
    Next
    
    wordDoc.Save
    wordDoc.Close
    Word.QuitApp wordApp
     
    Clipboard.Clear

End Sub
