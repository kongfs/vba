'This is based on the job https://www.upwork.com/jobs/~01d96cf0c0d66917a7
Option Explicit
 
Dim excelFile
excelFile = "C:\github\vba-studio\sample\extract data from json response.xlsm"

Dim macroName
macroName = "DashboardSheet.Run" 

Dim scheduledTime
scheduledTime = CDate(FormatDateTime(Now, vbShortDate) & " 22:50:00")
 
RunMacro
 
Sub RunMacro()

	'Wait in loop until time arrives.
	While Now < scheduledTime
		WScript.Sleep 60 * 1000
	Wend
	
	scheduledTime = DateAdd("d", 1, scheduledTime)
 
	Dim app
	Set app = CreateObject("Excel.Application")
	app.Visible = True
	app.DisplayAlerts = False
	
	Dim book
	Set book = app.Workbooks.Open(excelFile, 0, True)
	 
	app.Run macroName

	book.Close
	Set book = Nothing

	app.Quit
	Set app = Nothing
 
	RunMacro 'Run the Macro again
	
End Sub
