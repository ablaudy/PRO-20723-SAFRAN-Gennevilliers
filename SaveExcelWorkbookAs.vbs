'Writes all values found to a new .csv file
'(adapted from write to excel script)
'

' Constants 
const xlCalculationAutomatic	= -4105	'Excel controls recalculation.
const xlCalculationManual		= -4135	'Calculation is done when the user requests it.


' Setup variables for arguments
Dim excelFilePath, exportFilePath, errorDialogMsgTitle, errorDialogMsg1, errorDialogMsg2


' Get Script Arguments
'---------------------------------------------------------
'Dim Arg
Set Arg = WScript.Arguments

' Full path of excel file
excelFilePath = Arg(0)

' Export file
exportFilePath = Arg(1)

' Arguments for error boxe(s)
errorDialogMsgTitle = Arg(2) 
errorDialogMsg1 = Arg(3) 
errorDialogMsg2 = Arg(4) 


' Had coded testing
'------------------------------------------------------------------------------------
'excelFilePath = "C:\Program Files\InnovMetric\PolyWorks MS 2022\report\ExcelTemplates\FirstArticleInspection\Template-AS9102BFAIR.xltx"
'exportFilePath = "C:\Users\TravisGarrison\Documents\PolyWorks\FAI Examples - 2022PR_Files\report\Test.xlsx"
'errorDialogMsgTitle = "Message"
'errorDialogMsg1 = "No installation of Microsoft Excel was found on your computer."
'errorDialogMsg2 = "No installation of Microsoft Excel was found on your computer."
'MsgBox (excelFilePath)
'MsgBox (excelTabName)
'MsgBox (exportFilePath)


' Launch Excel with error handling
'---------------------------------------------------------
Dim objReadXL, goodWB, objWorksheet

' Activate Error Handling
On Error Resume Next 

' Check if Excel is installed
Set objReadXL = CreateObject ( "Excel.Application" )
objReadXL.DisplayAlerts = False

If Err.Number <> 0 Then
    ' Excel is not installed. Abort
    Dim message
    message = MsgBox( errorDialogMsg1 & vbCrLf & errorDialogMsg2, 48, errorDialogMsgTitle )
    WScript.Quit
End If

' Verify if Excel is opened 
Set objReadXL = GetObject( , "Excel.Application" )

If Err.Number <> 0 Then
	' Excel not opened, open it
	Set objReadXL = CreateObject ( "Excel.Application" )
	Err.Clear
End IF

' Verify if desired workbook is opened
Set goodWB = objReadXL.Workbooks.Item(excelFilePath)

If Err.Number <> 0 Then
	' Workbook is not opened 
	Err.Clear
	
	' Opening file
	Set goodWB = objReadXL.Workbooks.Open(excelFilePath)
	
	If Err.Number <> 0 Then
		' Couldnt open the file
		msgbox("Could not open the Excel file")
		' Ending script
		WScript.Quit
	End if

End if

' Make sure Excel is shown (easier to troubleshoot if something goes wrong with saving or closing)
objReadXL.visible = True


' Could later add logic to also support saving as .csv
' Select correct sheet and export as .csv
'Set objWorksheet = goodWB.Worksheets("Mapping")
'objWorksheet.SaveAs  exportFilePath, 6


'Save the workbook as
goodWB.SaveAs exportFilePath


' Quit Excel workbook (will leave other workbooks open)
' For templates, force the "saved" property to true to avoid the save as prompt
goodWB.Saved = True 
' Supress alerts when closing
objReadXL.DisplayAlerts = False
goodWB.Close False

'Disable Error Handling
On Error GoTo 0
