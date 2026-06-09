'Write values to cells in Excel
' Argument 1: [in] [string] The Excel file path
' Argument 2: [in] [string] The name of the Excel sheet in the Excel file specified
' Argument 3: [in] [string] The file path of a text file of the following format
'		Cell1!_!Value1
'		Cell2!_!Value2
'		Cell3!_!Value3
'		etc...
'		
'		Example:
'		E33!_!This will be written in cell E33


' Constants 
const xlCalculationAutomatic	= -4105	'Excel controls recalculation.
const xlCalculationManual		= -4135	'Calculation is done when the user requests it.

' Get Script Arguments
Dim Arg
Set Arg = WScript.Arguments

' Set the arguments to the script variables
Dim excelFilePath, excelTabName, txtFile 
excelFilePath = Arg(0)
excelTabName = Arg(1)
txtFile = Arg(2)

' Control Excel 
Dim objReadXL, goodWB, sheet

' Activate Error Handling
On Error Resume Next 

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
		WScript.Exit
	End if
	
	' Start ReportLoop
	
	' objReadXL.AddIns("PolyWorksReportLoop Add-In").Installed = False
	' objReadXL.AddIns("PolyWorksReportLoop Add-In").Installed = True
	
End if

' Make sure Excel is shown
objReadXL.visible = True

' Select correct sheet
Set sheet = goodWB.Sheets(excelTabName)

' Disable Excel "auto update" for performance
objReadXL.Calculation = xlCalculationManual

' Read the text file, Write the values

Dim cell, value, fso, file, line

Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.OpenTextFile(txtFile)

'sheet.Unprotect

Do Until file.AtEndOfStream
  
  line = file.ReadLine
  line = Split(line, "!_!")
  cell = line(0)
  value = line(1)
  
  sheet.Range(cell).Value = value
 
Loop

file.Close

' Reactivate Excel "auto update" 
objReadXL.Calculation = xlCalculationAutomatic

'Disable Error Handling
On Error GoTo 0




