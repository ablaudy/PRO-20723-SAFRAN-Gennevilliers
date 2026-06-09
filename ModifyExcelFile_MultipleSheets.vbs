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
Dim excelFilePath, excelTabName, txtFile, errorDialogMsgTitle, errorDialogMsg1, errorDialogMsg2
excelFilePath = Arg(0)
txtFile = Arg(1)
errorDialogMsgTitle = Arg(2) 
errorDialogMsg1 = Arg(3) 
errorDialogMsg2 = Arg(4) 




' Launch Excel with error handling
'---------------------------------------------------------
Dim objReadXL, goodWB, sheet

' Activate Error Handling
On Error Resume Next 

' Check if Excel is installed
Set objReadXL = CreateObject ( "Excel.Application" )

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
	
	' Start ReportLoop
	'(turn it off and on again hack for excel plugins run from vbs)
	objReadXL.AddIns("PolyWorksReportLoop Add-In").Installed = False
	objReadXL.AddIns("PolyWorksReportLoop Add-In").Installed = True
	' and again for the 2022 version of ReportLoop
	objReadXL.AddIns("PolyWorks|ReportLoop for Inspector Add-In").Installed = False
	objReadXL.AddIns("PolyWorks|ReportLoop for Inspector Add-In").Installed = True
	
End if

' Make sure Excel is shown
objReadXL.visible = True

' Select correct sheet  (move inside Do loop)
'Set sheet = goodWB.Sheets(excelTabName)

' Disable Excel "auto update" for performance
objReadXL.Calculation = xlCalculationManual

' Read the text file, Write the values
Dim cell, value, fso, file, line, intAsc1Chr, intAsc2Chr, openAsUnicode

Set fso = CreateObject("Scripting.FileSystemObject")

'Detect Unicode Files (UTF-16 only?)
Set file = fso.OpenTextFile(txtFile, 1, False)
intAsc1Chr = Asc(file.Read(1))
intAsc2Chr = Asc(file.Read(1))
Stream.Close
If intAsc1Chr = 255 And intAsc2Chr = 254 Then 
    openAsUnicode = True
Else
    openAsUnicode = False
End If


' Open either as Unicode or ASCII
Set file = fso.OpenTextFile(txtFile, 1, 0,OpenAsUnicode)

'sheet.Unprotect




' Read the text file, Write the values
'----------------------------------------------
Do Until file.AtEndOfStream
	' Read i-th line from controling text file
	line = file.ReadLine
	line = Split(line, "!_!")
	excelTabName = line(0)
	cell = line(1)
	value = line(2)
	
	'Select correct sheet  
	Set sheet = goodWB.Sheets(excelTabName)
	
	' Set i-th cell
	sheet.Range(cell).Value = value
 
Loop
' Close text file with cell values
file.Close





' Cleanup excel
'---------------------------------------------
' Reactivate Excel "auto update" 
objReadXL.Calculation = xlCalculationAutomatic

' Refresh on the excel workbook & save
goodWB.RefreshAll
goodWB.Save

'Disable Error Handling
On Error GoTo 0
