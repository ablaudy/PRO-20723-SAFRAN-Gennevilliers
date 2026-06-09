Sub ExporterFeuillePDFDepuisFichierFerme()

	' Get Script Arguments
	Dim Arg
	Set Arg = WScript.Arguments
    
    ' Set the arguments to the script variables
	Dim excelFilePath, CheminPDF 
	excelFilePath = Arg(0)
	CheminPDF = Arg(1)

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

    ' Définir la feuille à exporter (ici, nous prenons la première feuille)
    ' Vous pouvez spécifier une feuille particulière par son nom, par exemple "Feuille1"
    Set Feuille = Classeur.Sheets(1)

    ' Exporter la feuille active (Feuille) au format PDF
    Feuille.ExportAsFixedFormat Type:=PDF, Filename:=CheminPDF, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False

    ' Fermer le fichier Excel sans enregistrer les modifications
    Classeur.Close SaveChanges:=False

    ' Quitter l'application Excel
    ApplicationExcel.Quit

    ' Libérer la mémoire
    Set Feuille = Nothing
    
    ' Message de confirmation
    MsgBox "La feuille a été exportée en PDF avec succès!", vbInformation
End Sub
