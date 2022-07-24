Imports System.IO
Imports System.Text

Sub Main()

'----------------------------Your custom message

Result = MessageBox.Show("Aby Macro poprawnie zadziałało:" & vbNewLine & vbNewLine &
"- rysunki muszą się znajodwać w lokalizacji o końcówce _RYSUNKI\WYKONAWCZE\" & vbNewLine &
"- rysunki muszą mieć takie same nazwy jak złożenia" & vbNewLine &
"- macro trzeba odpalić z poziomu głównego złożenia" & vbNewLine &
"- macro generuje PDFy zgodnie z kolejnościa jaka jest w BOMie Structural","assemblerPDF",MessageBoxButtons.OKCancel,MessageBoxIcon.Information)

'----------------------------

If Result =1

Dim openDoc As AssemblyDocument
openDoc = ThisDoc.Document
Dim oDoc As Document = ThisApplication.ActiveDocument
Dim oDocNameMain As String = oDoc.FullFileName
Dim mainPath As String = Split(oDocNameMain, oDoc.DisplayName)(0)

ThisApplication.ActiveDocument.ComponentDefinition.BOM.StructuredViewEnabled = True
ThisApplication.ActiveDocument.ComponentDefinition.BOM.PartsOnlyViewEnabled = True
Dim count As Integer = openDoc.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Count

Dim counter As Integer = 1


creatingPDF(counter, mainPath)

firstLoop(openDoc, oDoc, count, counter, mainPath)

Else
End If


Exit Sub
End Sub


'-------------------------------------------------------------------------- first loop -----------------------------------------------------------------


Function firstLoop (openDoc As AssemblyDocument, oDoc As Document, count As Integer, counter As Integer, mainPath As String)




For item As Integer = 1 To count 
	On Error Resume Next

	counter= counter + 10
	
	

	fullName=openDoc.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(item).ComponentDefinitions.Item(1).Document.FullDocumentName
	assemblyType = openDoc.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(item).ComponentDefinitions.Item(1).Type
	bomStructure=openDoc.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(item).BOMStructure
	Dim oPart As AssemblyDocument
	
		
	if assemblyType <>	kAssemblyComponentDefinitionObject Or bomStructure <> kNormalBOMStructure  Then 
			
		'ThisApplication.ActiveDocument.Close(True)
		
		counter = counter - 10
		
		Continue For
		
	Else if assemblyType = kAssemblyComponentDefinitionObject
		oPart = ThisApplication.Documents.Open(fullName, True)

		ThisApplication.ActiveDocument.ComponentDefinition.BOM.StructuredViewEnabled = True
		ThisApplication.ActiveDocument.ComponentDefinition.BOM.PartsOnlyViewEnabled = True
		newCount = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Count
			
		creatingPDF(counter, mainPath)
		counter = secondLoop(newCount, counter, mainPath)
			
	End if
Next


return counter

End Function


'-------------------------------------------------------------------------- Second loop -----------------------------------------------------------------


Function secondLoop (newCount, counter, mainPath)



For newItem As Integer = 1 To newCount 
			
	On Error Resume Next
	
	
	
	
	
	counter = counter + 10
			
			
	newFullName=ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem).ComponentDefinitions.Item(1).Document.FullDocumentName
	newBomStructure= ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newitem).BOMStructure
	newAssemblyType = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem).ComponentDefinitions.Item(1).Type
	Dim NewoPart As AssemblyDocument
	Dim closePart As PartDocument
	
			
			
	if newAssemblyType = kAssemblyComponentDefinitionObject
		NewoPart = ThisApplication.Documents.Open(newFullName, True)
		
		ThisApplication.ActiveDocument.ComponentDefinition.BOM.StructuredViewEnabled = True
		ThisApplication.ActiveDocument.ComponentDefinition.BOM.PartsOnlyViewEnabled = True
		newCount1 = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Count
		
		creatingPDF(counter, mainPath)		
		counter = secondLoop(newCount1, counter, mainPath)
				
	Else if newAssemblyType <>	kAssemblyComponentDefinitionObject Or newAssemblyType <> kNormalBOMStructure Then 
		
		
		counter = counter - 10
		'ThisApplication.ActiveDocument.Close(True)
		
		
		Continue For
				
			End if
		Next
		
		
	

			ThisApplication.ActiveDocument.Close(True)

		return counter
		
End Function





'-------------------------------------------------------------------------- PDF creation -----------------------------------------------------------------


Function creatingPDF(counter, mainPath)

Dim pdfCounter As Integer= counter 
Dim oDoc As Document
oDoc = ThisApplication.ActiveDocument

	
	Dim oDocName As String = oDoc.FullFileName
	Dim oDocJustName As String = oDoc.DisplayName
	
	
	Dim sFileName As String = Split(oDocName, oDocJustName)(0)
	
	
	Dim  displayNameCut As String = Split(oDocJustName, ".iam")(0)
	
	Dim sDrawingName As String = sFileName & "_RYSUNKI\WYKONAWCZE\" & displayNameCut & ".idw" '<------ drawing file path


	Dim directionTry As String = sFileName & "_RYSUNKI\WYKONAWCZE"
	
	

	Try 
		If Not System.IO.Directory.Exists(directionTry) Then
	
		MessageBox.Show("Twój projekt nie posiada właściwej ścieżki do rysunków zgodnej z instrukcją PB 19 I 1"&
		vbNewLine &vbNewLine &"Twoja ścieżka powinna wyglądać tak:" &
		vbNewLine &vbNewLine &
		sFileName & "_RYSUNKI\WYKONAWCZE\","assemblerPDF",MessageBoxButtons.OK,MessageBoxIcon.Exclamation)
		
		End If
	Catch
		
	
	End Try
	
Try	
	If Not System.IO.File.Exists(sDrawingName) Then
	
		
		errorLog (sDrawingName, mainPath, pdfCounter)
		
	End If
Catch


End Try	
ThisApplication.Documents.Open(sDrawingName, True)


Dim drwoDoc As Document
drwoDoc = ThisApplication.ActiveDocument

MakePDFFromDoc(drwoDoc, sFileName, pdfCounter, mainPath)

ThisApplication.ActiveDocument.Close

End Function

Function MakePDFFromDoc(docDrw, sFileName, counter, mainPath)
 
 
	
 

	Dim docJustName As String = docDrw.DisplayName
	Dim docNameCut As String = Split(docJustName, ".idw")(0)

	Dim newPDFname As String = counter & "." & docNameCut
	
	
	
 oPDFAddIn = ThisApplication.ApplicationAddIns.ItemById _
  ("{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}")
  oContext = ThisApplication.TransientObjects.CreateTranslationContext
  oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism
  oOptions = ThisApplication.TransientObjects.CreateNameValueMap
  oDataMedium = ThisApplication.TransientObjects.CreateDataMedium
  
  


  oOptions.Value("All_Color_AS_Black") = 0
  oOptions.Value("Remove_Line_Weights") = 0
  oOptions.Value("Vector_Resolution") = 400
  oOptions.Value("Sheet_Range") = Inventor.PrintRangeEnum.kPrintAllSheets
 


	if  sFileName = mainPath Then

  oDataMedium.FileName =  sFileName & "_RYSUNKI\WYKONAWCZE\" & newPDFname  & ".pdf"'<------ pdf file path


  oPDFAddIn.SaveCopyAs(docDrw, oContext, oOptions, oDataMedium)
  
  Else
  
  oDataMedium.FileName =  mainPath & "_RYSUNKI\WYKONAWCZE\" & newPDFname  & ".pdf"'<------ pdf file path


	oPDFAddIn.SaveCopyAs(docDrw, oContext, oOptions, oDataMedium)
  
  End if
  
End Function

'-------------------------------------------------------------------------- Error log -----------------------------------------------------------------

Function errorLog (sDrawingName, mainPath, pdfCounter)

'Dim myDate As String = Now().ToString("yyyy-MM-dd HH.m.ss")
'myDate = myDate.Replace(":","")  

Dim path As String = mainPath & "_RYSUNKI\WYKONAWCZE\" &"drwLog.txt" '<------ error log file path

Dim file As System.IO.StreamWriter
file = My.Computer.FileSystem.OpenTextFileWriter(path, True)
	
file.WriteLine(pdfCounter & "." & sDrawingName)

file.Close()
End Function








