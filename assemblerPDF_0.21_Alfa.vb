Sub Main()

Dim openDoc As AssemblyDocument
openDoc = ThisDoc.Document
Dim oDoc As Document


ThisApplication.ActiveDocument.ComponentDefinition.BOM.StructuredViewEnabled = True
ThisApplication.ActiveDocument.ComponentDefinition.BOM.PartsOnlyViewEnabled = True
Dim count As Integer = openDoc.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Count

Dim counter As Integer = 1

'MsgBox(counter & "- Glowne zlozenie")

creatingPDF(counter)

firstLoop(openDoc, oDoc, count, counter)

End Sub


'-------------------------------------------------------------------------- pierwsza petla -----------------------------------------------------------------


Function firstLoop (openDoc As AssemblyDocument, oDoc As Document, count As Integer, counter As Integer)

For item As Integer = 1 To count 
	On Error Resume Next

	counter= counter + 1
	'pathMap = openDoc.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(item).ComponentDefinitions.Item(1).Document.DisplayName

	'MsgBox(counter &  "- Zlozenie I poz.")
		
		
	fullName=openDoc.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(item).ComponentDefinitions.Item(1).Document.FullDocumentName
	assemblyType = openDoc.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(item).ComponentDefinitions.Item(1).Type
	Dim oPart As AssemblyDocument
		
	if assemblyType <>	kAssemblyComponentDefinitionObject Then 
			
		'ThisApplication.ActiveDocument.Close(True)
		Exit For
		
	Else if assemblyType = kAssemblyComponentDefinitionObject
		oPart = ThisApplication.Documents.Open(fullName, True)

		ThisApplication.ActiveDocument.ComponentDefinition.BOM.StructuredViewEnabled = True
		ThisApplication.ActiveDocument.ComponentDefinition.BOM.PartsOnlyViewEnabled = True
		newCount = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Count
			
		creatingPDF(counter)
		counter = secondLoop(newCount, counter)
			
	End if
Next

return counter

End Function


'-------------------------------------------------------------------------- druga petla -----------------------------------------------------------------


Function secondLoop (newCount, counter)

For newItem As Integer = 1 To newCount 
			
	On Error Resume Next
	
	counter= counter + 1
			
	'pathMap0 = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem).ComponentDefinitions.Item(1).Document.DisplayName

	'MsgBox(counter & "- Zlozenie II poz.")
			
	newFullName=ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem).ComponentDefinitions.Item(1).Document.FullDocumentName
	newAssemblyType = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem).ComponentDefinitions.Item(1).Type
	Dim NewoPart As AssemblyDocument
	Dim closePart As PartDocument
			
			
	if newAssemblyType = kAssemblyComponentDefinitionObject
		NewoPart = ThisApplication.Documents.Open(newFullName, True)
		
		ThisApplication.ActiveDocument.ComponentDefinition.BOM.StructuredViewEnabled = True
		ThisApplication.ActiveDocument.ComponentDefinition.BOM.PartsOnlyViewEnabled = True
		newCount1 = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Count
		
		creatingPDF(counter)		
		counter = thirdLoop(newCount1, counter)
				
	Else if newAssemblyType <>	kAssemblyComponentDefinitionObject Then 
		ThisApplication.ActiveDocument.Close(True)
		Exit For
				
			End if
		Next
		
		return counter
		
End Function



'-------------------------------------------------------------------------- trzecia petla -----------------------------------------------------------------


Function thirdLoop (newCount1, counter)
For newItem1 As Integer = 1 To newCount1 
				
						
On Error Resume Next
					
	counter= counter + 1	
	
	'pathMap1 = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem).ComponentDefinitions.Item(1).Document.DisplayName

	'MsgBox(counter & "- Zlozenie III poz.")
					
	newFullName1=ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem1).ComponentDefinitions.Item(1).Document.FullDocumentName
	newAssemblyType1 = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem1).ComponentDefinitions.Item(1).Type
	Dim NewoPart1 As AssemblyDocument
	Dim closePart1 As PartDocument
					
	if newAssemblyType1 = kAssemblyComponentDefinitionObject
		NewoPart1 = ThisApplication.Documents.Open(newFullName1, True)
			
		ThisApplication.ActiveDocument.ComponentDefinition.BOM.StructuredViewEnabled = True
		ThisApplication.ActiveDocument.ComponentDefinition.BOM.PartsOnlyViewEnabled = True
		newCount2 = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Count
		
		creatingPDF(counter)
		counter =fourthLoop(newCount2, counter)
		
	Else if newAssemblyType1 <>	kAssemblyComponentDefinitionObject Then 
		ThisApplication.ActiveDocument.Close(True)
		Exit For
						
	End if
Next

return counter

End Function


'-------------------------------------------------------------------------- czwarta petla -----------------------------------------------------------------

Function fourthLoop (newCount2, counter)
For newItem2 As Integer = 1 To newCount2 
				
						
On Error Resume Next
					
	counter= counter + 1	
	
	'pathMap1 = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem).ComponentDefinitions.Item(1).Document.DisplayName

	'MsgBox(counter & "- Zlozenie III poz.")
					
	newFullName2=ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem1).ComponentDefinitions.Item(1).Document.FullDocumentName
	newAssemblyType2 = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem1).ComponentDefinitions.Item(1).Type
	Dim NewoPart2 As AssemblyDocument
	Dim closePart2 As PartDocument
					
	if newAssemblyType2 = kAssemblyComponentDefinitionObject
		NewoPart2 = ThisApplication.Documents.Open(newFullName2, True)
			
		ThisApplication.ActiveDocument.ComponentDefinition.BOM.StructuredViewEnabled = True
		ThisApplication.ActiveDocument.ComponentDefinition.BOM.PartsOnlyViewEnabled = True
		newCount3 = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Count
		
		creatingPDF(counter)
		counter = fifthLoop(newCount3, counter)
		
	Else if newAssemblyType2 <>	kAssemblyComponentDefinitionObject Then 
		ThisApplication.ActiveDocument.Close(True)
		Exit For
						
	End if
Next

return counter

End Function


'-------------------------------------------------------------------------- piąta petla -----------------------------------------------------------------

Function fifthLoop(newCount3, counter)

For newItem3 As Integer = 1 To newCount3 
							
	On Error Resume Next
	counter= counter + 1					
							
	'pathMap2 = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem).ComponentDefinitions.Item(1).Document.DisplayName

	'MsgBox(counter & "- Zlozenie IV poz.")
	
	newFullName3=ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem2).ComponentDefinitions.Item(1).Document.FullDocumentName
	newAssemblyType3 = ThisApplication.ActiveDocument.ComponentDefinition.BOM.BOMViews.Item(2).BOMRows.Item(newItem2).ComponentDefinitions.Item(1).Type
	Dim NewoPart3 As AssemblyDocument
	Dim closePart3 As PartDocument
							
					
	if newAssemblyType3 = kAssemblyComponentDefinitionObject
		NewoPart3 = ThisApplication.Documents.Open(newFullName3, True)
		creatingPDF(counter)
						
	Else if newAssemblyType3 <>	kAssemblyComponentDefinitionObject Then 
						
		ThisApplication.ActiveDocument.Close(True)
		Exit For
								
	End if
Next

return counter

End Function





'-------------------------------------------------------------------------- funkcja tworząca PDF -----------------------------------------------------------------


Function creatingPDF(counter)

Dim pdfCounter As Integer= counter 
Dim oDoc As Document
oDoc = ThisApplication.ActiveDocument

	
	Dim oDocName As String = oDoc.FullFileName
	Dim oDocJustName As String = oDoc.DisplayName
	
	
	Dim sFileName As String = Split(oDocName, oDocJustName)(0)
	
	Dim  displayNameCut As String = Split(oDocJustName, ".iam")(0)
	
	Dim sDrawingName As String = sFileName & "_Rysunki\Wykonawcze\" & displayNameCut & ".idw"
	

ThisApplication.Documents.Open(sDrawingName, True)

Dim drwoDoc As Document
drwoDoc = ThisApplication.ActiveDocument

MakePDFFromDoc(drwoDoc, sFileName, pdfCounter)

ThisApplication.ActiveDocument.Close

End Function

Function MakePDFFromDoc(docDrw, sFileName, counter)
 

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
 

  oDataMedium.FileName =  sFileName & "_Rysunki\Wykonawcze\" & newPDFname  & ".pdf"

  oPDFAddIn.SaveCopyAs(docDrw, oContext, oOptions, oDataMedium)
  
  
End Function



